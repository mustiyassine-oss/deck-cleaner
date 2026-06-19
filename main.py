"""Template Cleaner API — upload a PPTX → get a layout-cleaned PPTX.

This service is cleaning-only: it runs visual QC on an uploaded deck, applies
deterministic geometric fixes, and (optionally) an LLM geometric fallback pass
for slides that still fail. There is no deck-generation pipeline.
"""
from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import tempfile
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(name)-10s  %(message)s",
    datefmt="%H:%M:%S",
)

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic_settings import BaseSettings

from qc import run_visual_qc
from pdf_geometry import run_geometry_qc, offslide_issues, keepout_issues
from template_cleaner import clean_template_from_qc
from llm_fix_fallback import apply_llm_fallback_fixes
from verify import (
    index_slides,
    build_report_from_index,
    changed_slides_from_clean_report,
    apply_verification_gate,
)

log = logging.getLogger("cleaner")

# Issue types that repositioning shapes cannot resolve (tiny fonts, minor
# alignment). Slides failing ONLY for these are still reported, but they do not
# trigger further repair passes or the LLM geometric fallback.
_SOFT_ISSUE_TYPES = {"overflow", "alignment"}


def _augment_with_geometry_signals(
    report: dict, pptx_bytes: bytes, slide_numbers: set[int] | None = None
) -> dict:
    """
    Merge the always-on deterministic geometry signals into a QC report:
    off-slide/clipping AND side-panel/footer-band intrusions.

    The vision engine is structurally blind to both — its screenshots clip
    off-slide content and it routinely misses a card spilling onto a dark sidebar
    or into the footer bar. We add the matching issue to the slide entry and flip
    it to ``fail`` so the existing repair loop (reflow / panel + footer fixers)
    acts on it. Dedupes by issue type, so the deterministic engine — which already
    reports these — is unaffected. Degrades to a no-op on any error / skip.
    """
    if not isinstance(report, dict):
        return report
    slides = report.get("slides")
    if not isinstance(slides, list) or not slides:
        return report

    combined: dict[int, list[dict]] = {}
    for source in (offslide_issues, keepout_issues):
        try:
            found = source(pptx_bytes)
        except Exception as e:
            log.warning("[Geometry] signal %s unavailable: %s", getattr(source, "__name__", "?"), e)
            continue
        for n, issues in (found or {}).items():
            combined.setdefault(n, []).extend(issues)
    if not combined:
        return report
    off = combined

    target = {int(n) for n in (slide_numbers or [])}
    by_no = {
        int(s.get("slide_number", -1)): s for s in slides if isinstance(s, dict)
    }
    changed = False
    for slide_no, issues in off.items():
        if target and slide_no not in target:
            continue
        entry = by_no.get(slide_no)
        if entry is None:
            continue
        existing = entry.get("issues") if isinstance(entry.get("issues"), list) else []
        existing_types = {
            str(i.get("type", "")).lower() for i in existing if isinstance(i, dict)
        }
        added = False
        for iss in issues:
            if str(iss.get("type", "")).lower() in existing_types:
                continue
            existing.append(iss)
            added = True
        if added:
            entry["issues"] = existing
            entry["status"] = "fail"
            entry["severity"] = "high"
            changed = True

    if changed:
        fails = sum(
            1 for s in slides
            if isinstance(s, dict) and str(s.get("status", "")).lower() == "fail"
        )
        report["failed_slide_count"] = fails
        report["overall_status"] = "fail" if fails else report.get("overall_status", "pass")
    return report


def _user_friendly_error(exc: BaseException) -> str:
    """Map API/server exceptions to clear, actionable messages."""
    msg = str(exc).lower()
    name = type(exc).__name__.lower()

    if "401" in msg or "authentication" in msg or ("invalid" in msg and "api" in msg):
        return "Invalid or expired API key. Please check OPENAI_API_KEY in .env and try again."
    if "auth" in name or "authentication" in name:
        return "Invalid or expired API key. Please check OPENAI_API_KEY in .env and try again."

    if "429" in msg or ("rate" in msg and "limit" in msg):
        return "Rate limit exceeded. Please wait a minute and try again."
    if "rate" in name and "limit" in name:
        return "Rate limit exceeded. Please wait a minute and try again."

    if "timeout" in msg or "timed out" in msg or "timeout" in name:
        return "Request timed out. Try again or use a smaller template."

    if "connection" in msg or "connection refused" in msg or "connection" in name:
        return "Cannot connect to the AI service. Check your internet connection and try again."

    if "500" in msg or "503" in msg or "unavailable" in msg:
        return "The AI service is temporarily unavailable. Please try again in a few minutes."

    if "400" in msg or "bad request" in msg:
        return "Invalid request. Please check your PPTX file and try again."

    s = str(exc)
    return s[:300] + "…" if len(s) > 300 else s


BASE_DIR = Path(__file__).resolve().parent


class Settings(BaseSettings):
    OPENAI_API_KEY: str | None = None
    OPENAI_BASE_URL: str | None = None

    # Vision QC model (reads slide screenshots and flags layout failures).
    OPENAI_VISION_QC_MODEL: str = "gpt-4.1"
    ENABLE_VISUAL_QC_REPAIR: bool = True

    # LLM geometric fallback for slides still failing after deterministic fixes.
    ENABLE_LLM_FIX_FALLBACK: bool = True
    OPENAI_LLM_FIX_MODEL: str = "gpt-4.1"
    LLM_FIX_MAX_PASSES: int = 1

    class Config:
        env_file = ".env"
        extra = "ignore"


settings = Settings()
app = FastAPI(title="Template Cleaner", version="1.0.0")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/")
def root():
    return FileResponse(str(BASE_DIR / "static" / "index.html"))


@app.get("/healthz")
def healthz():
    return {
        "ok": True,
        "qc_model": settings.OPENAI_VISION_QC_MODEL,
        "llm_fix_model": settings.OPENAI_LLM_FIX_MODEL,
        "openai_key_configured": bool(settings.OPENAI_API_KEY),
    }


def _sse_event(data: dict) -> str:
    """Format a dict as an SSE data line."""
    return f"data: {json.dumps(data)}\n\n"


async def _clean_template_stream(
    template_bytes: bytes,
    template_filename: str,
    engine: str = "vision",
):
    """Async generator: clean uploaded PPTX via QC + deterministic fixes + LLM fallback.

    `engine` selects the layout detector:
      * "vision"   — GPT-4.1 vision QC on slide screenshots (default).
      * "geometry" — deterministic PDF-geometry analysis (no OpenAI key needed).

    Both engines emit the same QC schema and share the cleanup + verification
    pipeline. Every pass is guarded by a per-slide verification gate: after
    applying fixes we re-QC the affected slides and revert any slide that did not
    improve, so a bad edit can never ship. A running full-deck QC state
    (`qc_state`) is kept so regressions on previously-passing slides are caught.
    """
    qc_api_key = settings.OPENAI_API_KEY
    qc_base_url = settings.OPENAI_BASE_URL or None
    use_vision = engine != "geometry"
    if use_vision and not qc_api_key:
        yield _sse_event({"error": "Vision QC requires OPENAI_API_KEY in .env.", "done": True})
        return
    # LLM geometric fallback always needs an OpenAI key regardless of detector.
    llm_enabled = settings.ENABLE_LLM_FIX_FALLBACK and bool(qc_api_key)

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
        current_pptx = template_bytes

        cleanup_attempts: list[dict] = []
        llm_fallback_attempts: list[dict] = []
        reverted_slides_all: set[int] = set()

        def _state_failing(state: dict[int, dict]) -> set[int]:
            return {n for n, e in state.items() if str(e.get("status", "")).lower() == "fail"}

        def _actionable_failing(state: dict[int, dict]) -> set[int]:
            """
            Failing slides that have at least one issue a *move* can actually fix.

            Issues in ``_SOFT_ISSUE_TYPES`` (tiny fonts, minor alignment) cannot be
            resolved by repositioning shapes. A slide failing ONLY for those must
            not trigger more cleanup passes or the LLM geometric fallback — otherwise
            the model moves unrelated shapes (e.g. a title under the logo) chasing a
            problem it can't fix. Detection still reports them; we just stop acting.
            """
            out: set[int] = set()
            for n, e in state.items():
                if str(e.get("status", "")).lower() != "fail":
                    continue
                issues = e.get("issues") if isinstance(e.get("issues"), list) else []
                types = {str(i.get("type", "")).lower() for i in issues if isinstance(i, dict)}
                # No parsed types -> be safe and treat as actionable.
                if not types or (types - _SOFT_ISSUE_TYPES):
                    out.add(n)
            return out

        def _detect(pptx: bytes, subdir: str, slide_numbers: set[int] | None = None) -> dict:
            """Run the selected layout detector; both share the QC report schema."""
            if use_vision:
                report = run_visual_qc(
                    pptx,
                    os.path.join(td, subdir),
                    qc_api_key,
                    settings.OPENAI_VISION_QC_MODEL,
                    openai_base_url=qc_base_url,
                    slide_numbers=slide_numbers,
                )
            else:
                report = run_geometry_qc(
                    pptx,
                    os.path.join(td, subdir),
                    slide_numbers=slide_numbers,
                )
            # Always-on geometry signals (off-slide clipping + side-panel/footer
            # intrusions). The vision engine is blind to these; merging them in
            # lets BOTH engines catch and then repair them. Deterministic-engine
            # reports already include them, so the merge dedupes by issue type.
            return _augment_with_geometry_signals(report, pptx, slide_numbers)

        try:
            engine_label = "geometry" if not use_vision else "vision"
            yield _sse_event({"step": 0, "pct": 10, "label": "Template uploaded", "done": True})
            yield _sse_event({"step": 1, "pct": 25, "label": f"Running initial QC ({engine_label})..."})
            qc_initial = await asyncio.to_thread(_detect, current_pptx, "clean_qc_initial")
            if isinstance(qc_initial, dict) and qc_initial.get("overall_status") == "skipped":
                yield _sse_event(
                    {"error": f"QC unavailable: {qc_initial.get('skipped_reason', 'detector skipped')}", "done": True}
                )
                return
            if not isinstance(qc_initial, dict) or qc_initial.get("overall_status") != "fail":
                b64 = base64.b64encode(current_pptx).decode("ascii")
                yield _sse_event(
                    {
                        "step": 3,
                        "pct": 100,
                        "label": "Template already clean",
                        "done": True,
                        "file": b64,
                        "engine": engine_label,
                        "qc_initial": qc_initial,
                        "qc": qc_initial,
                        "cleanup": {"attempt_count": 0, "applied": False, "attempts": []},
                    }
                )
                return

            total_slides = int(qc_initial.get("total_slide_count") or 0)
            qc_state = index_slides(qc_initial)

            # ── Deterministic cleanup passes (each verified, with rollback) ──
            if settings.ENABLE_VISUAL_QC_REPAIR:
                for level in range(1, 3):
                    if not _actionable_failing(qc_state):
                        break
                    before_report = build_report_from_index(qc_state, total_slides)
                    yield _sse_event({"step": 2, "pct": 40 + level * 10, "label": f"Applying cleanup pass {level}..."})
                    candidate_pptx, clean_report = await asyncio.to_thread(
                        clean_template_from_qc,
                        current_pptx,
                        before_report,
                        repair_level=level,
                    )
                    cleanup_attempts.append(clean_report)
                    if not clean_report.get("applied"):
                        break

                    recheck = changed_slides_from_clean_report(clean_report) | _state_failing(qc_state)
                    if not recheck:
                        current_pptx = candidate_pptx
                        break

                    yield _sse_event(
                        {"step": 3, "pct": min(95, 65 + level * 10), "label": f"Verifying cleanup (pass {level})..."}
                    )
                    qc_after = await asyncio.to_thread(
                        _detect, candidate_pptx, f"clean_qc_{level}", recheck
                    )
                    after_index = index_slides(qc_after)
                    current_pptx, qc_state, reverted = apply_verification_gate(
                        baseline_bytes=current_pptx,
                        candidate_bytes=candidate_pptx,
                        before_index=qc_state,
                        after_index=after_index,
                        rechecked=recheck,
                    )
                    if reverted:
                        clean_report["reverted_slides"] = reverted
                        reverted_slides_all.update(reverted)
                        log.info("[Clean] Pass %s reverted regressed slides: %s", level, reverted)

            # ── LLM geometric fallback (verified, with rollback) ──
            if llm_enabled and _actionable_failing(qc_state):
                for llm_pass in range(1, max(1, int(settings.LLM_FIX_MAX_PASSES)) + 1):
                    unresolved = _actionable_failing(qc_state)
                    if not unresolved:
                        break
                    before_report = build_report_from_index(qc_state, total_slides)
                    yield _sse_event(
                        {"step": 2, "pct": min(97, 75 + llm_pass * 8), "label": f"Applying LLM fallback fixes (pass {llm_pass})..."}
                    )
                    candidate_pptx, llm_report = await asyncio.to_thread(
                        apply_llm_fallback_fixes,
                        current_pptx,
                        before_report,
                        openai_api_key=qc_api_key,
                        model=settings.OPENAI_LLM_FIX_MODEL,
                        work_dir=os.path.join(td, f"clean_llm_fallback_{llm_pass}"),
                        openai_base_url=qc_base_url,
                    )
                    llm_fallback_attempts.append(llm_report)
                    if not llm_report.get("applied"):
                        break

                    touched: set[int] = set()
                    for n in llm_report.get("touched_slide_numbers", []):
                        try:
                            v = int(n)
                        except (TypeError, ValueError):
                            continue
                        if v > 0:
                            touched.add(v)
                    recheck = touched or unresolved

                    yield _sse_event(
                        {"step": 3, "pct": min(98, 86 + llm_pass * 4), "label": f"Verifying LLM fallback (pass {llm_pass})..."}
                    )
                    qc_after = await asyncio.to_thread(
                        _detect, candidate_pptx, f"clean_qc_llm_{llm_pass}", recheck
                    )
                    after_index = index_slides(qc_after)
                    current_pptx, qc_state, reverted = apply_verification_gate(
                        baseline_bytes=current_pptx,
                        candidate_bytes=candidate_pptx,
                        before_index=qc_state,
                        after_index=after_index,
                        rechecked=recheck,
                    )
                    if reverted:
                        llm_report["reverted_slides"] = reverted
                        reverted_slides_all.update(reverted)
                        log.info("[Clean] LLM pass %s reverted regressed slides: %s", llm_pass, reverted)

            qc_final = build_report_from_index(qc_state, total_slides)
            initial_failed = int(qc_initial.get("failed_slide_count", 0))
            final_failed = int(qc_final.get("failed_slide_count", 0))
            improved = final_failed < initial_failed
            resolved = final_failed == 0
            if resolved:
                final_label = "Template cleanup complete"
            elif improved:
                final_label = "Template partially cleaned (some issues remain)"
            else:
                final_label = "Cleanup attempted (issues remain)"

            b64 = base64.b64encode(current_pptx).decode("ascii")
            yield _sse_event(
                {
                    "step": 3,
                    "pct": 100,
                    "label": final_label,
                    "done": True,
                    "file": b64,
                    "engine": engine_label,
                    "qc_initial": qc_initial,
                    "qc": qc_final,
                    "cleanup": {
                        "attempt_count": len(cleanup_attempts),
                        "applied": (
                            any(a.get("applied") for a in cleanup_attempts)
                            or any(a.get("applied") for a in llm_fallback_attempts)
                        ),
                        "attempts": cleanup_attempts,
                        "llm_fallback": {
                            "attempt_count": len(llm_fallback_attempts),
                            "applied": any(a.get("applied") for a in llm_fallback_attempts),
                            "attempts": llm_fallback_attempts,
                        },
                        "improved": improved,
                        "resolved": resolved,
                        "reverted_slides": sorted(reverted_slides_all),
                        "remaining_failures": final_failed,
                    },
                }
            )
        except Exception as e:
            log.exception("Template cleanup failed")
            yield _sse_event({"error": _user_friendly_error(e), "done": True})


@app.post("/clean-template")
async def clean_template(
    template: UploadFile = File(...),
    engine: str = Form("vision"),
):
    """Clean an uploaded PPTX via a QC + deterministic-fix + LLM-fallback loop.

    `engine`: "vision" (GPT-4.1 screenshots, default) or "geometry" (deterministic
    PDF-geometry analysis, no OpenAI key required for detection).
    """
    if not template.filename or not template.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Template must be .pptx")

    engine_norm = (engine or "vision").strip().lower()
    if engine_norm not in ("vision", "geometry"):
        engine_norm = "vision"
    if engine_norm == "vision" and not settings.OPENAI_API_KEY:
        raise HTTPException(500, "OPENAI_API_KEY is not configured (required for the Vision engine).")

    template_bytes = await template.read()

    return StreamingResponse(
        _clean_template_stream(template_bytes, template.filename or "template.pptx", engine_norm),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )
