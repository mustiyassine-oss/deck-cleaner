"""Per-slide verification gate with slide-level rollback.

After a cleanup/repair pass, we re-QC the affected slides and keep the new
version of a slide ONLY if it actually improved (or stayed equal). Any slide
that regressed is reverted to its pre-pass state, so a bad edit can never ship.

Rollback works at the OOXML package level: a cleanup pass edits shapes/text in
place without adding or removing slides or parts, so the baseline and candidate
.pptx share identical part names. Reverting a slide therefore means swapping that
slide's XML part (plus its directly related chart/embedding parts) back from the
baseline package into the candidate package. Slide parts are independent, so this
is safe and leaves every other slide's fixes intact.
"""
from __future__ import annotations

import io
import logging
import zipfile
from typing import Any

from pptx import Presentation

log = logging.getLogger("cleaner")

_SEVERITY_RANK = {"none": 0, "low": 1, "medium": 2, "high": 3}

# Part-name fragments that belong to a single slide and must be reverted with it.
_SLIDE_RELATED_FRAGMENTS = ("/charts/", "/embeddings/", "/drawings/")


def severity_rank(entry: dict[str, Any] | None) -> int:
    if not isinstance(entry, dict):
        return 0
    return _SEVERITY_RANK.get(str(entry.get("severity", "none")).lower(), 0)


def _issue_count(entry: dict[str, Any] | None) -> int:
    if not isinstance(entry, dict):
        return 0
    issues = entry.get("issues")
    return len(issues) if isinstance(issues, list) else 0


def _is_fail(entry: dict[str, Any] | None) -> bool:
    return isinstance(entry, dict) and str(entry.get("status", "")).lower() == "fail"


def index_slides(qc_report: dict[str, Any] | None) -> dict[int, dict[str, Any]]:
    """Map slide_number -> slide QC entry."""
    out: dict[int, dict[str, Any]] = {}
    if not isinstance(qc_report, dict):
        return out
    slides = qc_report.get("slides") if isinstance(qc_report.get("slides"), list) else []
    for s in slides:
        if not isinstance(s, dict):
            continue
        try:
            n = int(s.get("slide_number"))
        except (TypeError, ValueError):
            continue
        if n > 0:
            out[n] = s
    return out


def is_regression(before: dict[str, Any] | None, after: dict[str, Any] | None) -> bool:
    """
    True if `after` is worse than `before` for the same slide.

    - Breaking a previously clean slide is always a regression.
    - For a slide that was already failing, it's a regression only if it got
      strictly worse (more issues, or higher max severity). Equal or improved
      states are kept (equal = harmless lateral change; improved = the goal).
    """
    before_fail = _is_fail(before)
    after_fail = _is_fail(after)
    if not after_fail:
        return False
    if after_fail and not before_fail:
        return True
    # both failing
    if _issue_count(after) > _issue_count(before):
        return True
    if severity_rank(after) > severity_rank(before):
        return True
    return False


def build_report_from_index(
    slide_index: dict[int, dict[str, Any]],
    total_slide_count: int,
) -> dict[str, Any]:
    """Reassemble a full QC report dict from a slide_number -> entry index."""
    slides = [slide_index[n] for n in sorted(slide_index)]
    failed = sum(1 for s in slides if _is_fail(s))
    return {
        "overall_status": "fail" if failed else "pass",
        "failed_slide_count": failed,
        "total_slide_count": total_slide_count,
        "checked_slide_count": len(slides),
        "checked_slide_numbers": sorted(slide_index),
        "slides": slides,
    }


def changed_slides_from_clean_report(clean_report: dict[str, Any] | None) -> set[int]:
    """Slide numbers a deterministic cleanup pass actually modified."""
    out: set[int] = set()
    if not isinstance(clean_report, dict):
        return out
    for e in clean_report.get("slides") or []:
        if not isinstance(e, dict) or not e.get("actions"):
            continue
        try:
            n = int(e.get("slide_number"))
        except (TypeError, ValueError):
            continue
        if n > 0:
            out.add(n)
    return out


def _zip_name(partname: Any) -> str:
    """Convert a PackURI ('/ppt/slides/slide1.xml') to a zip entry name."""
    s = str(partname)
    return s[1:] if s.startswith("/") else s


def _slide_part_paths(pptx_bytes: bytes) -> dict[int, set[str]]:
    """
    Map 1-based slide number -> set of zip entry names that constitute that slide
    (its XML part plus directly related chart/embedding/drawing parts and their
    own sub-parts).
    """
    paths: dict[int, set[str]] = {}
    try:
        pres = Presentation(io.BytesIO(pptx_bytes))
    except Exception as e:
        log.warning("[Verify] Could not open PPTX to map slide parts: %s", e)
        return paths

    for idx, slide in enumerate(pres.slides):
        slide_no = idx + 1
        names: set[str] = set()
        try:
            names.add(_zip_name(slide.part.partname))
        except Exception:
            continue
        try:
            for rel in slide.part.rels.values():
                if getattr(rel, "is_external", False):
                    continue
                try:
                    target = rel.target_part
                except Exception:
                    continue
                pn = _zip_name(target.partname)
                if any(frag in "/" + pn for frag in _SLIDE_RELATED_FRAGMENTS):
                    names.add(pn)
                    # include the related part's own sub-parts (e.g. chart colors,
                    # style, embedded workbook) so reverts are complete.
                    try:
                        for sub in target.rels.values():
                            if getattr(sub, "is_external", False):
                                continue
                            names.add(_zip_name(sub.target_part.partname))
                    except Exception:
                        pass
        except Exception:
            pass
        paths[slide_no] = names
    return paths


def revert_slides(
    baseline_bytes: bytes,
    candidate_bytes: bytes,
    slide_numbers: set[int] | list[int] | tuple[int, ...],
) -> bytes:
    """
    Return a PPTX equal to `candidate_bytes` except that the given slides are
    restored from `baseline_bytes`. If anything goes wrong, returns the candidate
    unchanged (fail-safe).
    """
    targets = {int(n) for n in (slide_numbers or []) if int(n) > 0}
    if not targets:
        return candidate_bytes

    slide_paths = _slide_part_paths(baseline_bytes)
    restore: set[str] = set()
    for n in targets:
        restore |= slide_paths.get(n, set())
    if not restore:
        log.warning("[Verify] No part names resolved for revert of slides %s", sorted(targets))
        return candidate_bytes

    try:
        with zipfile.ZipFile(io.BytesIO(baseline_bytes)) as zbase:
            base_names = set(zbase.namelist())
            base_data = {name: zbase.read(name) for name in restore if name in base_names}

        missing = restore - set(base_data)
        if missing:
            log.warning("[Verify] Baseline missing parts for revert (skipping those): %s", sorted(missing))

        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(candidate_bytes)) as zcand:
            with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zcand.infolist():
                    data = base_data.get(item.filename)
                    if data is None:
                        data = zcand.read(item.filename)
                    # Preserve original metadata (date, external attrs, comment).
                    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                    zi.compress_type = item.compress_type
                    zi.external_attr = item.external_attr
                    zi.internal_attr = item.internal_attr
                    zi.create_system = item.create_system
                    zout.writestr(zi, data)
        log.info("[Verify] Reverted %d slide(s) %s (%d part(s))", len(targets), sorted(targets), len(base_data))
        return out.getvalue()
    except Exception as e:
        log.warning("[Verify] Slide revert failed, keeping candidate: %s", e)
        return candidate_bytes


def _overlap_counts(pptx_bytes: bytes) -> dict[int, int]:
    """
    Deterministic per-slide collision counts: text-vs-text overlaps PLUS
    text-vs-picture overlaps (e.g. a title pushed under the logo). Empty dict on
    any failure so the gate degrades to QC-only comparison.
    """
    try:
        from pdf_geometry import overlap_counts, picture_overlap_counts
        text_text = overlap_counts(pptx_bytes)
        text_pic = picture_overlap_counts(pptx_bytes)
        keys = set(text_text) | set(text_pic)
        return {k: text_text.get(k, 0) + text_pic.get(k, 0) for k in keys}
    except Exception as e:
        log.warning("[Verify] Overlap cross-check unavailable: %s", e)
        return {}


def apply_verification_gate(
    *,
    baseline_bytes: bytes,
    candidate_bytes: bytes,
    before_index: dict[int, dict[str, Any]],
    after_index: dict[int, dict[str, Any]],
    rechecked: set[int],
) -> tuple[bytes, dict[int, dict[str, Any]], list[int]]:
    """
    Compare before/after for every rechecked slide, revert regressions, and
    return (possibly-reverted PPTX bytes, updated slide index, reverted slides).

    A slide is a regression if EITHER the QC detector got worse (`is_regression`)
    OR a deterministic collision was newly introduced — text-over-text OR
    text-over-picture (e.g. a title pushed under the logo). The cross-check is
    engine-agnostic: the vision QC cannot see geometric overlap and neither QC sees
    text-over-image, so this catches fixes (e.g. LLM moves) the detector misses.

    The returned index reflects reality: kept slides take their after-state,
    reverted slides keep their before-state.
    """
    overlap_before = _overlap_counts(baseline_bytes)
    overlap_after = _overlap_counts(candidate_bytes)

    regressed = [
        n for n in sorted(rechecked)
        if is_regression(before_index.get(n), after_index.get(n))
        or overlap_after.get(n, 0) > overlap_before.get(n, 0)
    ]

    new_index = dict(before_index)
    for n in rechecked:
        if n in regressed:
            new_index[n] = before_index.get(n, after_index.get(n, {"slide_number": n, "status": "fail", "issues": []}))
        elif n in after_index:
            new_index[n] = after_index[n]

    if not regressed:
        return candidate_bytes, new_index, []

    reverted_bytes = revert_slides(baseline_bytes, candidate_bytes, set(regressed))
    return reverted_bytes, new_index, regressed
