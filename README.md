# Template Cleaner (PPTX layout cleanup)

Upload a PowerPoint deck → get back the **same deck with layout defects repaired**.
This service is **cleaning-only**: it runs visual QC on an existing `.pptx`, applies
deterministic geometric fixes, and (optionally) an LLM geometric fallback pass for
slides that still fail. It never generates new content — it only repairs layout.

## What it fixes

- **Chart legends overlaying the plot** — legends drawn on top of the bars/slices are
  given their own space so they no longer overshadow the chart.
- **Overlapping cards** — two filled card backgrounds colliding (the card + the text
  inside it are moved as one unit).
- **Text poking into a neighbouring card** — a stray text box spilling into an adjacent
  card is trimmed/reflowed back into its own column.
- **Text overflowing its card** — an over-stuffed card's body text is shrunk to fit.
- **Text-vs-text overlap**, **off-slide / clipped content**, and **side-panel / footer-band
  intrusions** — content is nudged back into its safe region.

Every fix is verified per-slide and **rolled back automatically if it doesn't improve
the slide**, so a bad edit can never ship.

## How it works

```
PPTX  ──►  QC detector  ──►  deterministic cleanup  ──►  LLM fallback  ──►  verify gate  ──►  PPTX
            (per slide)        (move / shrink / fit)      (safe moves)       (revert regressions)
```

Two interchangeable **detectors** emit the same report schema:

| Engine | How it detects | Needs |
|--------|----------------|-------|
| `vision` (default) | LibreOffice → PDF → PNG → GPT-4.1 reads each slide screenshot | OpenAI key, LibreOffice, Poppler |
| `geometry` | Deterministic `python-pptx` object-model geometry (+ optional PyMuPDF rendered-text pass) | nothing required (PDF pass needs LibreOffice + PyMuPDF; degrades gracefully) |

Both engines feed the **same repair pipeline** (`template_cleaner.py`) and the same
**verification gate** (`verify.py`), which re-QCs touched slides and reverts any that
regressed (including a deterministic text-vs-text and text-vs-picture overlap cross-check
the detectors can't see on their own).

## Requirements

Python deps (see `requirements.txt`):

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
```

External tools (install separately — not pip-installable):

- **LibreOffice** — used to render `.pptx` → `.pdf`. Required for the **vision** engine and
  for the **geometry** engine's optional rendered-PDF complement.
  - Windows: install LibreOffice (auto-detected at `C:\Program Files\LibreOffice\...`).
  - macOS/Linux: ensure `soffice`/`libreoffice` is on `PATH`.
- **Poppler** — required by `pdf2image` to turn the PDF into PNG screenshots (vision engine).
  - Windows: install Poppler and put its `bin` on `PATH`.
  - macOS: `brew install poppler` · Debian/Ubuntu: `apt-get install poppler-utils`.

> The **geometry** engine runs with **no external tools and no API key** for its core
> object-model checks; LibreOffice/PyMuPDF only add the optional rendered-text pass.

## Configuration

Copy `.env.example` → `.env` and set values as needed:

| Variable | Default | Purpose |
|----------|---------|---------|
| `OPENAI_API_KEY` | — | Required for the **vision** engine and the LLM fallback. |
| `OPENAI_BASE_URL` | — | Optional custom OpenAI-compatible endpoint. |
| `OPENAI_VISION_QC_MODEL` | `gpt-4.1` | Vision QC model. |
| `ENABLE_VISUAL_QC_REPAIR` | `true` | Run the deterministic cleanup passes. |
| `ENABLE_LLM_FIX_FALLBACK` | `true` | Run the LLM geometric fallback for slides still failing. |
| `OPENAI_LLM_FIX_MODEL` | `gpt-4.1` | Model for the LLM fallback. |
| `LLM_FIX_MAX_PASSES` | `1` | Max LLM fallback passes. |

## Run

```bash
uvicorn main:app --reload --port 8080
```

On Windows you can also double-click **`run.bat`**, which activates `.venv` (if present),
starts the server on `http://127.0.0.1:8080`, and opens the browser.

Then open `http://127.0.0.1:8080` for the upload UI.

## API

- `GET /` — upload UI (`static/index.html`).
- `GET /healthz` — service + model/key status.
- `POST /clean-template` — multipart form, returns a **Server-Sent Events** progress
  stream ending with the cleaned deck as base64.
  - `template`: the `.pptx` file (required).
  - `engine`: `vision` (default) or `geometry`.

Example:

```bash
curl -N -X POST "http://127.0.0.1:8080/clean-template" \
  -F "template=@deck.pptx" \
  -F "engine=geometry"
```

## Project layout

| File | Responsibility |
|------|----------------|
| `main.py` | FastAPI app, SSE cleanup pipeline, engine selection, always-on geometry signals. |
| `qc.py` | Vision QC: PPTX→PDF→PNG + GPT-4.1 per-slide review (+ tiny/dense-text guardrail). |
| `pdf_geometry.py` | Geometry QC: object-model overlap/clipping/keep-out/card/chart detection + optional PyMuPDF pass. |
| `template_cleaner.py` | Deterministic fixers: de-overlap, off-slide reflow, panel/footer clearance, card-cluster collision handling, chart-legend normalization, shrink-to-fit. |
| `llm_fix_fallback.py` | LLM-planned, strictly-guardrailed "move" operations for slides still failing. |
| `verify.py` | Per-slide verification gate with OOXML-level rollback. |
| `static/index.html` | Browser upload UI. |

## Notes & limitations

- Pure object-model geometry can't see rendered-only effects (e.g. a font that renders
  larger than declared); those rely on the vision engine or the optional PyMuPDF pass.
- The cleaner repositions/shrinks existing shapes only — it never rewrites copy or adds
  content.
