# ProposalGPT — Document Generation (Python)

A **Docmosis replacement**: merge ProposalGPT `proposalData` JSON into a Word
`.docx` template and load the generated document in the **SuperDoc** editor.

Scope is **document generation only** — no collaboration, realtime sync, or
versioning. Pure Python merge; SuperDoc runs in the browser to view/edit/download.

## Stack
- **FastAPI + Uvicorn** — REST API + serves the dashboard/editor pages.
- **docxtpl (Jinja2)** — fills `{{ tags }}`, `{% for %}` loops, `{% if %}` conditions.
- **python-docx** — builds the example template and converts rich HTML → Word
  paragraphs/tables (injected in place of markers, so subdoc fidelity is exact).
- **SuperDoc** (CDN) — loads `base.docx` with Word fidelity; single-user editing.

## How it works
1. Phase 1: docxtpl renders all text + structure (metadata, auto-numbered
   sections `1 / 1.1 / 1.1.1`, loops, conditionals). Rich fields emit unique markers.
2. Phase 2: each marker paragraph is replaced with real Word content converted
   from the field's HTML (`<p> <b> <i> <ul>/<ol> <table>`), mirroring Docmosis output.

## Setup & run
```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
# open http://localhost:8000
```

## API
| Method | Path | Purpose |
|---|---|---|
| GET  | `/api/health` | health check |
| GET  | `/api/sample` | bundled sample `proposalData` |
| GET  | `/api/template/example` | download the example `.docx` template |
| POST | `/api/templates` (multipart `docx`) | store a template → `{id}` |
| POST | `/api/documents` | generate — JSON `{templateId?, proposalData}` or multipart `docx` → `{id, embedUrl}` |
| GET  | `/api/documents/:id` | document metadata |
| GET  | `/api/documents/:id/base.docx` | the generated `.docx` SuperDoc loads |
| GET  | `/` , `/editor?doc=:id` | dashboard / embedded editor |

## Data model
See the documented `proposalData` schema: top-level `metadata`, `sections[]`
(`title` + `section_html` + nested `subsections[]`), optional
`has_past_performance` + `past_performance[]` and `has_resumes` + `resumes[]`.
`app/sample_proposal.json` is a complete example.

## Authoring a template
Download the example, restyle it in Word, keep the `{{ tags }}`:
- Values: `{{ company_name }}`, `{{ title }}`, `{{ metadata.company_name }}`
- Sections loop with auto-number: `{% for s in sections %}` … `{{ s.number }} {{ s.title }}` … `{{ s.body }}`
- Conditionals: `{% if has_resumes %}` … `{% endif %}`
- Loop/condition controls use docxtpl's paragraph tag form `{%p ... %}` so the
  control lines disappear cleanly.

## Tests
```bash
python -m tests.test_merge      # or: pytest -q
```

## Notes
- Storage is in-memory (generation-only). Swap `TEMPLATES`/`DOCUMENTS` in
  `app/main.py` for a DB if persistence is needed.
- The editor can run **fully offline** (bundled SuperDoc) or fall back to a CDN.

## Offline editor (optional)
By default `/editor` serves a bundled SuperDoc build if present, otherwise it
falls back to loading SuperDoc from `esm.sh` (needs internet). To build the
offline bundle:

```bash
cd frontend
npm install
npm run build        # outputs app/static/dist/{editor.html, editor-*.js, editor-*.css}
```

FastAPI then serves `app/static/dist/editor.html` at `/editor` and the hashed
assets at `/assets`. `node_modules/` and `app/static/dist/` are git-ignored —
run the build after cloning to enable offline mode.
