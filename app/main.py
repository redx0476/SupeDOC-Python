"""FastAPI service: generate proposal .docx files and load them in SuperDoc.

Scope: document generation only (Docmosis replacement). No collaboration,
versioning, or realtime sync. Templates and generated documents are kept in
memory (generation-only); swap ``STORE`` for a DB if persistence is needed.
"""
from __future__ import annotations

import io
import json
import secrets
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, Response

from .example_template import build_example_template
from .merge import merge

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
APP_DIR = Path(__file__).resolve().parent

app = FastAPI(title="ProposalGPT Document Generation", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# In-memory stores (generation-only).
TEMPLATES: dict[str, dict] = {}
DOCUMENTS: dict[str, dict] = {}


def _id(prefix=""):
    return prefix + secrets.token_urlsafe(7)[:10]


def _docx_response(data: bytes, filename: str, inline=False):
    disp = "inline" if inline else "attachment"
    return Response(
        content=data,
        media_type=DOCX_MIME,
        headers={"Content-Disposition": f'{disp}; filename="{filename}"'},
    )


# --------------------------------------------------------------------------- #
# Health / sample
# --------------------------------------------------------------------------- #
@app.get("/api/health")
def health():
    return {"ok": True, "framework": "fastapi", "mode": "generation-only"}


@app.get("/api/sample")
def sample():
    return JSONResponse(json.loads((APP_DIR / "sample_proposal.json").read_text()))


# --------------------------------------------------------------------------- #
# Templates
# --------------------------------------------------------------------------- #
@app.get("/api/template/example")
def template_example():
    return _docx_response(build_example_template(), "ProposalGPT-Example-Template.docx")


@app.post("/api/templates")
async def upload_template(docx: UploadFile = File(...)):
    if not docx:
        raise HTTPException(status_code=400, detail="docx file required")
    data = await docx.read()
    tid = _id("t")
    TEMPLATES[tid] = {"id": tid, "name": docx.filename, "docx": data}
    return JSONResponse({"id": tid, "name": docx.filename}, status_code=201)


# --------------------------------------------------------------------------- #
# Documents (generation)
# --------------------------------------------------------------------------- #
@app.post("/api/documents")
async def create_document(
    request: Request,
    docx: UploadFile | None = File(default=None),
    templateId: str | None = Form(default=None),
    proposalData: str | None = Form(default=None),
    title: str | None = Form(default=None),
):
    """Create a generated document.

    Accepts either JSON ``{templateId?, proposalData}`` or multipart
    (``docx`` upload, or ``templateId`` + ``proposalData`` form fields).
    """
    body_template_id = templateId
    data = _parse_form_json(proposalData)
    doc_title = title or "Untitled Proposal"

    if docx is not None:
        base = await docx.read()
        seed_mode = "docx-upload"
        doc_title = title or Path(docx.filename or "proposal").stem
    else:
        if data is None and request.headers.get("content-type", "").startswith("application/json"):
            payload = await request.json()
            body_template_id = body_template_id or payload.get("templateId")
            data = payload.get("proposalData")
            doc_title = payload.get("title") or doc_title

        data = data or {}
        doc_title = (data.get("metadata") or {}).get("title") or doc_title

        if body_template_id:
            tmpl = TEMPLATES.get(body_template_id)
            if not tmpl:
                raise HTTPException(status_code=400, detail="unknown templateId")
            seed_mode = "template-merge"
            base = merge(tmpl["docx"], data)
        else:
            seed_mode = "json-render"
            base = merge(build_example_template(), data)

    did = _id()
    DOCUMENTS[did] = {"id": did, "title": doc_title, "seedMode": seed_mode, "base": base}
    embed = f"{str(request.base_url).rstrip('/')}/editor?doc={did}"
    return JSONResponse(
        {"id": did, "title": doc_title, "seedMode": seed_mode, "embedUrl": embed},
        status_code=201,
    )


@app.get("/api/documents/{doc_id}")
def get_document(doc_id: str, request: Request):
    doc = DOCUMENTS.get(doc_id)
    if not doc:
        raise HTTPException(status_code=404, detail="not found")
    embed = f"{str(request.base_url).rstrip('/')}/editor?doc={doc_id}"
    return {"id": doc_id, "title": doc["title"], "seedMode": doc["seedMode"], "embedUrl": embed}


@app.get("/api/documents/{doc_id}/base.docx")
def get_document_base(doc_id: str):
    doc = DOCUMENTS.get(doc_id)
    if not doc:
        raise HTTPException(status_code=404, detail="not found")
    return _docx_response(doc["base"], "proposal.docx", inline=True)


def _parse_form_json(value):
    if not value:
        return None
    try:
        return json.loads(value)
    except (ValueError, TypeError):
        return None


# --------------------------------------------------------------------------- #
# Frontend pages
# --------------------------------------------------------------------------- #
@app.get("/", response_class=HTMLResponse)
def dashboard():
    return (APP_DIR / "static" / "dashboard.html").read_text()


@app.get("/editor", response_class=HTMLResponse)
def editor():
    return (APP_DIR / "static" / "editor.html").read_text()
