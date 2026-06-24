"""Merge ProposalGPT ``proposalData`` JSON into a Word template (Docmosis replacement).

Two-phase render:

1. **Text + structure** via docxtpl (Jinja2 tags): metadata, auto-numbered
   sections (1, 1.1, 1.1.1), loops and conditionals. Rich fields are emitted as
   unique one-line **markers**.
2. **Rich content injection**: each marker paragraph is replaced in place with
   real Word paragraphs/tables converted from the field's HTML (``html_to_docx``).

This avoids docxtpl's broken subdoc insertion while keeping full Word fidelity.
Plain-text fields are HTML-entity decoded; missing data renders empty, never raises.
"""
from __future__ import annotations

import copy
import html as html_lib
import io
import uuid

from docx import Document
from docx.oxml.ns import qn
from docxtpl import DocxTemplate

from .html_to_docx import build_html_doc

_MARKER_PREFIX = "@@RICH:"
_MARKER_SUFFIX = "@@"


def merge(template_bytes: bytes, proposal_data: dict) -> bytes:
    """Render ``proposal_data`` into ``template_bytes``; return .docx bytes."""
    data = _decode_deep(proposal_data or {})
    markers: dict[str, str] = {}

    tpl = DocxTemplate(io.BytesIO(template_bytes))
    context = _build_context(data, markers)
    # autoescape so plain-text values containing &, <, > stay valid in the XML
    tpl.render(context, autoescape=True)

    rendered = io.BytesIO()
    tpl.save(rendered)

    return _inject_rich(rendered.getvalue(), markers)


# --------------------------------------------------------------------------- #
# Phase 1: build the docxtpl context (text + markers for rich fields)
# --------------------------------------------------------------------------- #
def _rich(markers, raw_html):
    """Register HTML under a unique marker and return the marker text."""
    text = (raw_html or "").strip()
    if not text:
        return ""
    token = f"{_MARKER_PREFIX}{uuid.uuid4().hex}{_MARKER_SUFFIX}"
    markers[token] = text
    return token


def _build_context(data, markers):
    meta = data.get("metadata") or {}
    past = data.get("past_performance") or []
    resumes = data.get("resumes") or []
    return {
        "title": meta.get("title", ""),
        "solicitation_number": meta.get("solicitation_number", ""),
        "date": meta.get("date", ""),
        "company_name": meta.get("company_name", ""),
        "company_city": meta.get("company_city", ""),
        "company_state": meta.get("company_state", ""),
        "company_zip": meta.get("company_zip", ""),
        "company_contact_person": meta.get("company_contact_person", ""),
        "company_contact_email": meta.get("company_contact_email", ""),
        "company_contact_phone": meta.get("company_contact_phone", ""),
        "metadata": meta,
        "sections": _build_sections(data.get("sections") or [], markers),
        "has_past_performance": _present(data.get("has_past_performance"), past),
        "past_performance": [_past(p, markers) for p in past],
        "has_resumes": _present(data.get("has_resumes"), resumes),
        "resumes": [_resume(r, markers) for r in resumes],
    }


def _build_sections(arr, markers, prefix=""):
    out = []
    for i, s in enumerate(arr, start=1):
        number = str(i) if not prefix else f"{prefix}.{i}"
        children = s.get("subsections") or s.get("subsubsections") or []
        out.append({
            "number": number,
            "title": str(s.get("title") or ""),
            "body": _rich(markers, s.get("section_html")),
            "subsections": _build_sections(children, markers, number),
        })
    return out


def _past(p, markers):
    return {
        "award_title": p.get("award_title", ""),
        "award_recipient_name": p.get("award_recipient_name", ""),
        "role": p.get("role", ""),
        "task_order_number": p.get("task_order_number", ""),
        "total_contract_value": p.get("total_contract_value", ""),
        "period_of_performance_start": p.get("period_of_performance_start", ""),
        "period_of_performance_end": p.get("period_of_performance_end", ""),
        "naics": p.get("naics", ""),
        "psc": p.get("psc", ""),
        "customer_poc": p.get("customer_poc", ""),
        "body": _rich(markers, p.get("description")),
    }


def _resume(r, markers):
    first = (r.get("first_name") or "").strip()
    last = (r.get("last_name") or "").strip()
    full = (r.get("full_name") or f"{first} {last}").strip()
    return {
        "full_name": full,
        "first_name": first,
        "last_name": last,
        "job_title": r.get("job_title", ""),
        "citizenship": r.get("citizenship", ""),
        "professional_experience": [
            {
                "position": e.get("position", ""),
                "company_name": e.get("company_name", ""),
                "start_date": e.get("start_date", ""),
                "end_date": e.get("end_date", ""),
                "body": _rich(markers, e.get("description")),
            }
            for e in (r.get("professional_experience") or [])
        ],
        "education": [
            {"degree": ed.get("degree", ""), "institution_name": ed.get("institution_name", "")}
            for ed in (r.get("education") or [])
        ],
        "technical_skills": r.get("technical_skills", ""),
        "certifications_and_courses": r.get("certifications_and_courses", ""),
        "notable_projects": r.get("notable_projects", ""),
    }


# --------------------------------------------------------------------------- #
# Phase 2: replace marker paragraphs with rich Word content
# --------------------------------------------------------------------------- #
def _inject_rich(docx_bytes, markers):
    if not markers:
        return docx_bytes

    doc = Document(io.BytesIO(docx_bytes))
    for para in list(doc.paragraphs):
        token = para.text.strip()
        if token not in markers:
            continue
        anchor = para._p
        scratch = build_html_doc(markers[token])
        for child in list(scratch.element.body):
            if child.tag == qn("w:sectPr"):
                continue
            anchor.addprevious(copy.deepcopy(child))
        anchor.getparent().remove(anchor)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _present(flag, arr):
    return bool(flag) and len(arr) > 0


# HTML-keyed fields hold rich markup handled in phase 2; everything else is text.
_RICH_KEYS = {"section_html", "description"}


def _decode_deep(value):
    """Recursively HTML-unescape plain-text strings (e.g. ``&amp;`` -> ``&``)."""
    if isinstance(value, dict):
        return {k: (v if k in _RICH_KEYS else _decode_deep(v)) for k, v in value.items()}
    if isinstance(value, list):
        return [_decode_deep(v) for v in value]
    if isinstance(value, str):
        return html_lib.unescape(value)
    return value
