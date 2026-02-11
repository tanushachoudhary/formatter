"""
FastAPI backend for the Legal Document Formatter.
Run from project root: uvicorn api.main:app --reload --port 8000
"""
import os
import sys

from dotenv import load_dotenv
load_dotenv()

# Ensure project root is on path when running as api.main
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel

from backend import extract_and_store_styles, process_document
from utils.html_to_docx import html_to_docx_bytes, normalize_editor_html, plain_text_to_simple_html

app = FastAPI(title="Legal Document Formatter API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class DocxFromHtmlBody(BaseModel):
    html: str


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/extract-styles")
async def api_extract_styles(file: UploadFile = File(...)):
    """Upload a DOCX template; returns extracted style schema (no file stored)."""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Upload a .docx file")
    try:
        os.makedirs(os.path.join(_ROOT, "output"), exist_ok=True)
        contents = await file.read()
        with open(os.path.join(_ROOT, "output", "upload_template.docx"), "wb") as f:
            f.write(contents)
        with open(os.path.join(_ROOT, "output", "upload_template.docx"), "rb") as f:
            schema = extract_and_store_styles(f)
        # Remove heavy keys for response (template_page_images, etc.)
        out = {k: v for k, v in schema.items() if k not in ("template_page_images", "template_page_ocr_texts")}
        return out
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/format")
async def api_format(file: UploadFile = File(...), text: str = Form(...)):
    """Format raw text with the template; returns preview_text and docx as base64."""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Upload a .docx template")
    try:
        os.makedirs(os.path.join(_ROOT, "output"), exist_ok=True)
        contents = await file.read()
        with open(os.path.join(_ROOT, "output", "upload_template.docx"), "wb") as f:
            f.write(contents)
        with open(os.path.join(_ROOT, "output", "upload_template.docx"), "rb") as f:
            output_path, preview_text = process_document(text.strip(), f)
        with open(output_path, "rb") as f:
            docx_bytes = f.read()
        import base64
        docx_b64 = base64.b64encode(docx_bytes).decode("ascii")
        html = plain_text_to_simple_html(preview_text)
        return {"preview_text": preview_text, "preview_html": html, "docx_base64": docx_b64}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/docx-from-html")
async def api_docx_from_html(body: DocxFromHtmlBody):
    """Build a DOCX from editor HTML (e.g. after user edits). Returns DOCX file."""
    try:
        html_norm = normalize_editor_html(body.html or "")
        docx_bytes = html_to_docx_bytes(html_norm)
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=formatted_output.docx"},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
