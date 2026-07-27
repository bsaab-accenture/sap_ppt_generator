"""
docx_core/triggers.py
─────────────────────
Azure Functions Blueprint — DOCX converter endpoints.

Endpoints
---------
POST /api/docx_to_txt
    Accept a base64-encoded .docx file and return full plain text.
    Preserves document order: headers → body (paragraphs + tables inline) → footers.

Registration
------------
In function_app.py add:
    from docx_core.triggers import bp as docx_bp
    app.register_blueprint(docx_bp)
"""

from __future__ import annotations

import base64
import io
import json
import logging

from docx import Document
from docx.oxml.ns import qn

log = logging.getLogger(__name__)

# azure.functions is only available inside the Function App runtime.
# When importing this module for local tests, we skip the Blueprint setup.
try:
    import azure.functions as func
    bp = func.Blueprint()
    _AZURE_AVAILABLE = True
except ImportError:
    func = None  # type: ignore[assignment]
    bp = None    # type: ignore[assignment]
    _AZURE_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# Extraction logic — importable directly for local tests
# ─────────────────────────────────────────────────────────────────────────────

def extract_text_from_docx(doc_bytes: bytes) -> str:
    """
    Convert a .docx binary to plain text with zero content loss.

    Extraction order:
      1. Section headers (labelled === HEADER ===)
      2. Document body in original order — paragraphs joined by \\n,
         tables formatted as pipe-delimited rows, inline where they appear.
      3. Section footers (labelled === FOOTER ===)
    """
    doc = Document(io.BytesIO(doc_bytes))
    parts: list[str] = []

    # ── 1. Headers ────────────────────────────────────────────────────────────
    header_lines: list[str] = []
    for section in doc.sections:
        try:
            if not section.header.is_linked_to_previous:
                text = _extract_part_text(section.header)
                if text:
                    header_lines.append(text)
        except Exception:
            pass
    if header_lines:
        parts.append("=== HEADER ===")
        parts.extend(header_lines)
        parts.append("=== END HEADER ===")
        parts.append("")

    # ── 2. Body in document order ─────────────────────────────────────────────
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            # Collect all w:t text nodes, preserving runs and spaces
            text = ''.join(node.text or '' for node in child.iter(qn('w:t')))
            parts.append(text)

        elif tag == 'tbl':
            for tr in child.iter(qn('w:tr')):
                cells: list[str] = []
                for tc in tr.iter(qn('w:tc')):
                    cell_text = ''.join(
                        node.text or '' for node in tc.iter(qn('w:t'))
                    ).strip()
                    cells.append(cell_text)
                if cells:
                    parts.append('| ' + ' | '.join(cells) + ' |')
            parts.append('')  # blank line after table

    # ── 3. Footers ────────────────────────────────────────────────────────────
    footer_lines: list[str] = []
    for section in doc.sections:
        try:
            if not section.footer.is_linked_to_previous:
                text = _extract_part_text(section.footer)
                if text:
                    footer_lines.append(text)
        except Exception:
            pass
    if footer_lines:
        parts.append("")
        parts.append("=== FOOTER ===")
        parts.extend(footer_lines)
        parts.append("=== END FOOTER ===")

    return '\n'.join(parts)


def _extract_part_text(part) -> str:
    """Extract plain text from a header/footer part, preserving paragraph breaks."""
    lines: list[str] = []
    for child in part._element:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            text = ''.join(node.text or '' for node in child.iter(qn('w:t')))
            lines.append(text)
        elif tag == 'tbl':
            for tr in child.iter(qn('w:tr')):
                cells = [
                    ''.join(node.text or '' for node in tc.iter(qn('w:t'))).strip()
                    for tc in tr.iter(qn('w:tc'))
                ]
                if cells:
                    lines.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(lines).strip()


# ─────────────────────────────────────────────────────────────────────────────
# HTTP endpoint — only wired up when running inside the Azure Functions runtime
# ─────────────────────────────────────────────────────────────────────────────

if _AZURE_AVAILABLE:

    @bp.route(route="docx_to_txt", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
    def docx_to_txt(req: func.HttpRequest) -> func.HttpResponse:
        """POST /api/docx_to_txt — convert base64-encoded .docx to plain text."""
        log.info("[docx_to_txt] Request received")

        # ── Parse JSON body ───────────────────────────────────────────────────
        try:
            body = req.get_json()
        except ValueError:
            return _json_error(400, "BadRequest", "Request body is not valid JSON.")

        docx_b64: str | None = body.get("docx_base64") if body else None
        if not docx_b64:
            return _json_error(400, "BadRequest", "Missing required field: docx_base64")

        # ── Decode base64 → bytes ─────────────────────────────────────────────
        try:
            doc_bytes = base64.b64decode(docx_b64)
        except Exception as exc:
            return _json_error(400, "BadRequest", f"Invalid base64 encoding: {exc}")

        log.info("[docx_to_txt] Input size: %d bytes", len(doc_bytes))

        # ── Extract text ──────────────────────────────────────────────────────
        try:
            text = extract_text_from_docx(doc_bytes)
        except Exception as exc:
            log.exception("[docx_to_txt] Extraction failed: %s", exc)
            return _json_error(500, "InternalError", f"Failed to extract text: {exc}")

        log.info("[docx_to_txt] Extraction complete. Output: %d chars", len(text))

        return func.HttpResponse(
            body=text.encode("utf-8"),
            status_code=200,
            mimetype="text/plain; charset=utf-8",
            headers={"Content-Disposition": 'attachment; filename="transcript.txt"'},
        )

    def _json_error(status: int, code: str, message: str) -> func.HttpResponse:
        return func.HttpResponse(
            json.dumps({"error": code, "message": message}),
            status_code=status,
            mimetype="application/json",
        )
