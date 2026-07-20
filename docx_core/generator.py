"""docx_core/generator.py
Core Word DOCX generation engine — zero Azure dependencies.

Consumes an agendadocx_input.json payload and a .docx template (for styles /
page setup / header-footer) and produces a Preliminary Minutes document that
mirrors the two-page layout of Preliminary_Minutes_BM20260219.docx:

  Page 1  ─  Title block  +  Participants / Topics / Minutes summary table
  Page 2  ─  3-column minutes detail table (Topic | Lead | Due)
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Optional

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── Brand / style constants ─────────────────────────────────────────────────

FONT = "Segoe UI"
SAP_BLUE = RGBColor(0x44, 0x72, 0xC4)   # #4472C4 — SAP brand accent
BLACK = RGBColor(0x00, 0x00, 0x00)
HDR_FILL = "D9E2F3"                      # Light SAP blue table-header fill

PT_TITLE = 16    # Meeting title / "Preliminary Minutes" heading
PT_BODY = 10     # General body text
PT_HDR_FTR = 8   # Running header / footer text

# Exact column widths in twips (1 twip = 1/1440 inch), calibrated from template
# Page 1 two-column summary table
P1_COL1 = 2268   # label column  (~4.00 cm)
P1_COL2 = 7052   # content column (~12.44 cm)

# Page 2 three-column minutes table
P2_COL1 = 7225   # topic content  (~12.74 cm)
P2_COL2 = 1750   # lead           (~3.09 cm)
P2_COL3 = 1226   # due date       (~2.16 cm)


@dataclass
class DocxConfig:
    """Reserved for future generator tuning knobs."""
    pass


# ─── Schema validation ────────────────────────────────────────────────────────

def validate_input(payload: dict) -> dict:
    """Return the validated ``body`` dict from the raw JSON payload.

    Expected top-level shape::

        {
            "body": {
                "Meeting_Title": "...",
                "Date": "...",
                "Topics_Discussed": [ { "Topic": "...", ... }, ... ]
            }
        }

    Raises:
        ValueError: if any required key is missing or has the wrong type.
    """
    if not isinstance(payload, dict) or "body" not in payload:
        got = list(payload.keys()) if isinstance(payload, dict) else type(payload).__name__
        raise ValueError(
            f"Input JSON must have a top-level 'body' key.  Got: {got}"
        )
    body = payload["body"]
    for key in ("Meeting_Title", "Date", "Topics_Discussed"):
        if key not in body:
            raise ValueError(f"body is missing required field '{key}'")
    if not isinstance(body["Topics_Discussed"], list) or not body["Topics_Discussed"]:
        raise ValueError("body.Topics_Discussed must be a non-empty list")
    return body


# ─── Low-level XML helpers ────────────────────────────────────────────────────

def _shade_cell(cell, hex_fill: str) -> None:
    """Apply a solid background colour to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _set_cell_width(cell, twips: int) -> None:
    """Explicitly set a cell's preferred width (in twips / dxa units)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcW")):
        tcPr.remove(existing)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.insert(0, tcW)


def _set_tbl_width(table, twips: int) -> None:
    """Set the preferred total width of a table."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for existing in tblPr.findall(qn("w:tblW")):
        tblPr.remove(existing)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(twips))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)


# ─── Paragraph helpers ────────────────────────────────────────────────────────

def _fmt_run(run, bold: bool = False, size: int = PT_BODY,
             color: RGBColor = BLACK) -> None:
    """Apply run-level character formatting."""
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color


def _write_para(
    container,
    text: str = "",
    bold: bool = False,
    size: int = PT_BODY,
    color: RGBColor = BLACK,
    space_before: float = 0.0,
    space_after: float = 0.0,
) -> None:
    """Write a paragraph into a Document or table-cell container.

    Reuses the first paragraph of a freshly-created cell (which is always
    empty) so we never leave a spurious blank line at the top of a cell.
    Falls back to ``add_paragraph()`` everywhere else.
    """
    reuse = (
        hasattr(container, "paragraphs")
        and container.paragraphs
        and container.paragraphs[0].text == ""
    )
    p = container.paragraphs[0] if reuse else container.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        _fmt_run(p.add_run(text), bold=bold, size=size, color=color)


def _add_para(
    cell,
    text: str = "",
    bold: bool = False,
    size: int = PT_BODY,
    color: RGBColor = BLACK,
    space_before: float = 0.0,
) -> None:
    """Always append a new paragraph to a table cell."""
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(0)
    if text:
        _fmt_run(p.add_run(text), bold=bold, size=size, color=color)


# ─── Header updater ───────────────────────────────────────────────────────────

def _update_running_header(doc: Document, meeting_title: str, date_str: str) -> None:
    """Replace the running-header text (left side) with the current meeting info.

    The template header layout is:  «Meeting Title   Date  <tab>  Preliminary Minutes»
    We rebuild this structure with correct meeting data.
    """
    new_left = f"{meeting_title}   {date_str}"
    for section in doc.sections:
        hdr = section.header
        if hdr.is_linked_to_previous:
            continue
        for para in hdr.paragraphs:
            if not para.text.strip():
                continue
            full = "".join(r.text for r in para.runs)
            if not full.strip():
                continue
            # Remove all existing runs from this paragraph
            for run in list(para.runs):
                run._element.getparent().remove(run._element)
            # Rebuild: left text + tab stop + right label
            r1 = para.add_run(new_left)
            r1.font.name = FONT
            r1.font.size = Pt(PT_HDR_FTR)
            r1.font.color.rgb = BLACK
            para.add_run("\t").font.size = Pt(PT_HDR_FTR)
            r3 = para.add_run("Preliminary Minutes")
            r3.font.name = FONT
            r3.font.size = Pt(PT_HDR_FTR)
            r3.font.color.rgb = BLACK


# ─── Page 1: cover / summary ─────────────────────────────────────────────────

def _build_page1(doc: Document, body: dict) -> None:
    """Build the title block and participants/topics/minutes summary table."""
    meeting_title: str = body["Meeting_Title"]
    date_str: str = body["Date"]
    time_str: str = body.get("Time", "")
    attendees: list = body.get("Attendees", [])
    topics: list = body["Topics_Discussed"]

    # ── Title line ────────────────────────────────────────────────────────────
    date_display = f"{date_str}  {time_str}".strip() if time_str else date_str
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(0)
    _fmt_run(p_title.add_run(f"{meeting_title}   {date_display}"), size=PT_TITLE)

    # ── Subtitle ──────────────────────────────────────────────────────────────
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(8)
    _fmt_run(p_sub.add_run("Preliminary Minutes"), bold=True, size=PT_TITLE)

    # ── Two-column summary table ──────────────────────────────────────────────
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"
    _set_tbl_width(tbl, P1_COL1 + P1_COL2)

    def _row(label: str) -> tuple:
        """Add a labeled row; return (label_cell, content_cell)."""
        r = tbl.add_row()
        lc, cc = r.cells[0], r.cells[1]
        _set_cell_width(lc, P1_COL1)
        _set_cell_width(cc, P1_COL2)
        _write_para(lc, label, size=PT_BODY)
        return lc, cc

    # ── Row: Participants (header label only) ─────────────────────────────────
    _row("Participants")

    # ── Row: Executive Board attendees ────────────────────────────────────────
    _, cc = _row("Executive Board")
    if attendees:
        _write_para(cc, attendees[0])
        for name in attendees[1:]:
            _add_para(cc, name)

    # ── Row: Topics | Lead, Guests ────────────────────────────────────────────
    _, cc = _row("Topics | Lead, Guests")
    first = True
    for i, topic in enumerate(topics, 1):
        name = topic.get("Topic", "")
        duration = topic.get("Duration", "")
        board_att = topic.get("Board_Attendees", [])

        header_line = f"{i}. {name}"
        if first:
            _write_para(cc, header_line, bold=True)
            first = False
        else:
            _add_para(cc, header_line, bold=True, space_before=3)

        parts = ([duration] if duration else []) + board_att
        if parts:
            _add_para(cc, " | ".join(parts))

    # ── Row: Minutes (per-topic summary) ──────────────────────────────────────
    _, cc = _row("Minutes")
    first = True
    for i, topic in enumerate(topics, 1):
        name = topic.get("Topic", "")
        minutes_text = topic.get("Topic_Minutes", "")

        header_line = f"{i}. {name}"
        if first:
            _write_para(cc, header_line, bold=True)
            first = False
        else:
            _add_para(cc, header_line, bold=True, space_before=4)

        if minutes_text:
            _add_para(cc, minutes_text)

    # ── Row: Distribution List ────────────────────────────────────────────────
    _, cc = _row("Distribution List")
    _write_para(cc, "Executive Board")
    _add_para(cc, "Guests (extracts only)")


# ─── Page 2: minutes detail ───────────────────────────────────────────────────

def _build_page2(doc: Document, body: dict) -> None:
    """Build the 3-column minutes detail table."""
    topics: list = body["Topics_Discussed"]

    tbl = doc.add_table(rows=0, cols=3)
    tbl.style = "Table Grid"
    _set_tbl_width(tbl, P2_COL1 + P2_COL2 + P2_COL3)

    # ── Header row ────────────────────────────────────────────────────────────
    hdr = tbl.add_row()
    col_widths = (P2_COL1, P2_COL2, P2_COL3)
    hdr_labels = (
        "Topic, context, feedback, decisions, and action items",
        "Lead",
        "Due",
    )
    for cell, w, label in zip(hdr.cells, col_widths, hdr_labels):
        _set_cell_width(cell, w)
        _shade_cell(cell, HDR_FILL)
        _write_para(cell, label, bold=True)

    # ── Data rows — one per topic ─────────────────────────────────────────────
    for i, topic in enumerate(topics, 1):
        name: str = topic.get("Topic", "")
        lead: str = topic.get("Topic_Lead", "")
        description: str = topic.get("Topic_Description", "")
        minutes_text: str = topic.get("Topic_Minutes", "")
        actions: list = topic.get("Actions", [])

        row = tbl.add_row()
        for cell, w in zip(row.cells, col_widths):
            _set_cell_width(cell, w)

        mc = row.cells[0]   # main content cell

        # Bold topic heading
        _write_para(mc, f"{i}. {name}", bold=True)

        # Topic description
        if description:
            _add_para(mc, description)

        # Spacer + blue context subheading
        mc.add_paragraph()
        p_ctx = mc.add_paragraph()
        p_ctx.paragraph_format.space_after = Pt(0)
        _fmt_run(
            p_ctx.add_run(
                "Context, Board feedback, decisions & resulting action items:"
            ),
            color=SAP_BLUE,
        )

        # Meeting minutes text
        if minutes_text:
            _add_para(mc, minutes_text)

        # Action items
        for action in actions:
            action_text = action.get("Action", "")
            owner = action.get("Owner", "")
            if action_text:
                _add_para(mc, f"▶  Action: {action_text}", bold=True, space_before=4)
            if owner:
                _add_para(mc, f"    Owner: {owner}")

        # Lead cell
        _write_para(row.cells[1], lead)

        # Due date — use first action's due date, or em-dash if none
        due = actions[0].get("Due_Date", "–") if actions else "–"
        _write_para(row.cells[2], due)


# ─── Public entry point ───────────────────────────────────────────────────────

def generate_docx(
    template_bytes: bytes,
    content_json: dict,
    cfg: Optional[DocxConfig] = None,
) -> bytes:
    """Generate a Preliminary Minutes Word document from JSON meeting data.

    Opens the DOCX template (preserving its styles, theme, and header/footer),
    clears the body, injects data-driven content, then serialises to bytes.

    Args:
        template_bytes: Raw bytes of the ``.docx`` template file.  Provides SAP
                        brand styles, A4 page setup, and running header/footer.
        content_json:   Parsed JSON with an ``agendadocx_input.json`` shape —
                        must have a top-level ``"body"`` key.
        cfg:            Optional :class:`DocxConfig` (reserved for future knobs).

    Returns:
        Raw bytes of the generated ``.docx`` document.

    Raises:
        ValueError: If ``content_json`` does not match the expected schema.
    """
    if cfg is None:
        cfg = DocxConfig()

    body = validate_input(content_json)

    # ── Open template — inherits styles, theme, page setup, header/footer ────
    doc = Document(io.BytesIO(template_bytes))

    # ── Clear body content; preserve the final <w:sectPr> (A4 margins etc.) ──
    body_el = doc.element.body
    sect_pr = body_el.find(qn("w:sectPr"))   # direct-child sectPr only
    for child in list(body_el):
        body_el.remove(child)

    # Restore sectPr immediately: python-docx's add_table / add_paragraph insert
    # new elements *before* sectPr, so it must be present before we build content.
    if sect_pr is not None:
        body_el.append(sect_pr)

    # ── Update running header with correct meeting info ───────────────────────
    _update_running_header(doc, body["Meeting_Title"], body["Date"])

    # ── Page 1: title + summary table ────────────────────────────────────────
    _build_page1(doc, body)

    # ── Page break ────────────────────────────────────────────────────────────
    doc.add_page_break()

    # ── Page 2: minutes detail table ─────────────────────────────────────────
    _build_page2(doc, body)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
