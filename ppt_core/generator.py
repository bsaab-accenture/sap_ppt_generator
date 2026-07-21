"""
ppt_core/generator.py
─────────────────────
Agenda PPT generation using python-pptx.

Template: Agenda_BM20260116.pptx  (2 slides)
  Slide 1 – Cover  : meeting title, formatted date, attendees block
  Slide 2 – Agenda : header bar | Opening row | up to 2 numbered topic rows
                     | Board-only closing section (kept from template)

Input JSON schema
─────────────────
{
  "meeting": {
    "meetingTitle"         : str,           # e.g. "Board Meeting July"
    "meetingDate"          : "YYYY-MM-DD",  # e.g. "2026-07-30"
    "meetingTime"          : "HH:MM",       # 24-h start, e.g. "13:00"
    "attendees"            : "a@x;b@x",     # semicolon-separated
    "openingDurationMinutes": int
  },
  "topics": [
    {
      "topicTitle"     : str,
      "sequence"       : int,
      "startTime"      : "HH:MM AM/PM",   # e.g. "01:10 PM"
      "durationMinutes": int,
      "boardMember"    : str,
      "xBoard"         : str,
      "lead"           : str,
      "guest"          : str,
      "status"         : str
    }
  ]
}

Capacity: 2 numbered topic rows per slide.
Topics beyond 2 are logged and dropped (extend with multi-slide support later).
"""

from __future__ import annotations

import logging
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Config (kept compatible with function_app.py interface)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class GeneratorConfig:
    scripts_dir: Path
    workdir: Path
    max_bullets: int = 4


# ─────────────────────────────────────────────────────────────────────────────
# Shape ID map — Agenda_BM20260116.pptx
# (IDs are the cNvPr id= attribute, scoped per slide)
# ─────────────────────────────────────────────────────────────────────────────

# Slide 1 – cover
_S1_TITLE_ID     = 2   # "Title 23"  : "SAP Executive Board Meeting \n <date>"
_S1_ATTENDEES_ID = 3   # "TextBox 5" : attendees block

# Slide 2 – header bar
_S2_HEADER_ID = 4      # "Titel 4"   : "SAP Executive Board Meeting | <date>"

# Slide 2 – Opening row
_S2_OPEN_TIME_ID  = 37  # start time
_S2_OPEN_DUR_ID   = 39  # duration
_S2_OPEN_BOARD_ID = 30  # board member (first name)

# Slide 2 – Topic row 1
_S2_T1_SEQ_ID    = 38
_S2_T1_TIME_ID   = 6
_S2_T1_DUR_ID    = 14
_S2_T1_TOPIC_ID  = 34
_S2_T1_BOARD_ID  = 21
_S2_T1_XBOARD_ID = 35
_S2_T1_LEAD_ID   = 27

# Slide 2 – Topic row 2
_S2_T2_SEQ_ID    = 45
_S2_T2_TIME_ID   = 31
_S2_T2_DUR_ID    = 40
_S2_T2_TOPIC_ID  = 43
_S2_T2_BOARD_ID  = 44
_S2_T2_XBOARD_ID = 48
_S2_T2_LEAD_ID   = 11

# All shape IDs that belong to each topic row (used for clearing unused rows)
_S2_T1_ALL_IDS = [_S2_T1_SEQ_ID, _S2_T1_TIME_ID, _S2_T1_DUR_ID,
                  _S2_T1_TOPIC_ID, _S2_T1_BOARD_ID, _S2_T1_XBOARD_ID, _S2_T1_LEAD_ID]
_S2_T2_ALL_IDS = [_S2_T2_SEQ_ID, _S2_T2_TIME_ID, _S2_T2_DUR_ID,
                  _S2_T2_TOPIC_ID, _S2_T2_BOARD_ID, _S2_T2_XBOARD_ID, _S2_T2_LEAD_ID]


# ─────────────────────────────────────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────────────────────────────────────

def generate_pptx(
    template_bytes: bytes,
    content_json: dict,
    cfg: GeneratorConfig,
) -> bytes:
    """
    Generate an Agenda PPTX from the template and input JSON.

    Returns raw .pptx bytes.
    """
    meeting = content_json.get("meeting", {})
    topics  = sorted(
        content_json.get("topics", []),
        key=lambda t: int(t.get("sequence", 0)),
    )

    if not meeting:
        raise ValueError("'meeting' key is required in the input JSON.")

    if len(topics) > 2:
        log.warning(
            "Template supports 2 topic rows; %d topics provided — "
            "only the first 2 will be rendered.",
            len(topics),
        )

    prs = Presentation(BytesIO(template_bytes))

    _fill_cover(prs.slides[0], meeting)
    _fill_agenda(prs.slides[1], meeting, topics[:2])

    out = BytesIO()
    prs.save(out)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Slide 1 – cover
# ─────────────────────────────────────────────────────────────────────────────

def _fill_cover(slide, meeting: dict) -> None:
    title     = meeting.get("meetingTitle", "")
    date_str  = meeting.get("meetingDate", "")
    attendees = meeting.get("attendees", "")

    date_fmt = _fmt_date(date_str) if date_str else ""

    shape_map = _shape_id_map(slide)

    # Title + date text box (id=2): "Meeting Title\n<date>"
    if _S1_TITLE_ID in shape_map:
        _set_two_line_text(shape_map[_S1_TITLE_ID], title, date_fmt)
        log.info("Cover title set: %r / %r", title, date_fmt)

    # Attendees block (id=3)
    if _S1_ATTENDEES_ID in shape_map:
        attendee_list = [a.strip() for a in attendees.split(";") if a.strip()]
        _set_attendees_block(shape_map[_S1_ATTENDEES_ID], attendee_list)
        log.info("Cover attendees set: %d entries", len(attendee_list))


def _set_two_line_text(shape, line1: str, line2: str) -> None:
    """
    Replace shape content with two lines separated by a line break.
    Preserves the formatting of the first existing run on each line.
    """
    tf = shape.text_frame
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))
    if not paras:
        return

    # Use first paragraph for both lines (with <a:br> between them)
    first_para = paras[0]

    # Snapshot run properties before clearing
    runs = first_para.findall(qn("a:r"))
    first_rpr = deepcopy(runs[0].find(qn("a:rPr"))) if runs else None

    # Find the <a:br> element to get its rPr (for the line break)
    br_elem = first_para.find(qn("a:br"))
    br_rpr  = deepcopy(br_elem.find(qn("a:rPr"))) if br_elem is not None else None

    # Clear first paragraph content
    for child in list(first_para):
        tag = child.tag
        if tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            first_para.remove(child)

    # Remove extra paragraphs
    for para in paras[1:]:
        txBody.remove(para)

    # Build: <a:r>line1</a:r> <a:br/> <a:r>line2</a:r>
    _append_run(first_para, line1, first_rpr)

    br = etree.SubElement(first_para, qn("a:br"))
    if br_rpr is not None:
        br.append(br_rpr)

    _append_run(first_para, line2, first_rpr)


def _set_attendees_block(shape, attendee_list: list[str]) -> None:
    """
    Replace the attendees shape with one paragraph per attendee.
    The first paragraph is bold (group label), the rest are the names.
    """
    if not attendee_list:
        _set_single_run_text(shape, "")
        return

    tf = shape.text_frame
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))

    # Collect run property templates from existing paragraphs
    bold_rpr   = _snapshot_rpr(paras[0]) if paras else None
    normal_rpr = _snapshot_rpr(paras[1]) if len(paras) > 1 else bold_rpr

    # Clear all paragraphs
    for para in list(paras):
        txBody.remove(para)

    # First paragraph: bold header showing meeting attendees label
    _append_paragraph(txBody, "Attendees:", bold_rpr, paras[0] if paras else None)

    # Subsequent paragraphs: one per attendee (show readable name from email)
    for email in attendee_list:
        name = _email_to_name(email)
        _append_paragraph(txBody, name, normal_rpr, paras[1] if len(paras) > 1 else None)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 2 – agenda table
# ─────────────────────────────────────────────────────────────────────────────

def _fill_agenda(slide, meeting: dict, topics: list[dict]) -> None:
    shape_map = _shape_id_map(slide)

    # ── Header bar ────────────────────────────────────────────────────────────
    if _S2_HEADER_ID in shape_map:
        title    = meeting.get("meetingTitle", "")
        date_str = meeting.get("meetingDate", "")
        date_fmt = _fmt_date(date_str) if date_str else ""
        _set_single_run_text(shape_map[_S2_HEADER_ID], f"{title}  |  {date_fmt}")
        log.info("Agenda header set: %r", f"{title} | {date_fmt}")

    # ── Opening row ───────────────────────────────────────────────────────────
    opening_start = _fmt_time_24h(meeting.get("meetingTime", "09:00"))
    opening_dur   = meeting.get("openingDurationMinutes", 5)

    _set_single_run_text_by_id(shape_map, _S2_OPEN_TIME_ID, opening_start)
    _set_single_run_text_by_id(shape_map, _S2_OPEN_DUR_ID,  f"{opening_dur} min")
    # Opening board member: first name from first attendee email
    first_attendee = meeting.get("attendees", "").split(";")[0].strip()
    _set_single_run_text_by_id(
        shape_map, _S2_OPEN_BOARD_ID,
        _email_to_first_name(first_attendee) if first_attendee else "",
    )
    log.info("Opening row: time=%s dur=%s min", opening_start, opening_dur)

    # ── Topic rows ────────────────────────────────────────────────────────────
    _fill_topic_row(shape_map, topics[0] if len(topics) > 0 else None, slot=1)
    _fill_topic_row(shape_map, topics[1] if len(topics) > 1 else None, slot=2)


def _fill_topic_row(shape_map: dict, topic: dict | None, slot: int) -> None:
    """
    Populate (slot=1 or slot=2) with a topic dict, or clear it if topic=None.
    """
    if slot == 1:
        ids = dict(
            seq=_S2_T1_SEQ_ID, time=_S2_T1_TIME_ID, dur=_S2_T1_DUR_ID,
            topic=_S2_T1_TOPIC_ID, board=_S2_T1_BOARD_ID,
            xboard=_S2_T1_XBOARD_ID, lead=_S2_T1_LEAD_ID,
        )
    else:
        ids = dict(
            seq=_S2_T2_SEQ_ID, time=_S2_T2_TIME_ID, dur=_S2_T2_DUR_ID,
            topic=_S2_T2_TOPIC_ID, board=_S2_T2_BOARD_ID,
            xboard=_S2_T2_XBOARD_ID, lead=_S2_T2_LEAD_ID,
        )

    if topic is None:
        # Clear all shapes in this slot
        for shape_id in ids.values():
            _set_single_run_text_by_id(shape_map, shape_id, "")
        return

    seq       = str(topic.get("sequence", slot))
    start     = _normalise_time(topic.get("startTime", ""))
    dur       = f"{topic.get('durationMinutes', '')} min"
    title     = topic.get("topicTitle", "")
    board     = topic.get("boardMember", "")
    xboard    = topic.get("xBoard", "")
    lead      = topic.get("lead", "")
    guest     = topic.get("guest", "")
    lead_str  = ", ".join(filter(None, [lead, guest]))

    _set_single_run_text_by_id(shape_map, ids["seq"],    seq)
    _set_single_run_text_by_id(shape_map, ids["time"],   start)
    _set_single_run_text_by_id(shape_map, ids["dur"],    dur)
    _set_topic_text(shape_map, ids["topic"], title)
    _set_single_run_text_by_id(shape_map, ids["board"],  board)
    _set_single_run_text_by_id(shape_map, ids["xboard"], xboard)
    _set_single_run_text_by_id(shape_map, ids["lead"],   lead_str)

    log.info("Topic row %d: [%s] %s — %s, board=%r", slot, seq, title, dur, board)


def _set_topic_text(shape_map: dict, shape_id: int, title: str) -> None:
    """
    Set the topic title in the topic text box.
    The template uses a bold gradient heading in the first paragraph;
    we keep that run style and just replace its text.
    """
    shape = shape_map.get(shape_id)
    if shape is None:
        return

    tf = shape.text_frame
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))
    if not paras:
        return

    first_para = paras[0]
    runs = first_para.findall(qn("a:r"))
    first_rpr = deepcopy(runs[0].find(qn("a:rPr"))) if runs else None

    # Clear first paragraph, keep format
    for child in list(first_para):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            first_para.remove(child)

    # Remove extra paragraphs (sub-bullets from template)
    for para in paras[1:]:
        txBody.remove(para)

    # Set title as single run
    _append_run(first_para, title, first_rpr)


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────────────────────────────────────

def _shape_id_map(slide) -> dict[int, object]:
    """Return {shape_id: shape} for all shapes on the slide."""
    return {shape.shape_id: shape for shape in slide.shapes}


def _set_single_run_text(shape, text: str) -> None:
    """Replace shape text with a single run, preserving first run's rPr."""
    tf = shape.text_frame
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))
    if not paras:
        return

    first_para = paras[0]
    runs = first_para.findall(qn("a:r"))
    first_rpr = deepcopy(runs[0].find(qn("a:rPr"))) if runs else None

    # Clear
    for child in list(first_para):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            first_para.remove(child)
    for para in paras[1:]:
        txBody.remove(para)

    _append_run(first_para, text, first_rpr)


def _set_single_run_text_by_id(shape_map: dict, shape_id: int, text: str) -> None:
    shape = shape_map.get(shape_id)
    if shape is not None:
        _set_single_run_text(shape, text)


def _append_run(para_elem, text: str, rpr=None) -> None:
    """Append a new <a:r> with optional <a:rPr> to a paragraph element."""
    r = etree.SubElement(para_elem, qn("a:r"))
    if rpr is not None:
        r.append(deepcopy(rpr))
    t = etree.SubElement(r, qn("a:t"))
    t.text = text


def _append_paragraph(txBody_elem, text: str, rpr=None, template_para=None) -> None:
    """Append a new <a:p> with one run to a txBody element."""
    new_p = etree.SubElement(txBody_elem, qn("a:p"))

    # Copy paragraph-level properties from the template paragraph if available
    if template_para is not None:
        tmpl_ppr = template_para.find(qn("a:pPr"))
        if tmpl_ppr is not None:
            new_p.append(deepcopy(tmpl_ppr))

    _append_run(new_p, text, rpr)


def _snapshot_rpr(para_elem):
    """Return a deep copy of the first run's <a:rPr> in a paragraph element."""
    run = para_elem.find(qn("a:r"))
    if run is None:
        return None
    rpr = run.find(qn("a:rPr"))
    return deepcopy(rpr) if rpr is not None else None


# ─────────────────────────────────────────────────────────────────────────────
# Formatting helpers
# ─────────────────────────────────────────────────────────────────────────────

def _fmt_date(date_str: str) -> str:
    """'2026-07-30' → 'July 30, 2026'"""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%B %-d, %Y")       # Linux/Mac
    except ValueError:
        pass
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%B %d, %Y").replace(" 0", " ")  # Windows fallback
    except ValueError:
        return date_str


def _fmt_time_24h(time_str: str) -> str:
    """'13:00' → '1:00 PM'"""
    try:
        dt = datetime.strptime(time_str, "%H:%M")
        raw = dt.strftime("%I:%M %p")
        return raw.lstrip("0") or raw
    except ValueError:
        return time_str


def _normalise_time(time_str: str) -> str:
    """'01:10 PM' → '1:10 PM' (strip leading zero)."""
    if not time_str:
        return ""
    try:
        dt = datetime.strptime(time_str, "%I:%M %p")
        raw = dt.strftime("%I:%M %p")
        return raw.lstrip("0") or raw
    except ValueError:
        return time_str


def _email_to_name(email: str) -> str:
    """'christian.klein@sap.com' → 'Christian Klein'"""
    local = email.split("@")[0] if "@" in email else email
    return " ".join(part.capitalize() for part in local.replace(".", " ").split())


def _email_to_first_name(email: str) -> str:
    """'christian.klein@sap.com' → 'Christian'"""
    local = email.split("@")[0] if "@" in email else email
    first = local.split(".")[0]
    return first.capitalize()
