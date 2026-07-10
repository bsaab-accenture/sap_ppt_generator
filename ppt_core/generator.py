"""
ppt_core/generator.py
─────────────────────
Pure-Python PPT generation logic.

Zero Azure dependencies — fully unit-testable in isolation.
The Azure Function (function_app.py) handles all I/O; this module
handles only PPTX construction.

Public API
----------
    generate_pptx(
        template_bytes : bytes,           # raw .pptx file content
        content_json   : dict,            # parsed JSON payload
        scripts_dir    : Path,            # path to pptx skill scripts
        workdir        : Path,            # temp directory for unpack/pack
    ) -> bytes                            # raw .pptx file content
"""

from __future__ import annotations

import json
import logging
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import NamedTuple

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Constants — calibrated to SAP Board Meeting Governance template
# ─────────────────────────────────────────────────────────────────────────────

MAX_BULLETS: int = 4


@dataclass(frozen=True)
class TemplateGeometry:
    """
    EMU coordinates extracted from the SAP Board Meeting Governance template.
    914400 EMU = 1 inch.

    Override via GeneratorConfig if you switch to a different template.
    """
    # y-position of each numbered circle row (circles 01–04)
    circle_y: tuple[int, ...] = (1942204, 3029028, 4115852, 5202677)

    # Bullet text box: starts at x=7315937, same y as its circle row
    bullet_box_x: int = 7315937
    bullet_box_w: int = 4175883
    bullet_box_h: int = 950000      # ~1 inch; fits ~2 wrapped lines at 18pt

    # Font size in hundredths of a point
    bullet_font_sz: int = 1800


@dataclass(frozen=True)
class TemplateSlotMap:
    """
    Visual slide-order positions in the SAP template (1-indexed).
    """
    title_slide_idx: int = 1
    content_slide_range: tuple[int, int] = (4, 8)   # inclusive
    closing_slide_idx: int = 9


@dataclass
class GeneratorConfig:
    """
    All tuneable parameters for one generation run.
    Constructed by function_app.py from env-vars and request context.
    """
    scripts_dir: Path
    workdir: Path
    max_bullets: int = MAX_BULLETS
    geometry: TemplateGeometry = field(default_factory=TemplateGeometry)
    slot_map: TemplateSlotMap = field(default_factory=TemplateSlotMap)


# ─────────────────────────────────────────────────────────────────────────────
# Data model
# ─────────────────────────────────────────────────────────────────────────────

class ContentSlide(NamedTuple):
    title: str
    bullets: list[str]
    speaker_notes: str


# ─────────────────────────────────────────────────────────────────────────────
# JSON validation
# ─────────────────────────────────────────────────────────────────────────────

def validate_content(payload: dict) -> list[dict]:
    """
    Validate the parsed JSON payload and return the 'slides' list.

    Expected shape
    --------------
    {
      "slides": [
        {
          "title":         "string (required)",
          "bullets":       ["string", ...],   // at least 1 required
          "speaker_notes": "string (optional)"
        }
      ]
    }

    Raises
    ------
    ValueError  on any schema violation, with a human-readable message.
    """
    if not isinstance(payload, dict) or "slides" not in payload:
        raise ValueError("Payload must be a JSON object with a top-level 'slides' array.")

    slides = payload["slides"]
    if not isinstance(slides, list) or not slides:
        raise ValueError("'slides' must be a non-empty array.")

    for i, slide in enumerate(slides):
        if not isinstance(slide.get("title"), str) or not slide["title"].strip():
            raise ValueError(f"slides[{i}].title must be a non-empty string.")
        bullets = slide.get("bullets")
        if not isinstance(bullets, list) or not bullets:
            raise ValueError(f"slides[{i}].bullets must be a non-empty array.")
        for j, b in enumerate(bullets):
            if not isinstance(b, str) or not b.strip():
                raise ValueError(f"slides[{i}].bullets[{j}] must be a non-empty string.")

    return slides


# ─────────────────────────────────────────────────────────────────────────────
# Slide splitting
# ─────────────────────────────────────────────────────────────────────────────

def split_slide(slide: dict, max_bullets: int) -> list[ContentSlide]:
    """
    Convert one JSON slide dict into one or more ContentSlide objects.

    Rule
    ----
    If len(bullets) <= max_bullets → single slide, title unchanged.
    Otherwise                      → chunk into "Title - Part 1", "Part 2", …
                                     speaker_notes only on Part 1.
    """
    title   = slide["title"].strip()
    bullets = [b.strip() for b in slide["bullets"]]
    notes   = slide.get("speaker_notes", "").strip()

    if len(bullets) <= max_bullets:
        return [ContentSlide(title=title, bullets=bullets, speaker_notes=notes)]

    chunks = [bullets[i: i + max_bullets] for i in range(0, len(bullets), max_bullets)]
    return [
        ContentSlide(
            title=f"{title} - Part {n}",
            bullets=chunk,
            speaker_notes=notes if n == 1 else "",
        )
        for n, chunk in enumerate(chunks, start=1)
    ]


def build_content_slides(raw_slides: list[dict], max_bullets: int) -> list[ContentSlide]:
    """Flatten all JSON slides into the final ordered ContentSlide list."""
    output: list[ContentSlide] = []
    for slide in raw_slides:
        output.extend(split_slide(slide, max_bullets))
    return output


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────────────────────────────────────

def xml_escape(text: str) -> str:
    """Escape the five XML special characters."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def make_bullet_sp(
    text: str,
    y_pos: int,
    sp_id: int,
    row_index: int,
    geo: TemplateGeometry,
) -> str:
    """
    Build a standalone <p:sp> text box for one bullet point, vertically
    aligned with its numbered circle row.

    Parameters
    ----------
    text      : bullet text (will be XML-escaped)
    y_pos     : EMU y-offset (from TemplateGeometry.circle_y)
    sp_id     : unique integer for <p:cNvPr id="...">
    row_index : 1-based row number (used only for the shape name label)
    geo       : TemplateGeometry constants
    """
    safe = xml_escape(text)
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="BulletRow{row_index}"/>'
        f'<p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm>'
        f'<a:off x="{geo.bullet_box_x}" y="{y_pos}"/>'
        f'<a:ext cx="{geo.bullet_box_w}" cy="{geo.bullet_box_h}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr vert="horz"'
        f' lIns="108000" tIns="36000" rIns="108000" bIns="36000" rtlCol="0">'
        f'<a:normAutofit/>'
        f'</a:bodyPr>'
        f'<a:lstStyle/>'
        f'<a:p>'
        f'<a:pPr algn="l"/>'
        f'<a:r>'
        f'<a:rPr lang="en-GB" sz="{geo.bullet_font_sz}" dirty="0"/>'
        f'<a:t>{safe}</a:t>'
        f'</a:r>'
        f'</a:p>'
        f'</p:txBody>'
        f'</p:sp>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# Slide XML mutators
# ─────────────────────────────────────────────────────────────────────────────

def inject_title_slide(xml: str, meeting_title: str, date_str: str) -> str:
    """
    Replace template placeholders on the cover slide.

    Placeholders
    ------------
    {meetingTitle}   – split across 3 <a:r> runs in the template
    {presenter}, SAP – replaced with "Board Meeting"
    {date}           – replaced with extracted date string
    """
    safe_title = xml_escape(meeting_title)

    # {meetingTitle}: three-run pattern; <a:rPr> may be self-closing OR have children
    xml = re.sub(
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>\{</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>meetingTitle</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>\}</a:t>\s*</a:r>",
        (
            f'<a:r>'
            f'<a:rPr lang="en-US" b="1" sz="3600">'
            f'<a:cs typeface="Arial"/>'
            f'</a:rPr>'
            f'<a:t>{safe_title}</a:t>'
            f'</a:r>'
        ),
        xml,
        flags=re.DOTALL,
    )
    xml = xml.replace("{presenter}, SAP", "Board Meeting")
    xml = re.sub(r"<a:t>\s*\{date\}</a:t>", f"<a:t>{xml_escape(date_str)}</a:t>", xml)
    return xml


def inject_section_title(xml: str, title: str) -> str:
    """
    Replace the {sectionTitle} placeholder on a content slide.

    The placeholder is split across three <a:r> runs: " {" | "sectionTitle" | "}"
    Handles both self-closing and block-form <a:rPr>.
    """
    safe_title = xml_escape(title)
    replacement = f'<a:r><a:rPr lang="en-GB" b="1"/><a:t>{safe_title}</a:t></a:r>'

    # Pattern A: block-form rPr (has child elements)
    xml = re.sub(
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>\s* \{</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>sectionTitle</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^>]*>.*?</a:rPr>\s*<a:t>}</a:t>\s*</a:r>",
        replacement,
        xml,
        flags=re.DOTALL,
    )
    # Pattern B: self-closing rPr
    xml = re.sub(
        r"<a:r>\s*<a:rPr[^/]*\/>\s*<a:t>\s* \{</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^/]*\/>\s*<a:t>sectionTitle</a:t>\s*</a:r>\s*"
        r"<a:r>\s*<a:rPr[^/]*\/>\s*<a:t>}</a:t>\s*</a:r>",
        replacement,
        xml,
        flags=re.DOTALL,
    )
    return xml


def inject_bullets(xml: str, bullets: list[str], geo: TemplateGeometry) -> str:
    """
    Replace the monolithic {sectionPoints} text box with N individual
    text boxes, one per bullet, each aligned to its numbered circle row.

    Steps
    -----
    1. Strip the original {sectionPoints} <p:sp> element entirely.
    2. Build one <p:sp> per bullet at the correct circle-row y-position.
    3. Insert all new blocks just before </p:spTree>.
    """
    # Remove original sectionPoints shape
    xml = re.sub(
        r"<p:sp>(?:(?!</p:sp>).)*?\{sectionPoints\}(?:(?!</p:sp>).)*?</p:sp>",
        "",
        xml,
        flags=re.DOTALL,
    )

    sp_blocks = [
        make_bullet_sp(
            text=bullet,
            y_pos=geo.circle_y[idx],
            sp_id=300 + idx,
            row_index=idx + 1,
            geo=geo,
        )
        for idx, bullet in enumerate(bullets[: len(geo.circle_y)])
    ]

    xml = xml.replace(
        "</p:spTree>",
        "\n      ".join(sp_blocks) + "\n    </p:spTree>",
    )
    return xml


def hide_unused_circles(xml: str, bullet_count: int, geo: TemplateGeometry) -> str:
    """
    Remove circle shapes (rows 02–04) that have no corresponding bullet.

    Circle shapes are identified by their exact y-position in the template.
    Circle 01 (index 0) is always retained — it is always populated.

    Parameters
    ----------
    bullet_count : number of bullets on this slide (1–4)
    """
    if bullet_count >= 4:
        return xml  # all four rows occupied

    for slot in range(bullet_count + 1, 5):          # 1-based slot numbers without a bullet
        y_pos = geo.circle_y[slot - 1]
        xml = re.sub(
            r"<p:sp>(?:(?!</p:sp>).)*?"
            r'<a:off x="636[34]\d{3}" y="' + str(y_pos) + r'"/>'
            r"(?:(?!</p:sp>).)*?</p:sp>",
            "",
            xml,
            flags=re.DOTALL,
        )

    return xml


# ─────────────────────────────────────────────────────────────────────────────
# presentation.xml helpers
# ─────────────────────────────────────────────────────────────────────────────

def parse_rid_to_file(unpacked: Path) -> dict[str, str]:
    """Return {rId: slideN.xml} from the presentation relationships file."""
    rels = (unpacked / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
    return {
        m.group(1): m.group(2)
        for m in re.finditer(
            r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', rels
        )
    }


def get_ordered_rids(prs_xml: str) -> list[str]:
    """Return rIds in visual slide order from <p:sldIdLst>."""
    return re.findall(r'<p:sldId[^>]*r:id="(rId\d+)"[^/]*/?>', prs_xml)


def rebuild_sld_id_list(
    prs_xml: str,
    title_rid: str,
    content_rids: list[str],
    closing_rid: str,
) -> str:
    """
    Rewrite <p:sldIdLst> to: title | content × N | closing.
    Numeric id attributes start above the current maximum to stay unique.
    """
    all_entries = re.findall(r'<p:sldId[^/]*/>', prs_xml)
    base = max(int(re.search(r'id="(\d+)"', e).group(1)) for e in all_entries) + 100

    entries = [f'<p:sldId id="{base}" r:id="{title_rid}"/>']
    for i, rid in enumerate(content_rids, start=1):
        entries.append(f'<p:sldId id="{base + i}" r:id="{rid}"/>')
    entries.append(f'<p:sldId id="{base + len(content_rids) + 1}" r:id="{closing_rid}"/>')

    new_list = "<p:sldIdLst>\n      " + "\n      ".join(entries) + "\n    </p:sldIdLst>"
    return re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", new_list, prs_xml, flags=re.DOTALL)


def clone_content_slides(
    needed: int,
    have: int,
    unpacked: Path,
    scripts_dir: Path,
    source_file: str,
    rid_to_file: dict[str, str],
    content_rids: list[str],
) -> list[str]:
    """
    Clone content slides until `len(content_rids) == needed`.

    Uses add_slide.py from the skills scripts directory, which correctly
    handles notes stubs, Content_Types.xml, and relationship IDs.

    Returns the updated rId list trimmed to `needed` entries.
    """
    for _ in range(needed - have):
        result = subprocess.run(
            [sys.executable, str(scripts_dir / "add_slide.py"), str(unpacked), source_file],
            capture_output=True,
            text=True,
            check=True,
        )
        m = re.search(r'r:id="(rId\d+)"', result.stdout)
        if not m:
            raise RuntimeError(
                f"add_slide.py produced unexpected output:\n{result.stdout}"
            )
        new_rid = m.group(1)
        content_rids.append(new_rid)

        # Refresh rid→file map after each clone (rels file is updated in-place)
        updated_rels = (unpacked / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
        for rel_m in re.finditer(
            r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', updated_rels
        ):
            rid_to_file[rel_m.group(1)] = rel_m.group(2)

    return content_rids[:needed]


def _run_script(scripts_dir: Path, script: str, *args: str) -> str:
    """Run a pptx skill script; raise RuntimeError on non-zero exit."""
    result = subprocess.run(
        [sys.executable, str(scripts_dir / script), *args],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"{script} failed (rc={result.returncode}):\n"
            f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
        )
    return result.stdout


# ─────────────────────────────────────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────────────────────────────────────

def generate_pptx(
    template_bytes: bytes,
    content_json: dict,
    cfg: GeneratorConfig,
) -> bytes:
    """
    Generate a populated Board Meeting PPTX.

    Parameters
    ----------
    template_bytes : raw bytes of the .pptx template file
    content_json   : parsed JSON payload (validated inside this function)
    cfg            : GeneratorConfig (scripts_dir, workdir, geometry, …)

    Returns
    -------
    bytes  – raw .pptx file content ready for upload or HTTP response

    Raises
    ------
    ValueError    on invalid JSON schema
    RuntimeError  on PPTX generation failure
    """
    # ── 1. Validate & plan ────────────────────────────────────────────────────
    raw_slides = validate_content(content_json)
    content_slides = build_content_slides(raw_slides, cfg.max_bullets)

    log.info(
        "Generation plan: %d JSON slides → %d output slides",
        len(raw_slides),
        len(content_slides),
    )
    for cs in content_slides:
        log.info("  [%d bullet(s)] %s", len(cs.bullets), cs.title)

    # ── 2. Write template bytes to workdir ────────────────────────────────────
    template_path = cfg.workdir / "template.pptx"
    template_path.write_bytes(template_bytes)

    # ── 3. Unpack ─────────────────────────────────────────────────────────────
    unpacked = cfg.workdir / "unpacked"
    _run_script(cfg.scripts_dir, "office/unpack.py", str(template_path), str(unpacked))

    # ── 4. Parse template structure ───────────────────────────────────────────
    prs_xml_path = unpacked / "ppt/presentation.xml"
    prs_xml      = prs_xml_path.read_text(encoding="utf-8")
    rid_to_file  = parse_rid_to_file(unpacked)
    ordered_rids = get_ordered_rids(prs_xml)

    sm = cfg.slot_map
    title_rid            = ordered_rids[sm.title_slide_idx - 1]
    closing_rid          = ordered_rids[sm.closing_slide_idx - 1]
    content_rids_tpl     = ordered_rids[sm.content_slide_range[0] - 1 : sm.content_slide_range[1]]
    source_file          = rid_to_file[content_rids_tpl[0]]
    content_rids: list[str] = list(content_rids_tpl)

    # ── 5. Clone extra slides if needed ───────────────────────────────────────
    needed, have = len(content_slides), len(content_rids_tpl)
    if needed > have:
        log.info("Cloning %d extra content slide(s)", needed - have)
        content_rids = clone_content_slides(
            needed, have, unpacked, cfg.scripts_dir,
            source_file, rid_to_file, content_rids,
        )
    else:
        content_rids = content_rids[:needed]

    # ── 6. Rebuild slide order ────────────────────────────────────────────────
    prs_xml = prs_xml_path.read_text(encoding="utf-8")   # re-read after cloning
    prs_xml = rebuild_sld_id_list(prs_xml, title_rid, content_rids, closing_rid)
    prs_xml_path.write_text(prs_xml, encoding="utf-8")
    log.info("Slide order: 1 title + %d content + 1 closing", len(content_rids))

    # ── 7a. Title slide ───────────────────────────────────────────────────────
    first         = raw_slides[0]
    meeting_title = first["title"].strip()
    date_bullet   = next(
        (b for b in first["bullets"] if b.lower().startswith("date:")), ""
    )
    date_str = date_bullet.split(":", 1)[-1].strip() if date_bullet else ""

    title_path = unpacked / "ppt/slides" / rid_to_file[title_rid]
    title_xml  = title_path.read_text(encoding="utf-8")
    title_xml  = inject_title_slide(title_xml, meeting_title, date_str)
    title_path.write_text(title_xml, encoding="utf-8")
    log.info("Title slide: '%s' / %s", meeting_title, date_str or "(no date)")

    # ── 7b. Content slides ────────────────────────────────────────────────────
    geo = cfg.geometry
    for i, cs in enumerate(content_slides):
        slide_path = unpacked / "ppt/slides" / rid_to_file[content_rids[i]]
        xml = slide_path.read_text(encoding="utf-8")
        xml = inject_section_title(xml, cs.title)
        xml = inject_bullets(xml, cs.bullets, geo)
        xml = hide_unused_circles(xml, len(cs.bullets), geo)
        slide_path.write_text(xml, encoding="utf-8")
        log.info(
            "  [%d/%d] '%s' — %d bullet(s)",
            i + 1, len(content_slides), cs.title, len(cs.bullets),
        )

    # ── 8. Clean ──────────────────────────────────────────────────────────────
    _run_script(cfg.scripts_dir, "clean.py", str(unpacked))

    # ── 9. Pack → return bytes ────────────────────────────────────────────────
    output_path = cfg.workdir / "output.pptx"
    stdout = _run_script(
        cfg.scripts_dir, "office/pack.py",
        str(unpacked), str(output_path),
        "--original", str(template_path),
    )
    for line in stdout.strip().splitlines():
        if any(kw in line for kw in ("PASS", "packed", "Error", "Auto-repaired")):
            log.info("  pack: %s", line.strip())

    return output_path.read_bytes()
