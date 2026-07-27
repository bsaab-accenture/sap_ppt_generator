"""
test_local_docx_to_txt.py
─────────────────────────
Local test for the docx_to_txt extraction logic.

Usage:
    python test_local_docx_to_txt.py

What it does:
  1. Reads the sample .docx from local_storage/templates-docx/
  2. Base64-encodes it (mirrors the HTTP contract)
  3. Imports and calls extract_text_from_docx directly — no HTTP, no API calls
  4. Saves output to local_storage/output-docx/transcript.txt
  5. Prints stats and PASS/FAIL

PASS condition: extracted chars >= 80% of all <w:t> text visible inside the
docx XML. This is the true ground truth — raw file size is dominated by binary
overhead (styles, fonts, image data) that has nothing to do with visible text.
"""

from __future__ import annotations

import base64
import io
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent

# Priority: SAP transcript → convertor template → meeting minutes template
_SAP = ROOT / "local_storage" / "SAP_Executive_Board_Meeting_Dummy_Transcript.docx"
_CONVERTOR = ROOT / "local_storage" / "templates-convertor" / "Transcript.docx.docx"
_TEMPLATE = ROOT / "local_storage" / "templates-docx" / "Preliminary_Minutes_BM20260219.docx"
DOCX_PATH = _SAP if _SAP.exists() else (_CONVERTOR if _CONVERTOR.exists() else _TEMPLATE)

OUTPUT_DIR = ROOT / "local_storage" / "output-docx"
OUTPUT_PATH = OUTPUT_DIR / "transcript.txt"


def _xml_visible_text(raw: bytes) -> str:
    """Extract all <w:t> text from the docx ZIP — the true text ground truth."""
    texts: list[str] = []
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            for name in z.namelist():
                if name.startswith("word/") and name.endswith(".xml"):
                    content = z.read(name).decode("utf-8", errors="replace")
                    texts.extend(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", content))
    except Exception as exc:
        print(f"  (warning: could not read docx XML for threshold: {exc})")
    return "".join(texts)


def main() -> None:
    print("=" * 60)
    print("DOCX → TXT  local extraction test")
    print("=" * 60)

    if not DOCX_PATH.exists():
        print(f"FAIL: sample .docx not found at {DOCX_PATH}")
        sys.exit(1)

    # 1. Read raw bytes
    raw = DOCX_PATH.read_bytes()
    print(f"Input file   : {DOCX_PATH.name}")
    print(f"Input size   : {len(raw):,} bytes")

    # 2. Base64-encode (round-trip mirrors HTTP contract)
    docx_b64 = base64.b64encode(raw).decode("ascii")
    doc_bytes = base64.b64decode(docx_b64)

    # 3. Import and call extraction logic — no HTTP, no external calls
    sys.path.insert(0, str(ROOT))
    from docx_core.triggers import extract_text_from_docx

    print("\nExtracting text...")
    text = extract_text_from_docx(doc_bytes)

    # 4. Save output
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(text, encoding="utf-8")
    print(f"Output saved : {OUTPUT_PATH}")

    # 5. Stats
    lines = text.splitlines()
    char_count = len(text)
    line_count = len(lines)

    print("\n--- FIRST 500 CHARS " + "-" * 40)
    print(text[:500])
    print("\n--- LAST 200 CHARS " + "-" * 41)
    print(text[-200:])

    # Ground-truth XML text for threshold comparison
    xml_text = _xml_visible_text(raw)
    xml_char_count = len(xml_text)
    threshold = max(xml_char_count * 0.8, 50)  # must capture ≥ 80% of XML text

    status = "PASS" if char_count >= threshold else "FAIL"

    print("\n" + "=" * 60)
    print(f"Total characters extracted : {char_count:,}")
    print(f"Total lines extracted      : {line_count:,}")
    print(f"XML visible text chars     : {xml_char_count:,}")
    print(f"Threshold (xml_chars*0.80) : {threshold:,.0f}")
    print(f"Result : {status}")
    print("=" * 60)

    if status == "FAIL":
        print("ERROR: output is smaller than expected — check extraction logic.")
        sys.exit(1)


if __name__ == "__main__":
    main()
