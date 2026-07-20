"""test_local_docx.py
Local smoke-test for the agenda DOCX generator — no Azure dependencies.

Usage
-----
    python test_local_docx.py [<json_file>]

    <json_file>  Path to an agendadocx_input.json file.
                 Defaults to: local_storage/input-docx-json/agendadocx_input.json

Output is written to: local_storage/output-docx/<stem>_generated.docx
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

# ── Ensure project root is on the import path ─────────────────────────────────
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from docx_core.generator import generate_docx, validate_input  # noqa: E402

INPUT_DEFAULT = ROOT / "local_storage" / "input-docx-json" / "agendadocx_input.json"
TEMPLATE_PATH = ROOT / "local_storage" / "templates-docx" / "Preliminary_Minutes_BM20260219.docx"
OUTPUT_DIR = ROOT / "local_storage" / "output-docx"


def main() -> None:
    input_path = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_DEFAULT

    # ── Validate paths ────────────────────────────────────────────────────────
    for p, label in ((input_path, "Input JSON"), (TEMPLATE_PATH, "DOCX template")):
        if not p.exists():
            print(f"ERROR  {label} not found: {p}")
            sys.exit(1)

    print(f"Input    : {input_path}")
    print(f"Template : {TEMPLATE_PATH}")

    # ── Load inputs ───────────────────────────────────────────────────────────
    content_json = json.loads(input_path.read_text(encoding="utf-8"))
    template_bytes = TEMPLATE_PATH.read_bytes()

    # ── Quick schema check before generation ─────────────────────────────────
    try:
        body = validate_input(content_json)
        print(f"Schema   : OK  ({len(body['Topics_Discussed'])} topics)")
    except ValueError as exc:
        print(f"Schema ERROR: {exc}")
        sys.exit(1)

    # ── Generate ──────────────────────────────────────────────────────────────
    t0 = time.monotonic()
    docx_bytes = generate_docx(template_bytes=template_bytes, content_json=content_json)
    elapsed = round(time.monotonic() - t0, 2)

    # ── Write output ──────────────────────────────────────────────────────────
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem
    out_path = OUTPUT_DIR / f"{stem}_generated.docx"
    out_path.write_bytes(docx_bytes)

    print(f"Output   : {out_path}  ({len(docx_bytes):,} bytes)  [{elapsed}s]")
    print("Done.")


if __name__ == "__main__":
    main()
