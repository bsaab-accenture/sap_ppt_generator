"""Run Agenda PPT generation locally without Azure Functions runtime."""
import json
import os
from pathlib import Path

os.environ["USE_LOCAL_STORAGE"] = "true"

from ppt_core import blob_store_local as blob_store
from ppt_core.generator import GeneratorConfig, generate_pptx


def main():
    json_path = Path("./local_storage/input-json/Agenda_Presentation_Inputs.json")
    content_json = json.loads(json_path.read_text(encoding="utf-8"))

    template_bytes = blob_store.download_blob("templates-ppt", "Agenda_BM20260116.pptx")

    cfg = GeneratorConfig(
        scripts_dir=Path("./scripts"),
        workdir=Path("./tmp"),
        max_bullets=4,
    )

    pptx_bytes = generate_pptx(template_bytes, content_json, cfg)

    output_path = Path("./local_storage/output-agenda_ppt/Agenda_PPT.pptx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(pptx_bytes)

    print(f"Generated: {output_path}")


if __name__ == "__main__":
    main()
