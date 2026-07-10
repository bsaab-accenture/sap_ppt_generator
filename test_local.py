#!/usr/bin/env python3
"""
test_local.py
─────────────
Standalone test script to run PPT generation locally without Azure Functions runtime.
Useful for quick testing and debugging.
"""

import json
import logging
import shutil
import tempfile
from pathlib import Path

# Add ppt_core to path
import sys
sys.path.insert(0, str(Path(__file__).parent))

from ppt_core import blob_store_local as blob_store
from ppt_core.generator import GeneratorConfig, TemplateGeometry, TemplateSlotMap, generate_pptx

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


def main():
    """Run PPT generation using local files."""
    
    # Configuration
    template_container = "templates"
    template_name = "PowerpointTemplate_BoardMeetingGovernance.pptx"
    input_container = "input-json"
    output_container = "output-pptx"
    
    # Find test JSON files
    input_dir = Path("local_storage") / input_container
    json_files = list(input_dir.glob("*.json"))
    
    if not json_files:
        log.error("❌ No JSON files found in %s", input_dir)
        log.info("   Add a test JSON file and try again.")
        return 1
    
    # Process first JSON file
    json_file = json_files[0]
    log.info("📄 Processing: %s", json_file.name)
    
    # Read input JSON
    content_json = json.loads(json_file.read_text())
    log.info("✅ Parsed JSON: %d top-level keys", len(content_json))
    
    # Download template
    try:
        template_bytes = blob_store.download_blob(template_container, template_name)
        log.info("✅ Template loaded: %d bytes", len(template_bytes))
    except FileNotFoundError as e:
        log.error("❌ Template not found: %s", e)
        log.info("   Run: ./setup_local_storage.sh")
        return 1
    
    # Generate PPTX
    workdir = Path(tempfile.mkdtemp(prefix="pptx_test_"))
    try:
        cfg = GeneratorConfig(
            scripts_dir=Path("./scripts"),
            workdir=workdir,
            max_bullets=4,
            geometry=TemplateGeometry(),
            slot_map=TemplateSlotMap(),
        )
        
        log.info("⚙️  Generating PPTX...")
        pptx_bytes = generate_pptx(
            template_bytes=template_bytes,
            content_json=content_json,
            cfg=cfg,
        )
        log.info("✅ PPTX generated: %d bytes", len(pptx_bytes))
        
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    
    # Save output
    output_name = f"{json_file.stem}_generated.pptx"
    blob_store.upload_blob(output_container, output_name, pptx_bytes)
    
    # Generate local URL
    sas_url = blob_store.generate_sas_url(output_container, output_name, 24)
    
    log.info("")
    log.info("🎉 SUCCESS!")
    log.info("   Output: local_storage/%s/%s", output_container, output_name)
    log.info("   URL: %s", sas_url)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
