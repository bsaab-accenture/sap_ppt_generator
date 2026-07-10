"""Run PPT generation locally without Azure Functions runtime."""
import json
from pathlib import Path
from ppt_core.generator import generate_pptx, GeneratorConfig, TemplateGeometry, TemplateSlotMap

def main():
    # Read input JSON
    json_path = Path("./local_storage/input-json/test_meeting.json")
    content_json = json.loads(json_path.read_text())
    
    # Read template
    template_path = Path("./local_storage/templates/PowerpointTemplate_BoardMeetingGovernance.pptx")
    template_bytes = template_path.read_bytes()
    
    # Generate
    cfg = GeneratorConfig(
        scripts_dir=Path("./scripts"),
        workdir=Path("./tmp"),
        max_bullets=4,
        geometry=TemplateGeometry(),
        slot_map=TemplateSlotMap(),
    )
    
    pptx_bytes = generate_pptx(template_bytes, content_json, cfg)
    
    # Save output
    output_path = Path("./local_storage/output-pptx/generated.pptx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(pptx_bytes)
    
    print(f"✅ Generated: {output_path}")

if __name__ == "__main__":
    main()