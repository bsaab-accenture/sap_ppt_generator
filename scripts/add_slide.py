#!/usr/bin/env python3
"""
add_slide.py
────────────
Clone a slide in an unpacked PPTX directory.

Clones the specified slide file, updates relationships, and registers the new slide
in Content_Types.xml. Returns the new relationship ID.

Usage:
    python add_slide.py <unpacked_dir> <source_slide_file>
    
Example:
    python add_slide.py ./unpacked slide3.xml
    
Output:
    r:id="rId42"
"""

import re
import sys
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        print("Usage: python add_slide.py <unpacked_dir> <source_slide_file>", file=sys.stderr)
        sys.exit(1)
    
    unpacked_dir = Path(sys.argv[1])
    source_file = sys.argv[2]  # e.g., "slide3.xml"
    
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = unpacked_dir / "ppt" / "_rels"
    prs_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"
    content_types_path = unpacked_dir / "[Content_Types].xml"
    
    # Find the source slide
    source_slide_path = slides_dir / source_file
    if not source_slide_path.exists():
        print(f"Error: Source slide not found: {source_slide_path}", file=sys.stderr)
        sys.exit(1)
    
    # Find the highest existing slide number
    existing_slides = list(slides_dir.glob("slide*.xml"))
    slide_numbers = []
    for slide_path in existing_slides:
        match = re.match(r'slide(\d+)\.xml', slide_path.name)
        if match:
            slide_numbers.append(int(match.group(1)))
    
    new_slide_num = max(slide_numbers) + 1 if slide_numbers else 1
    new_slide_file = f"slide{new_slide_num}.xml"
    new_slide_path = slides_dir / new_slide_file
    
    # Clone the slide file
    new_slide_path.write_bytes(source_slide_path.read_bytes())
    
    # Clone the slide's relationship file if it exists
    source_slide_rels = unpacked_dir / "ppt" / "slides" / "_rels" / f"{source_file}.rels"
    if source_slide_rels.exists():
        new_slide_rels = unpacked_dir / "ppt" / "slides" / "_rels" / f"{new_slide_file}.rels"
        new_slide_rels.parent.mkdir(parents=True, exist_ok=True)
        new_slide_rels.write_bytes(source_slide_rels.read_bytes())
    
    # Update presentation.xml.rels
    prs_rels_content = prs_rels_path.read_text(encoding="utf-8")
    
    # Find the highest existing rId
    rid_pattern = r'Id="rId(\d+)"'
    rid_numbers = [int(m.group(1)) for m in re.finditer(rid_pattern, prs_rels_content)]
    new_rid_num = max(rid_numbers) + 1 if rid_numbers else 1
    new_rid = f"rId{new_rid_num}"
    
    # Add new relationship entry before </Relationships>
    new_rel_entry = f'  <Relationship Id="{new_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{new_slide_file}"/>\n'
    prs_rels_content = prs_rels_content.replace('</Relationships>', new_rel_entry + '</Relationships>')
    prs_rels_path.write_text(prs_rels_content, encoding="utf-8")
    
    # Update [Content_Types].xml
    content_types_content = content_types_path.read_text(encoding="utf-8")
    new_override = f'  <Override PartName="/ppt/slides/{new_slide_file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n'
    
    # Add before </Types> if not already present
    if new_slide_file not in content_types_content:
        content_types_content = content_types_content.replace('</Types>', new_override + '</Types>')
        content_types_path.write_text(content_types_content, encoding="utf-8")
    
    # Output the new rId in the expected format
    print(f'r:id="{new_rid}"')


if __name__ == "__main__":
    main()
