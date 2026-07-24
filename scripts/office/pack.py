#!/usr/bin/env python3
"""
office/pack.py
──────────────
Pack a directory back into a PPTX file (ZIP archive).

Usage:
    python pack.py <unpacked_dir> <output_pptx> [--original <original_pptx>]
"""

import sys
import zipfile
from pathlib import Path


def main():
    if len(sys.argv) < 3:
        print("Usage: python pack.py <unpacked_dir> <output_pptx> [--original <original_pptx>]", file=sys.stderr)
        sys.exit(1)
    
    unpacked_dir = Path(sys.argv[1])
    output_pptx = Path(sys.argv[2])
    
    # Optional: --original parameter (for reference)
    original_pptx = None
    if len(sys.argv) > 3 and sys.argv[3] == "--original":
        original_pptx = Path(sys.argv[4]) if len(sys.argv) > 4 else None
    
    if not unpacked_dir.exists():
        print(f"Error: Unpacked directory not found: {unpacked_dir}", file=sys.stderr)
        sys.exit(1)
    
    # Create PPTX (ZIP file) from directory
    with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through all files in the directory
        for file_path in unpacked_dir.rglob('*'):
            if file_path.is_file():
                # Calculate archive name (relative path from unpacked_dir)
                arcname = file_path.relative_to(unpacked_dir)
                zipf.write(file_path, arcname)
    
    print(f"PASS: Packed {unpacked_dir} -> {output_pptx}")
    print(f"      Size: {output_pptx.stat().st_size} bytes")


if __name__ == "__main__":
    main()
