#!/usr/bin/env python3
"""
office/unpack.py
────────────────
Unpack a PPTX file (ZIP archive) to a directory.

Usage:
    python unpack.py <pptx_file> <output_dir>
"""

import sys
import zipfile
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        print("Usage: python unpack.py <pptx_file> <output_dir>", file=sys.stderr)
        sys.exit(1)
    
    pptx_file = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    
    if not pptx_file.exists():
        print(f"Error: PPTX file not found: {pptx_file}", file=sys.stderr)
        sys.exit(1)
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Extract PPTX (which is a ZIP file)
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
        zip_ref.extractall(output_dir)
    
    print(f"Unpacked {pptx_file} → {output_dir}")


if __name__ == "__main__":
    main()
