#!/usr/bin/env python3
"""
clean.py
────────
Clean up the unpacked PPTX directory structure.

This script performs housekeeping tasks on the unpacked Office Open XML structure:
- Remove unnecessary metadata
- Clean up temporary files
- Normalize structure

Usage:
    python clean.py <unpacked_dir>
"""

import sys
from pathlib import Path


def main():
    if len(sys.argv) != 2:
        print("Usage: python clean.py <unpacked_dir>", file=sys.stderr)
        sys.exit(1)
    
    unpacked_dir = Path(sys.argv[1])
    
    if not unpacked_dir.exists():
        print(f"Error: Unpacked directory not found: {unpacked_dir}", file=sys.stderr)
        sys.exit(1)
    
    # Remove common temporary/cache files
    patterns = ['**/.DS_Store', '**/~$*', '**/*.tmp']
    removed_count = 0
    
    for pattern in patterns:
        for file_path in unpacked_dir.glob(pattern):
            if file_path.is_file():
                file_path.unlink()
                removed_count += 1
    
    print(f"Cleaned {unpacked_dir} ({removed_count} temporary file(s) removed)")


if __name__ == "__main__":
    main()
