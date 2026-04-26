#!/usr/bin/env python3
"""Package the skill into a .skill zip file."""
import zipfile
import os
from pathlib import Path

SKILL_NAME = "ppt-translator"
FILES = [
    "SKILL.md",
    "scripts/translate_pptx.py",
    "quick_validate.py",
]

def package():
    out = f"{SKILL_NAME}.skill"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in FILES:
            if Path(f).exists():
                zf.write(f)
                print(f"  + {f}")
    print(f"\nCreated: {out} ({Path(out).stat().st_size} bytes)")

if __name__ == "__main__":
    package()
