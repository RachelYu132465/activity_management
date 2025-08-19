
"""Validate that program_data.json contains valid JSON."""

from pathlib import Path
import json


# Cross-platform path construction
p = Path("data") / "shared" / "program_data.json"

# Read the file and attempt to parse as JSON
s = p.read_text(encoding="utf-8")
try:
    json.loads(s)
    print("OK, length:", len(s))
except json.JSONDecodeError as e:
    print("ERROR @ line", e.lineno, "col", e.colno, "char", e.pos)
    print("Tail preview:", s[e.pos:e.pos+120].replace("\n","\\n"))

