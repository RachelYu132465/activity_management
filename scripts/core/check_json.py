
import json, pathlib
p=pathlib.Path(r"data\shared\program_data.json")
s=p.read_text(encoding="utf-8")
try:
    json.loads(s)
    print("OK, length:", len(s))
except json.JSONDecodeError as e:
    print("ERROR @ line", e.lineno, "col", e.colno, "char", e.pos)
    print("Tail preview:", s[e.pos:e.pos+120].replace("\n","\\n"))

