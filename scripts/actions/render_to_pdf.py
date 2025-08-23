
#!/usr/bin/env python3
# render_to_pdf.py — 修正版
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML
import json

from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, PROGRAM_JSON

DATA_FILE = PROGRAM_JSON



OUTPUT_DIR.makedirs("output", exist_ok=True)


env = Environment(
    loader=FileSystemLoader(TEMPLATES_DIR),
    autoescape=select_autoescape(["html", "xml"])
)
tpl = env.get_template("template.html")
html = tpl.render(**data)

# save html for preview
with open(OUT_HTML, "w", encoding="utf8") as f:
    f.write(html)
print(f"Saved HTML preview to {OUT_HTML}")

# render PDF with WeasyPrint
HTML(string=html, base_url=TEMPLATES_DIR).write_pdf(OUT_PDF)
print(f"Saved PDF to {OUT_PDF}")
