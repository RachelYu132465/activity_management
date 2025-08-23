# render_to_pdf.py
from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML
import json
import os

# paths
TEMPLATES_DIR = "templates"
DATA_FILE = "data/program.json"
OUT_HTML = "output/program.html"
OUT_PDF = "output/program.pdf"

os.makedirs("output", exist_ok=True)

# load data
with open(DATA_FILE, "r", encoding="utf8") as f:
    data = json.load(f)

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
