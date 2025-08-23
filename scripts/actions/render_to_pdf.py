#!/usr/bin/env python3
# render_to_pdf.py — 修正版
from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML
import json
import sys

from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, PROGRAM_JSON

DATA_FILE = PROGRAM_JSON

# 確保 output 目錄存在
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 讀 JSON
try:
    with DATA_FILE.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
except Exception as e:
    print(f"Failed to load {DATA_FILE}: {e}", file=sys.stderr)
    sys.exit(1)

# 建立 Jinja2 環境（autoescape for html/xml）
env = Environment(
    loader=FileSystemLoader(str(TEMPLATE_DIR)),
    autoescape=select_autoescape(["html", "xml"])
)

try:
    tpl = env.get_template("template.html")
except Exception as e:
    print(f"Template not found in {TEMPLATE_DIR}: {e}", file=sys.stderr)
    sys.exit(1)

# 根據 JSON 結構選擇渲染方式：
# - 如果 JSON 頂層是 dict，tpl.render(**data) 會把每個 key 當成 template 變數
# - 否則用 tpl.render(data=data)
try:
    if isinstance(data, dict):
        html = tpl.render(**data)
    else:
        html = tpl.render(data=data)
except Exception as e:
    print(f"Template rendering failed: {e}", file=sys.stderr)
    sys.exit(1)

# 儲存 HTML（方便 debug / preview）
html_file = OUTPUT_DIR / "program.html"
with html_file.open("w", encoding="utf-8") as f:
    f.write(html)

# 產生 PDF（WeasyPr
