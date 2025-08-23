#!/usr/bin/env python3
"""Render HTML to PDF using headless Chrome with centralized paths."""
import json
import subprocess
import sys

from jinja2 import Environment, FileSystemLoader, select_autoescape

from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, PROGRAM_JSON

DATA_FILE = PROGRAM_JSON

# Ensure output directory exists
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Load program data
try:
    with DATA_FILE.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
except Exception as e:
    print(f"Failed to load {DATA_FILE}: {e}", file=sys.stderr)
    sys.exit(1)

# Prepare Jinja2 environment
env = Environment(
    loader=FileSystemLoader(str(TEMPLATE_DIR)),
    autoescape=select_autoescape(["html", "xml"]),
)

try:
    tpl = env.get_template("template.html")
except Exception as e:
    print(f"Template not found in {TEMPLATE_DIR}: {e}", file=sys.stderr)
    sys.exit(1)

# Render HTML
try:
    if isinstance(data, dict):
        html = tpl.render(**data)
    else:
        html = tpl.render(data=data)
except Exception as e:
    print(f"Template rendering failed: {e}", file=sys.stderr)
    sys.exit(1)

# Save intermediate HTML
html_file = OUTPUT_DIR / "program.html"
with html_file.open("w", encoding="utf-8") as f:
    f.write(html)

# Render PDF via headless Chrome
pdf_file = OUTPUT_DIR / "program.pdf"
cmd = [
    "google-chrome",
    "--headless",
    "--disable-gpu",
    f"--print-to-pdf={pdf_file}",
    str(html_file),
]

try:
    subprocess.run(cmd, check=True)
except FileNotFoundError:
    print("google-chrome not found; ensure Chrome is installed.", file=sys.stderr)
    sys.exit(1)
except subprocess.CalledProcessError as e:
    print(f"Chrome rendering failed: {e}", file=sys.stderr)
    sys.exit(1)
