#!/usr/bin/env python3
"""Render HTML to PDF using headless Chrome with centralized paths (direct import style)."""

from pathlib import Path
import sys
import json
import subprocess

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from jinja2 import Environment, FileSystemLoader, select_autoescape

# Direct import from bootstrap (requested "direct" style)
from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, DATA_DIR, CHROME_BIN

DATA_FILE = DATA_DIR / "shared" / "program_data.json"
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
try:
    with html_file.open("w", encoding="utf-8") as f:
        f.write(html)
except Exception as e:
    print(f"Failed to write HTML preview {html_file}: {e}", file=sys.stderr)
    sys.exit(1)

# Prepare PDF path
pdf_file = OUTPUT_DIR / "program.pdf"

# Use centralized CHROME_BIN (imported directly). If absent, instruct user how to fix.
if not CHROME_BIN:
    print(
        "Chrome executable not found (bootstrap.CHROME_BIN is None).\n"
        "Options:\n"
        "  1) Set CHROME_BIN environment variable to your chrome executable (e.g. C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe).\n"
        "  2) Add \"Chrome\": \"C:/path/to/chrome.exe\" to config/paths.json and re-run.\n"
        "  3) Install Chrome/Chromium/Edge so bootstrap can auto-detect it.\n",
        file=sys.stderr,
    )
    sys.exit(1)

# Build chrome command
cmd = [
    CHROME_BIN,
    "--headless",
    "--disable-gpu",
    f"--print-to-pdf={str(pdf_file)}",
    str(html_file),
]

# Run Chrome to print PDF
try:
    subprocess.run(cmd, check=True)
    print(f"Saved PDF to {pdf_file}")
except FileNotFoundError:
    print(f"Chrome binary not found at: {CHROME_BIN}. Check CHROME_BIN or config/paths.json.", file=sys.stderr)
    sys.exit(1)
except subprocess.CalledProcessError as e:
    print(f"Chrome rendering failed: {e}", file=sys.stderr)
    print("You can open the HTML preview to debug:", html_file, file=sys.stderr)
    sys.exit(1)
