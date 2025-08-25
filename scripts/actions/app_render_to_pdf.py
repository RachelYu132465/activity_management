#!/usr/bin/env python3
"""Render the Flask app's HTML to PDF using headless Chrome."""

from pathlib import Path
import sys
import subprocess
import argparse

# Ensure project root on sys.path
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from flask import render_template
from scripts.actions.app import app, get_context_for_event
from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, CHROME_BIN


def render_to_pdf(event_id: int | None = None) -> None:
    """Render template with context and convert to PDF.

    Saves the intermediate HTML and the final PDF into OUTPUT_DIR.
    If Chrome is not available, only the HTML is generated and a warning is printed.
    """
    ctx = get_context_for_event(event_id)
    # Render HTML using the Flask app's template context
    with app.app_context():
        html = render_template("template.html", **ctx)

    # Ensure output directory exists
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    html_file = OUTPUT_DIR / "app_render.html"
    pdf_file = OUTPUT_DIR / "app_render.pdf"

    html_file.write_text(html, encoding="utf-8")

    if not CHROME_BIN:
        print("Chrome executable not found; set CHROME_BIN or config/paths.json.")
        print(f"HTML saved to {html_file}")
        return

    cmd = [
        CHROME_BIN,
        "--headless", "--disable-gpu",
        f"--print-to-pdf={pdf_file}",
        str(html_file),
    ]

    subprocess.run(cmd, check=True)
    print(f"Saved PDF to {pdf_file}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Render the Flask app template to PDF.")
    parser.add_argument("--event-id", type=int, default=None, help="Program id to render")
    args = parser.parse_args()
    render_to_pdf(args.event_id)


if __name__ == "__main__":
    main()
