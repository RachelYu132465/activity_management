#!/usr/bin/env python3
"""Render HTML to PDF using headless Chrome with centralized paths (direct import style)."""

from pathlib import Path
import sys
import json
import subprocess
import argparse
from datetime import datetime, timedelta

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from jinja2 import Environment, FileSystemLoader, select_autoescape

# Direct import from bootstrap (requested "direct" style)
from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, DATA_DIR, CHROME_BIN

DATA_FILE = DATA_DIR / "shared" / "program_data.json"
INFLUENCER_FILE = DATA_DIR / "shared" / "influencer_data.json"
# Ensure output directory exists
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Load program and influencer data
try:
    with DATA_FILE.open("r", encoding="utf-8") as fh:
        program_list = json.load(fh)
    with INFLUENCER_FILE.open("r", encoding="utf-8") as fh:
        influencer_list = json.load(fh)
except Exception as e:
    print(f"Failed to load data files: {e}", file=sys.stderr)
    sys.exit(1)

# Select program entry
parser = argparse.ArgumentParser(description="Render program handbook")
parser.add_argument("--event-id", type=int, default=None, help="Program id to render")
args = parser.parse_args()

selected = None
if args.event_id is not None:
    for prog in program_list:
        if int(prog.get("id", -1)) == args.event_id:
            selected = prog
            break
if selected is None:
    selected = program_list[0] if program_list else {}

# Build schedule from agenda settings
def build_schedule(event):
    cfg = event.get("agenda_settings", {})
    try:
        current = datetime.strptime(cfg.get("start_time", "00:00"), "%H:%M")
    except ValueError:
        return []
    speaker_minutes = int(cfg.get("speaker_minutes") or 0)
    specials = cfg.get("special_sessions", [])

    def add_special(after_no, start):
        for s in specials:
            if int(s.get("after_speaker", -1)) == after_no:
                dur = int(s.get("duration") or 0)
                end = start + timedelta(minutes=dur)
                schedule.append({
                    "time": f"{start.strftime('%H:%M')}-{end.strftime('%H:%M')}",
                    "topic": s.get("title", ""),
                    "speaker": "",
                    "note": "",
                })
                start = end
        return start

    schedule = []
    current = add_special(0, current)
    for sp in event.get("speakers", []):
        end = current + timedelta(minutes=speaker_minutes)
        schedule.append({
            "time": f"{current.strftime('%H:%M')}-{end.strftime('%H:%M')}",
            "topic": sp.get("topic", ""),
            "speaker": sp.get("name", ""),
            "note": "",
        })
        current = end
        current = add_special(sp.get("no"), current)
    add_special(999, current)
    return schedule

# Build speaker and chair lists augmented from influencer data
infl_by_name = {p.get("name"): p for p in influencer_list}
chairs = []
speakers = []
for sp in selected.get("speakers", []):
    name = sp.get("name")
    info = infl_by_name.get(name, {})
    enriched = {
        "name": name,
        "title": info.get("current_position", {}).get("title", ""),
        "profile": "\n".join(info.get("experience", [])),
        "photo_url": info.get("photo_url", ""),
    }
    if sp.get("type") == "主持人":
        chairs.append(enriched)
    else:
        speakers.append(enriched)

program_data = {
    "title": selected.get("eventNames", [""])[0],
    "date": selected.get("date", ""),
    "locations": selected.get("locations", []),
    "organizers": selected.get("organizers", []),
    "co_organizers": selected.get("coOrganizers", []),
    "schedule": build_schedule(selected),
    "chairs": chairs,
    "speakers": speakers,
    "notes": selected.get("notes", []),
    "contact": selected.get("contact", ""),
}

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
context = {
    "program_data": program_data,
    "organizers": program_data.get("organizers", []),
    "co_organizers": program_data.get("co_organizers", []),
    "assets": {},
}
try:
    html = tpl.render(**context)
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
