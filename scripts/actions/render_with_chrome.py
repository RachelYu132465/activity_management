#!/usr/bin/env python3
"""Render HTML to PDF using headless Chrome with centralized paths (direct import style)."""

from __future__ import print_function, unicode_literals

try:  # Python 2 fallback
    from pathlib import Path
except ImportError:  # pragma: no cover - pathlib2 used only on legacy Python
    from pathlib2 import Path  # type: ignore

import sys
import json
import subprocess
import argparse
from datetime import datetime, timedelta

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


# <<<<<<< HEAD
# from jinja2 import Environment, FileSystemLoader, select_autoescape, Undefined
# import traceback
#
# from jinja2.exceptions import UndefinedError
# =======
from jinja2 import Environment, FileSystemLoader, select_autoescape, StrictUndefined, Undefined
from jinja2.exceptions import UndefinedError, TemplateNotFound


import traceback

# Direct import from bootstrap (requested "direct" style)
from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, DATA_DIR, CHROME_BIN


DATA_FILE = DATA_DIR / "shared" / "program_data.json"
INFLUENCER_FILE = DATA_DIR / "shared" / "influencer_data.json"

# Ensure output directory exists
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Load program and influencer data (expect list or dict)
try:
    with DATA_FILE.open("r", encoding="utf-8") as fh:
        programs_raw = json.load(fh)
except (OSError, ValueError) as e:
    print("Failed to load program data file {}: {}".format(DATA_FILE, e), file=sys.stderr)
    sys.exit(1)

try:
    with INFLUENCER_FILE.open("r", encoding="utf-8") as fh:
        influencer_list = json.load(fh)
except (OSError, ValueError) as e:
    print("Failed to load influencer data file {}: {}".format(INFLUENCER_FILE, e), file=sys.stderr)
    influencer_list = []

# Select program entry and normalize to a dict
parser = argparse.ArgumentParser(description="Render program handbook")
parser.add_argument("--event-id", type=int, default=None, help="Program id to render")
args = parser.parse_args()

program_data = {}
if isinstance(programs_raw, list):
    if args.event_id is not None:
        for prog in programs_raw:
            try:
                if int(prog.get("id", -1)) == args.event_id:
                    program_data = prog
                    break
            except (ValueError, TypeError):
                continue
    if not program_data and programs_raw:
        program_data = programs_raw[0]
elif isinstance(programs_raw, dict):
    program_data = programs_raw
else:
    print("Unexpected JSON structure in {} (expected list or dict).".format(DATA_FILE), file=sys.stderr)
    sys.exit(1)

# Build schedule from agenda settings
def build_schedule(event):
    cfg = event.get("agenda_settings", {}) or {}
    try:
        current = datetime.strptime(cfg.get("start_time", "00:00"), "%H:%M")
    except (ValueError, TypeError):
        return []
    speaker_minutes = int(cfg.get("speaker_minutes") or 0)
    specials = cfg.get("special_sessions", []) or []

    def add_special(after_no, start):
        for s in specials:
            try:
                if int(s.get("after_speaker", -1)) == int(after_no):
                    dur = int(s.get("duration") or 0)
                    end = start + timedelta(minutes=dur)
                    schedule.append({
                        "time": "{}-{}".format(start.strftime('%H:%M'), end.strftime('%H:%M')),
                        "topic": s.get("title", ""),
                        "speaker": "",

                    })
                    start = end
            except (ValueError, TypeError):
                continue
        return start

    schedule = []
    current = add_special(0, current)
    for speaker_item in event.get("speakers", []) or []:
        end = current + timedelta(minutes=speaker_minutes)
        schedule.append({
            "time": "{}-{}".format(current.strftime('%H:%M'), end.strftime('%H:%M')),
            "topic": speaker_item.get("topic", ""),
            "speaker": speaker_item.get("name", ""),
            "note": "",
        })
        current = end
        current = add_special(speaker_item.get("no", 0), current)
    add_special(999, current)
    return schedule

# Build speaker and chair lists augmented from influencer data
infl_by_name = {p.get("name"): p for p in influencer_list if isinstance(p, dict)}
chairs = []
speakers = []
for speaker_entry in program_data.get("speakers", []) or []:
    name = speaker_entry.get("name")
    info = infl_by_name.get(name, {}) or {}
    enriched = {
        "name": name,
        "title": info.get("current_position", {}).get("title", "") if isinstance(info.get("current_position"), dict) else "",
        "profile": "\n".join(info.get("experience", [])) if isinstance(info.get("experience"), list) else info.get("experience", "") or "",
        "photo_url": info.get("photo_url", ""),
    }
    if speaker_entry.get("type") == "主持人":
        chairs.append(enriched)
    else:
        speakers.append(enriched)

program_data["schedule"] = build_schedule(program_data)
program_data["chairs"] = chairs
program_data["speakers"] = speakers

# Prepare Jinja2 environment
# Custom undefined that logs missing variables instead of raising errors
class LoggingUndefined(Undefined):
    def __str__(self):
        if self._undefined_name is not None:
            print("[render] Missing template variable: {}".format(self._undefined_name), file=sys.stderr)
        return ""


env = Environment(
    loader=FileSystemLoader(str(TEMPLATE_DIR)),
    autoescape=select_autoescape(["html", "xml"]),
    undefined=LoggingUndefined,

)

try:
    tpl = env.get_template("template.html")
except TemplateNotFound as e:
    print("Template not found in {}: {}".format(TEMPLATE_DIR, e), file=sys.stderr)
    sys.exit(1)

# Render HTML directly with raw program data
try:
    render_args = dict(program_data)
    render_args["assets"] = {}
    html = tpl.render(**render_args)
except UndefinedError:
    print("Template rendering failed due to missing variable:", file=sys.stderr)
    traceback.print_exc()
    sys.exit(1)

# Save intermediate HTML
html_file = OUTPUT_DIR / "program.html"
try:
    with html_file.open("w", encoding="utf-8") as f:
        f.write(html)
except OSError as e:
    print("Failed to write HTML preview {}: {}".format(html_file, e), file=sys.stderr)
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
    "--print-to-pdf={}".format(str(pdf_file)),
    str(html_file),
]

# Run Chrome to print PDF
try:
    subprocess.run(cmd, check=True)
    print("Saved PDF to {}".format(pdf_file))
except FileNotFoundError:
    print("Chrome binary not found at: {}. Check CHROME_BIN or config/paths.json.".format(CHROME_BIN), file=sys.stderr)
    sys.exit(1)
except subprocess.CalledProcessError as e:
    print("Chrome rendering failed: {}".format(e), file=sys.stderr)
    print("You can open the HTML preview to debug:", html_file, file=sys.stderr)
    sys.exit(1)
