#!/usr/bin/env python3
#
# Minimal live preview server that renders templates/template.html.
# This version expands the selected program's top-level fields into the template context,
# so all JSON fields (not just a fixed subset) are available to Jinja templates.
#
# Usage:
#   cd C:\Users\User\activity_management
#   python scripts\core\simple_server.py --serve
#   http://127.0.0.1:5000/?event_id=3
#   http://127.0.0.1:5000/event/3

from __future__ import print_function, unicode_literals
from pathlib import Path
import sys
import json
import argparse
import traceback
from datetime import datetime, timedelta

from flask import Flask, render_template, request, Response
from jinja2 import Undefined
from jinja2.exceptions import TemplateNotFound

# Project layout: this file is scripts/core/simple_server.py
PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = PROJECT_ROOT / "templates"
DATA_FILE = PROJECT_ROOT / "data" / "shared" / "program_data.json"

# make sure project root is on sys.path (helps future imports)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# minimal Flask app pointed at project templates folder
app = Flask(__name__, template_folder=str(TEMPLATE_DIR))
app.jinja_env.undefined = Undefined  # non-strict: missing keys render empty

# ---------- helpers ----------
def load_json_safe(path):
    """Return parsed JSON or None if file missing/invalid."""
    if not path.exists():
        return None
    try:
        with path.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception as e:
        print("Warning: failed to load JSON", path, e)
        return None

def pick_program_by_id(programs_raw, event_id):
    """
    programs_raw: list or dict
    event_id: int or None
    Returns: a dict representing the selected program (or {})
    """
    if isinstance(programs_raw, dict):
        # single-object JSON
        return programs_raw
    if not isinstance(programs_raw, list) or not programs_raw:
        return {}

    # if event_id provided, search
    if event_id is not None:
        for p in programs_raw:
            try:
                if int(p.get("id", -1)) == int(event_id):
                    return p if isinstance(p, dict) else {}
            except Exception:
                continue

    # fallback: return first element
    first = programs_raw[0]
    return first if isinstance(first, dict) else {}

def _normalize_event_names(program):
    ev = []
    if isinstance(program.get("eventNames"), list):
        ev = program.get("eventNames")
    elif isinstance(program.get("eventNames"), str):
        ev = [program.get("eventNames")]
    else:
        # fallback to title if available
        title_cand = program.get("title") or (program.get("program") or {}).get("title", "")
        if title_cand:
            ev = [title_cand]
    return ev

def _build_schedule_from_agenda(program):
    """Attempt to build a simple schedule from agenda_settings and speakers."""
    schedule = []
    try:
        cfg = (program.get("agenda_settings") or {}) or {}
        start = cfg.get("start_time")
        minutes = int(cfg.get("speaker_minutes") or 0) if cfg else 0
        if start and minutes and isinstance(program.get("speakers"), list):
            cur = datetime.strptime(start, "%H:%M")
            for sp in program.get("speakers", []) or []:
                end = cur + timedelta(minutes=minutes)
                schedule.append({
                    "time": f"{cur.strftime('%H:%M')}-{end.strftime('%H:%M')}",
                    "topic": sp.get("topic", "") if isinstance(sp, dict) else "",
                    "speaker": sp.get("name", "") if isinstance(sp, dict) else "",
                })
                cur = end
    except Exception:
        schedule = []
    return schedule

def _enrich_speakers_and_chairs(program, influencer_list=None):
    """Return (speakers_list, chairs_list) with some enrichment if influencer data is available."""
    influencer_list = influencer_list or []
    infl_by_name = {p.get("name"): p for p in (influencer_list or []) if isinstance(p, dict)}
    chairs = []
    speakers = []
    for s in (program.get("speakers") or []) or []:
        if isinstance(s, dict):
            name = s.get("name", "") or ""
            info = infl_by_name.get(name, {}) or {}
            enriched = {
                "name": name,
                "topic": s.get("topic", "") or "",
                "type": s.get("type", "") or "",
                "no": s.get("no", None),
                "title": (info.get("current_position") or {}).get("title", "") if isinstance(info.get("current_position"), dict) else info.get("current_position", "") or "",
                "profile": "\n".join(info.get("experience", [])) if isinstance(info.get("experience"), list) else (info.get("experience", "") or ""),
                "photo_url": info.get("photo_url", "") or "",
            }
        else:
            # speaker entry not dict, coerce to minimal form
            enriched = {"name": str(s), "topic": "", "type": "", "no": None, "title": "", "profile": "", "photo_url": ""}
        if enriched.get("type") == "主持人":
            chairs.append(enriched)
        else:
            speakers.append(enriched)
    return speakers, chairs

def build_safe_context(program, influencer_list=None):
    """
    Build a template context that:
      - includes ALL top-level fields from the selected program (if it's a dict)
      - provides normalized, convenient fields (eventNames, title, date, locations, organizers, speakers, chairs, schedule, contact)
    """
    if not program or not isinstance(program, dict):
        # return safe empty context with expected keys
        return {
            "eventNames": [],
            "assets": {},
            "program": {},
            "title": "",
            "date": "",
            "locations": [],
            "organizers": [],
            "speakers": [],
            "chairs": [],
            "schedule": [],
            "contact": "",
        }

    # Start by copying all top-level keys so templates can access anything in JSON
    context = {}
    # shallow copy program's top-level keys
    for k, v in program.items():
        context[k] = v

    # normalized conveniences / overrides (ensures presence and safer shapes)
    ev = _normalize_event_names(program)
    context["eventNames"] = ev
    context["assets"] = program.get("assets", {}) or {}
    # keep original program nested if present, else set to program itself
    context["program"] = program.get("program", program) or program
    context["title"] = program.get("title", "") or (ev[0] if ev else "")
    context["date"] = program.get("date", "") or ""
    context["locations"] = program.get("locations", []) or []
    context["organizers"] = program.get("organizers", []) or []
    context["contact"] = program.get("contact", "") or ""

    # enrich speakers and chairs (these will override any raw 'speakers'/'chairs' keys
    # but we still keep the original raw speakers under 'raw_speakers' for full access)
    context["raw_speakers"] = program.get("speakers", []) or []
    speakers, chairs = _enrich_speakers_and_chairs(program)
    context["speakers"] = speakers
    context["chairs"] = chairs

    # schedule: if agenda_settings present try to compute; otherwise keep raw schedule if present
    computed_schedule = _build_schedule_from_agenda(program)
    if computed_schedule:
        context["schedule"] = computed_schedule
    else:
        context["schedule"] = program.get("schedule", []) or []

    # Provide a small helper: a flat 'all_keys' list for debugging/templates if needed
    context["_all_keys"] = list(program.keys())

    return context

def get_context_for_event(event_id):
    """Load JSON, select program, and return context dict with all fields expanded."""
    raw = load_json_safe(DATA_FILE)
    program = pick_program_by_id(raw, event_id)
    ctx = build_safe_context(program)
    return ctx

# ---------- routes ----------
@app.route("/")
def index():
    # query param ?event_id=...
    event_id = request.args.get("event_id", None)
    try:
        event_id = int(event_id) if event_id is not None else None
    except Exception:
        event_id = None

    ctx = get_context_for_event(event_id)
    try:
        return render_template("template.html", **ctx)
    except TemplateNotFound:
        return f"Template template.html not found in {TEMPLATE_DIR}", 404
    except Exception as e:
        traceback.print_exc()
        return f"Rendering error: {e}", 500

@app.route("/event/<int:event_id>")
def event_route(event_id):
    ctx = get_context_for_event(event_id)
    try:
        return render_template("template.html", **ctx)
    except TemplateNotFound:
        return f"Template template.html not found in {TEMPLATE_DIR}", 404
    except Exception as e:
        traceback.print_exc()
        return f"Rendering error: {e}", 500

# Optional: a debug endpoint that returns the context as JSON (handy while developing)
@app.route("/_ctx")
def show_ctx():
    event_id = request.args.get("event_id", None)
    try:
        event_id = int(event_id) if event_id is not None else None
    except Exception:
        event_id = None
    ctx = get_context_for_event(event_id)
    # convert non-serializable items defensively by using json.dumps roundtrip
    try:
        return Response(json.dumps(ctx, ensure_ascii=False, indent=2), mimetype="application/json; charset=utf-8")
    except Exception:
        # fallback: show keys only
        return Response(json.dumps({"keys": list(ctx.keys())}, ensure_ascii=False, indent=2), mimetype="application/json; charset=utf-8")

# ---------- CLI ----------
def main():
    parser = argparse.ArgumentParser(description="Serve templates/template.html with optional event_id")
    parser.add_argument("--serve", action="store_true", help="Run Flask dev server")
    parser.add_argument("--port", type=int, default=5000, help="Port for server")
    parser.add_argument("--event-id", type=int, default=None, help="Program id to render by default")
    args = parser.parse_args()

    print("Project root:", PROJECT_ROOT)
    print("Template dir:", TEMPLATE_DIR)
    print("Data file:", DATA_FILE)
    if args.serve:
        print("Starting server at http://127.0.0.1:%s/" % args.port)
        app.run(host="127.0.0.1", port=args.port, debug=True, use_reloader=True)
    else:
        # render once and print out minimal info (for CLI usage)
        ctx = get_context_for_event(args.event_id)
        print("Rendered context keys:", list(ctx.keys()))
        # also print title/date for quick sanity
        print("title:", ctx.get("title"))
        print("date:", ctx.get("date"))

if __name__ == "__main__":
    main()

#可用!/usr/bin/env python3

# Minimal live preview server that renders templates/template.html.
# Supports selecting a program by event_id via query string or path param.
#
# Usage:
#   cd C:\Users\User\activity_management
#   python scripts\core\simple_server.py --serve
#   # Browse:
#   http://127.0.0.1:5000/                -> default program (first in JSON or fallback)
#   http://127.0.0.1:5000/?event_id=3     -> program with id=3 (if exists)
#   http://127.0.0.1:5000/event/3         -> same as above
#
# from __future__ import print_function, unicode_literals
# from pathlib import Path
# import sys
# import json
# import argparse
# import traceback
#
# from flask import Flask, render_template, request, Response
# from jinja2 import Undefined
# from jinja2.exceptions import TemplateNotFound
#
# # Project layout: this file is scripts/core/simple_server.py
# PROJECT_ROOT = Path(__file__).resolve().parents[2]
# TEMPLATE_DIR = PROJECT_ROOT / "templates"
# DATA_FILE = PROJECT_ROOT / "data" / "shared" / "program_data.json"
#
# # make sure project root is on sys.path (helps future imports)
# if str(PROJECT_ROOT) not in sys.path:
#     sys.path.insert(0, str(PROJECT_ROOT))
#
# # minimal Flask app pointed at project templates folder
# app = Flask(__name__, template_folder=str(TEMPLATE_DIR))
# app.jinja_env.undefined = Undefined  # non-strict: missing keys render empty
#
# # ---------- helpers ----------
# def load_json_safe(path):
#     """Return parsed JSON or None if file missing/invalid."""
#     if not path.exists():
#         return None
#     try:
#         with path.open("r", encoding="utf-8") as fh:
#             return json.load(fh)
#     except Exception as e:
#         print("Warning: failed to load JSON", path, e)
#         return None
#
# def pick_program_by_id(programs_raw, event_id):
#     """
#     programs_raw: list or dict
#     event_id: int or None
#     Returns: a dict representing the selected program (or {})
#     """
#     if isinstance(programs_raw, dict):
#         # single-object JSON
#         return programs_raw
#     if not isinstance(programs_raw, list) or not programs_raw:
#         return {}
#
#     # if event_id provided, search
#     if event_id is not None:
#         for p in programs_raw:
#             try:
#                 if int(p.get("id", -1)) == int(event_id):
#                     return p if isinstance(p, dict) else {}
#             except Exception:
#                 continue
#
#     # fallback: return first element
#     first = programs_raw[0]
#     return first if isinstance(first, dict) else {}
#
# def build_safe_context(program):
#     """Return a safe dict with keys your template likely expects (avoid UndefinedError)."""
#     ev = []
#     if isinstance(program.get("eventNames"), list):
#         ev = program.get("eventNames")
#     elif isinstance(program.get("eventNames"), str):
#         ev = [program.get("eventNames")]
#     else:
#         # fallback to title if available
#         title_cand = program.get("title") or (program.get("program") or {}).get("title", "")
#         if title_cand:
#             ev = [title_cand]
#
#     # normalize speakers/chairs (basic)
#     speakers = []
#     chairs = []
#     for s in (program.get("speakers") or []) or []:
#         ent = {
#             "name": s.get("name", "") if isinstance(s, dict) else "",
#             "topic": s.get("topic", "") if isinstance(s, dict) else "",
#             "type": s.get("type", "") if isinstance(s, dict) else "",
#             "no": s.get("no", None) if isinstance(s, dict) else None,
#         }
#         if ent["type"] == "主持人":
#             chairs.append(ent)
#         else:
#             speakers.append(ent)
#
#     # simple schedule: if agenda_settings exist, attempt build (robustly)
#     schedule = []
#     try:
#         cfg = program.get("agenda_settings") or {}
#         start = cfg.get("start_time")
#         minutes = int(cfg.get("speaker_minutes") or 0) if cfg else 0
#         if start and minutes and isinstance(program.get("speakers"), list):
#             from datetime import datetime, timedelta
#             cur = datetime.strptime(start, "%H:%M")
#             for sp in program.get("speakers", []) or []:
#                 end = cur + timedelta(minutes=minutes)
#                 schedule.append({
#                     "time": f"{cur.strftime('%H:%M')}-{end.strftime('%H:%M')}",
#                     "topic": sp.get("topic", "") if isinstance(sp, dict) else "",
#                     "speaker": sp.get("name", "") if isinstance(sp, dict) else "",
#                 })
#                 cur = end
#     except Exception:
#         schedule = []
#
#     safe = {
#         "eventNames": ev,
#         "assets": program.get("assets", {}) if isinstance(program, dict) else {},
#         "program": program.get("program", program) if isinstance(program, dict) else {},
#         "title": program.get("title", "") if isinstance(program, dict) else (ev[0] if ev else ""),
#         "date": program.get("date", "") if isinstance(program, dict) else "",
#         "locations": program.get("locations", []) if isinstance(program, dict) else [],
#         "organizers": program.get("organizers", []) if isinstance(program, dict) else [],
#         "speakers": speakers,
#         "chairs": chairs,
#         "schedule": schedule,
#         "contact": program.get("contact", "") if isinstance(program, dict) else "",
#     }
#     return safe
#
# def get_context_for_event(event_id):
#     """Load JSON, select program, and return safe context dict."""
#     raw = load_json_safe(DATA_FILE)
#     program = pick_program_by_id(raw, event_id)
#     ctx = build_safe_context(program)
#     return ctx
#
# # ---------- routes ----------
# @app.route("/")
# def index():
#     # query param ?event_id=...
#     event_id = request.args.get("event_id", None)
#     try:
#         event_id = int(event_id) if event_id is not None else None
#     except Exception:
#         event_id = None
#
#     ctx = get_context_for_event(event_id)
#     try:
#         return render_template("template.html", **ctx)
#     except TemplateNotFound:
#         return f"Template template.html not found in {TEMPLATE_DIR}", 404
#     except Exception as e:
#         traceback.print_exc()
#         return f"Rendering error: {e}", 500
#
# @app.route("/event/<int:event_id>")
# def event_route(event_id):
#     # direct path param
#     ctx = get_context_for_event(event_id)
#     try:
#         return render_template("template.html", **ctx)
#     except TemplateNotFound:
#         return f"Template template.html not found in {TEMPLATE_DIR}", 404
#     except Exception as e:
#         traceback.print_exc()
#         return f"Rendering error: {e}", 500
#
# # ---------- CLI ----------
# def main():
#     parser = argparse.ArgumentParser(description="Serve templates/template.html with optional event_id")
#     parser.add_argument("--serve", action="store_true", help="Run Flask dev server")
#     parser.add_argument("--port", type=int, default=5000, help="Port for server")
#     args = parser.parse_args()
#
#     print("Project root:", PROJECT_ROOT)
#     print("Template dir:", TEMPLATE_DIR)
#     print("Data file:", DATA_FILE)
#     if args.serve:
#         print("Starting server at http://127.0.0.1:%s/" % args.port)
#         app.run(host="127.0.0.1", port=args.port, debug=True, use_reloader=True)
#     else:
#         # render once and print out minimal info (for CLI usage)
#         ctx = get_context_for_event(None)
#         print("Rendered context keys:", list(ctx.keys()))
#         print("Run with --serve to start web preview")
#
# if __name__ == "__main__":
#     main()



# #可用!/usr/bin/env python3
# # simple_server.py
# # Minimal Flask app that serves templates/template.html from project root.
# # Usage:
# #   cd C:\Users\User\activity_management
# #   python scripts\core/simple_server.py
# #
# # It uses an absolute TEMPLATE_FOLDER (project_root/templates) so it reliably finds template.html.
#
# from pathlib import Path
# import sys
# from flask import Flask, render_template
# from jinja2 import Undefined
#
# # project root is two parents up from this file (scripts/core/)
# PROJECT_ROOT = Path(__file__).resolve().parents[2]
#
# # absolute templates folder
# TEMPLATE_FOLDER = PROJECT_ROOT / "templates"
#
# # ensure project root is on sys.path (helps imports if you later expand)
# if str(PROJECT_ROOT) not in sys.path:
#     sys.path.insert(0, str(PROJECT_ROOT))
#
# # create Flask app pointing at the absolute templates folder
# app = Flask(__name__, template_folder=str(TEMPLATE_FOLDER))
#
# # use non-strict undefined so missing keys render as empty strings instead of raising
# app.jinja_env.undefined = Undefined
#
# @app.route("/")
# def index():
#     """
#     Render templates/template.html with a minimal safe context to avoid missing-variable errors.
#     If your template requires additional variables, add them to the `test_data` dict below.
#     """
#     test_data = {
#         "eventNames": [],   # used by many templates; set to [] so eventNames[0] won't exist unless you populate
#         "assets": {},
#         "program": {},
#         "title": "",
#         "date": "",
#         "locations": [],
#         "organizers": [],
#         "speakers": [],
#         "chairs": [],
#         "schedule": [],
#         "contact": "",
#     }
#     return render_template("template.html", **test_data)
#
# if __name__ == "__main__":
#     # run development server for local preview
#     # change host/port here if you want it accessible on the LAN
#     print("Serving templates from:", TEMPLATE_FOLDER)
#     app.run(host="127.0.0.1", port=5000, debug=True)


# # put at top of scripts/core/app.py (or replace your Flask(...) line)
# from pathlib import Path
# from flask import Flask, render_template
# import os
#
# # base = project root (two parents up from this file: activity_management)
# PROJECT_ROOT = Path(__file__).resolve().parents[2]
#
# # explicit absolute template folder (use the actual templates folder at project root)
# TEMPLATE_FOLDER = PROJECT_ROOT / "templates"
#
# # create app with absolute template folder
# app = Flask(__name__, template_folder=str(TEMPLATE_FOLDER))
#
# # debug info to confirm where Flask will look
# print("CWD:", os.getcwd())
# print("app.root_path:", app.root_path)
# print("app.template_folder:", app.template_folder)
# try:
#     print("Flask jinja searchpath:", app.jinja_loader.searchpath)
# except Exception as e:
#     print("Cannot read jinja_loader.searchpath:", e)
# print("List templates folder (abs):", [p.name for p in Path(str(app.template_folder)).iterdir()])
#
# # then your routes...
# @app.route("/")
# def index():
#     # minimal test data to avoid missing-key errors
#     test_data = {
#         "assets": {},
#         "program": {},
#         "date": "",
#         "locations": [],
#         # add other keys your template expects
#     }
#     return render_template("template.html", **test_data)
#
#
#
#
#
#
#
#
# # BEGIN PATCH: insert at very top of scripts/actions/app.py
# from pathlib import Path
# import sys
# # make project root explicit: this file is scripts/actions/app.py -> go up two levels
# PROJECT_ROOT = Path(__file__).resolve().parents[2]
# project_root_str = str(PROJECT_ROOT)
# if project_root_str not in sys.path:
#     # put project root first so imports like `import scripts.core.bootstrap` work
#     sys.path.insert(0, project_root_str)
#
# # optional debug (uncomment while testing)
# # import os
# # print("DEBUG: CWD =", os.getcwd())
# # print("DEBUG: PROJECT_ROOT =", project_root_str)
# # print("DEBUG: sys.path[0:5] =", sys.path[0:5])
# # END PATCH
#
# #!/usr/bin/env python3
# """Render HTML to PDF using headless Chrome with centralized paths (direct import style).
#    Added: --serve mode (Flask) to preview live rendering, and safer data loading.
# """
# from __future__ import print_function, unicode_literals
# from pathlib import Path
# import sys
# import json
# import subprocess
# import argparse
# from datetime import datetime, timedelta
# import traceback
# import os
#
# # Optional dev dependency: Flask for --serve mode
# try:
#     from flask import Flask, Response, request, send_file
#     FLASK_AVAILABLE = True
# except Exception:
#     FLASK_AVAILABLE = False
#
# from jinja2 import Environment, FileSystemLoader, select_autoescape, Undefined
# from jinja2.exceptions import UndefinedError, TemplateNotFound
#
# # Direct import from bootstrap (requested "direct" style)
# from scripts.core.bootstrap import TEMPLATE_DIR, OUTPUT_DIR, DATA_DIR, CHROME_BIN
#
# # Paths
# ROOT = Path(__file__).resolve().parents[2]
# DATA_FILE = DATA_DIR / "shared" / "program_data.json"
# INFLUENCER_FILE = DATA_DIR / "shared" / "influencer_data.json"
#
# OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
# OUTPUT_HTML = OUTPUT_DIR / "program.html"
# OUTPUT_PDF = OUTPUT_DIR / "program.pdf"
#
# # ---------- Jinja undefined that logs missing variables ----------
# class LoggingUndefined(Undefined):
#     def __str__(self):
#         if self._undefined_name is not None:
#             print("[render] Missing template variable: {}".format(self._undefined_name), file=sys.stderr)
#         return ""
#
# # ---------- Jinja env ----------
# env = Environment(
#     loader=FileSystemLoader(str(TEMPLATE_DIR)),
#     autoescape=select_autoescape(["html", "xml"]),
#     undefined=LoggingUndefined,
# )
#
# # Try to locate template now (but routes will re-get it each request)
# TEMPLATE_NAME_DEFAULT = "template.html"
#
# # ---------- utility: load raw JSON lists ----------
# def load_raw_json(path):
#     if not path.exists():
#         return None
#     try:
#         with path.open("r", encoding="utf-8") as fh:
#             return json.load(fh)
#     except Exception as e:
#         print(f"Warning: failed to load JSON {path}: {e}", file=sys.stderr)
#         return None
#
# # ---------- robust loader that builds safe context ----------
# def load_program_data(event_id=None):
#     """
#     Returns a safe dict to pass into tpl.render(**data).
#     Ensures eventNames, program, speakers, chairs, schedule exist with sensible defaults.
#     """
#     programs_raw = load_raw_json(DATA_FILE) or []
#     influencer_list = load_raw_json(INFLUENCER_FILE) or []
#
#     # choose program entry
#     program_data = {}
#     if isinstance(programs_raw, list):
#         if event_id is not None:
#             for p in programs_raw:
#                 try:
#                     if int(p.get("id", -1)) == int(event_id):
#                         program_data = p
#                         break
#                 except Exception:
#                     continue
#         if not program_data and programs_raw:
#             program_data = programs_raw[0]
#     elif isinstance(programs_raw, dict):
#         program_data = programs_raw
#     else:
#         program_data = {}
#
#     # safe defaults if keys missing
#     # normalize eventNames
#     ev = []
#     if isinstance(program_data.get("eventNames"), list):
#         ev = program_data.get("eventNames")
#     elif isinstance(program_data.get("eventNames"), str):
#         ev = [program_data.get("eventNames")]
#     else:
#         title_candidate = program_data.get("title") or (program_data.get("program") or {}).get("title", "")
#         if title_candidate:
#             ev = [title_candidate]
#
#     # Build schedule helper
#     def build_schedule(event):
#         cfg = (event.get("agenda_settings") or {}) or {}
#         try:
#             current = datetime.strptime(cfg.get("start_time", "00:00"), "%H:%M")
#         except Exception:
#             return []
#         speaker_minutes = int(cfg.get("speaker_minutes") or 0)
#         specials = cfg.get("special_sessions", []) or []
#
#         def add_special(after_no, start, schedule):
#             for s in specials:
#                 try:
#                     if int(s.get("after_speaker", -1)) == int(after_no):
#                         dur = int(s.get("duration") or 0)
#                         end = start + timedelta(minutes=dur)
#                         schedule.append({
#                             "time": "{}-{}".format(start.strftime('%H:%M'), end.strftime('%H:%M')),
#                             "topic": s.get("title", ""),
#                             "speaker": "",
#                         })
#                         start = end
#                 except Exception:
#                     continue
#             return start
#
#         schedule = []
#         current = add_special(0, current, schedule)
#         for speaker_item in event.get("speakers", []) or []:
#             try:
#                 end = current + timedelta(minutes=speaker_minutes)
#             except Exception:
#                 break
#             schedule.append({
#                 "time": "{}-{}".format(current.strftime('%H:%M'), end.strftime('%H:%M')),
#                 "topic": speaker_item.get("topic", ""),
#                 "speaker": speaker_item.get("name", ""),
#                 "note": "",
#             })
#             current = end
#             current = add_special(speaker_item.get("no", 0), current, schedule)
#         add_special(999, current, schedule)
#         return schedule
#
#     # Enrich speakers with influencer info if available
#     infl_by_name = {p.get("name"): p for p in (influencer_list or []) if isinstance(p, dict)}
#     chairs = []
#     speakers = []
#     for speaker_entry in (program_data.get("speakers") or []) or []:
#         name = speaker_entry.get("name")
#         info = infl_by_name.get(name, {}) or {}
#         enriched = {
#             "name": name or "",
#             "title": (info.get("current_position") or {}).get("title", "") if isinstance(info.get("current_position"), dict) else info.get("current_position", "") or "",
#             "profile": "\n".join(info.get("experience", [])) if isinstance(info.get("experience"), list) else info.get("experience", "") or "",
#             "photo_url": info.get("photo_url", "") or "",
#             "topic": speaker_entry.get("topic", ""),
#             "no": speaker_entry.get("no", None),
#             "type": speaker_entry.get("type", ""),
#         }
#         if speaker_entry.get("type") == "主持人":
#             chairs.append(enriched)
#         else:
#             speakers.append(enriched)
#
#     # compute schedule (works if agenda_settings present)
#     schedule = build_schedule(program_data)
#
#     safe = {
#         "eventNames": ev,
#         "assets": program_data.get("assets", {}),
#         "program": program_data.get("program", program_data),
#         "title": program_data.get("title", "") or (ev[0] if ev else ""),
#         "date": program_data.get("date", ""),
#         "locations": program_data.get("locations", []),
#         "organizers": program_data.get("organizers", []),
#         "speakers": speakers,
#         "chairs": chairs,
#         "schedule": schedule,
#         "contact": program_data.get("contact", ""),
#         # add more defaults here if your template expects them:
#         # "notes": program_data.get("notes", []),
#     }
#     return safe
#
# # ---------- render helper ----------
# def render_to_html_string(template_name=TEMPLATE_NAME_DEFAULT, context=None):
#     try:
#         tpl = env.get_template(template_name)
#     except TemplateNotFound as e:
#         raise
#     if context is None:
#         context = load_program_data()
#     return tpl.render(**context), context
#
# def write_output_html(html_str, path=OUTPUT_HTML):
#     try:
#         with path.open("w", encoding="utf-8") as f:
#             f.write(html_str)
#     except Exception as e:
#         print("Failed to write HTML preview {}: {}".format(path, e), file=sys.stderr)
#         raise
#
# def chrome_print_pdf(html_path=OUTPUT_HTML, pdf_path=OUTPUT_PDF):
#     if not CHROME_BIN:
#         raise FileNotFoundError("Chrome binary not configured (CHROME_BIN is None).")
#     cmd = [
#         CHROME_BIN,
#         "--headless",
#         "--disable-gpu",
#         "--print-to-pdf={}".format(str(pdf_path)),
#         str(html_path),
#     ]
#     subprocess.run(cmd, check=True)
#
# # ---------- CLI ----------
# def main():
#     parser = argparse.ArgumentParser(description="Render program handbook (serve or render-only)")
#     parser.add_argument("--event-id", type=int, default=None, help="Program id to render")
#     parser.add_argument("--render-only", action="store_true", help="Render HTML+PDF and exit (no server)")
#     parser.add_argument("--template", default=TEMPLATE_NAME_DEFAULT, help="Template path relative to templates/")
#     parser.add_argument("--serve", action="store_true", help="Run local Flask server for live preview")
#     parser.add_argument("--port", type=int, default=5000, help="Port for --serve")
#     args = parser.parse_args()
#
#     # quick debug print
#     print("TEMPLATE_DIR:", TEMPLATE_DIR)
#     print("OUTPUT_DIR:", OUTPUT_DIR)
#     print("DATA_FILE:", DATA_FILE)
#     print("Using template:", args.template)
#
#     if args.render_only and args.serve:
#         print("Error: --render-only and --serve are mutually exclusive", file=sys.stderr)
#         sys.exit(2)
#
#     # render-only flow: render -> save -> optionally pdf
#     if args.render_only:
#         try:
#             html_str, ctx = render_to_html_string(args.template, load_program_data(args.event_id))
#         except TemplateNotFound as e:
#             print("Template not found in {}: {}".format(TEMPLATE_DIR, e), file=sys.stderr)
#             sys.exit(1)
#         except Exception as e:
#             print("Render failed:", e, file=sys.stderr)
#             traceback.print_exc()
#             sys.exit(1)
#
#         write_output_html(html_str, OUTPUT_HTML)
#         print("Rendered HTML saved to:", OUTPUT_HTML)
#         # attempt PDF via chrome
#         try:
#             chrome_print_pdf(OUTPUT_HTML, OUTPUT_PDF)
#             print("Saved PDF to {}".format(OUTPUT_PDF))
#         except Exception as e:
#             print("Chrome print failed:", e, file=sys.stderr)
#             print("Open the HTML preview to debug:", OUTPUT_HTML, file=sys.stderr)
#             sys.exit(1)
#         return
#
#     # serve flow: run Flask app if Flask is installed
#     if args.serve:
#         if not FLASK_AVAILABLE:
#             print("Flask not installed. Install with: pip install flask", file=sys.stderr)
#             sys.exit(1)
#
#         app = Flask(__name__)
#
#         @app.route("/")
#         def index():
#             try:
#                 html_str, ctx = render_to_html_string(args.template, load_program_data(args.event_id))
#             except TemplateNotFound:
#                 return f"Template {args.template} not found in {TEMPLATE_DIR}", 404
#             except Exception as e:
#                 traceback.print_exc()
#                 return f"Rendering error: {e}", 500
#             return Response(html_str, mimetype="text/html")
#
#         @app.route("/save", methods=["POST", "GET"])
#         def save():
#             """Save current render to output/program.html for debugging."""
#             try:
#                 html_str, ctx = render_to_html_string(args.template, load_program_data(args.event_id))
#                 write_output_html(html_str, OUTPUT_HTML)
#                 return f"Saved HTML to {OUTPUT_HTML}", 200
#             except Exception as e:
#                 traceback.print_exc()
#                 return f"Save failed: {e}", 500
#
#         @app.route("/pdf", methods=["POST", "GET"])
#         def pdf():
#             """Trigger chrome PDF rendering of the current preview (blocks until completion)."""
#             try:
#                 html_str, ctx = render_to_html_string(args.template, load_program_data(args.event_id))
#                 write_output_html(html_str, OUTPUT_HTML)
#                 chrome_print_pdf(OUTPUT_HTML, OUTPUT_PDF)
#                 return f"Saved PDF to {OUTPUT_PDF}", 200
#             except Exception as e:
#                 traceback.print_exc()
#                 return f"PDF failed: {e}", 500
#
#         print(f"Starting Flask dev server on http://127.0.0.1:{args.port} (template: {args.template})")
#         app.run(host="127.0.0.1", port=args.port, debug=True, use_reloader=True)
#         return
#
#     # default behavior (if neither serve nor render-only): do the original render+pdf (safe mode)
#     try:
#         html_str, ctx = render_to_html_string(args.template, load_program_data(args.event_id))
#     except TemplateNotFound as e:
#         print("Template not found in {}: {}".format(TEMPLATE_DIR, e), file=sys.stderr)
#         sys.exit(1)
#     except Exception as e:
#         print("Render failed:", e, file=sys.stderr)
#         traceback.print_exc()
#         sys.exit(1)
#
#     write_output_html(html_str, OUTPUT_HTML)
#     print("Rendered HTML saved to:", OUTPUT_HTML)
#     try:
#         chrome_print_pdf(OUTPUT_HTML, OUTPUT_PDF)
#         print("Saved PDF to {}".format(OUTPUT_PDF))
#     except Exception as e:
#         print("Chrome print failed:", e, file=sys.stderr)
#         print("Open the HTML preview to debug:", OUTPUT_HTML, file=sys.stderr)
#         sys.exit(1)
#
# if __name__ == "__main__":
#     main()
