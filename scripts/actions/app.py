from datetime import datetime, timedelta
import argparse
import json
import sys
import traceback
from pathlib import Path

from flask import Flask, render_template, request, Response
from jinja2 import Undefined
from jinja2.exceptions import TemplateNotFound

# Project layout: this file is scripts/core/app.py
PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"
DATA_FILE = PROJECT_ROOT / "data" / "shared" / "program_data.json"

# make sure project root is on sys.path (helps future imports)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# minimal Flask app pointed at project templates folder and static folder
app = Flask(
    __name__,
    template_folder=str(TEMPLATE_DIR),
    static_folder=str(STATIC_DIR),
)
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
