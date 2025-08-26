import argparse
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, Response, url_for
from jinja2 import Undefined
from jinja2.exceptions import TemplateNotFound

# pdf rendering is optional; import lazily so the app still runs without it
try:  # pragma: no cover - optional dependency
    import pdfkit  # type: ignore
except Exception:  # pragma: no cover - if pdfkit not installed
    pdfkit = None

# Project layout: this file is scripts/core/app.py
PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"
DATA_FILE = PROJECT_ROOT / "data" / "shared" / "program_data.json"
INFLUENCER_FILE = PROJECT_ROOT / "data" / "shared" / "influencer_data.json"

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


def flatten_influencers(payload):
    """Flatten nested lists of influencer objects into a simple list."""
    result = []

    def _walk(item):
        if isinstance(item, list):
            for sub in item:
                _walk(sub)
        elif isinstance(item, dict):
            result.append(item)

    _walk(payload or [])
    return result


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


def _schedule_from_speakers(program, influencers=None):
    """Build a schedule with merged-column rules.

    - Entries with topic "主持" render as a centered row spanning all columns.
    - If topic is "休息", merge topic and speaker columns.
    - Otherwise keep the three-column layout.
    - Time column includes duration in minutes on a new line.
    - Speaker column includes title and organization pulled from influencer data.
    """

    influencers = influencers or {}

    schedule = []
    for sp in program.get("speakers", []) or []:
        if not isinstance(sp, dict):
            continue

        start = sp.get("start_time") or ""
        end = sp.get("end_time") or ""
        time = ""
        if start or end:
            if start and end:
                try:
                    start_dt = datetime.strptime(start, "%H:%M")
                    end_dt = datetime.strptime(end, "%H:%M")
                    mins = int((end_dt - start_dt).total_seconds() // 60)
                    time = f"{start}-{end}\n({mins}分鐘)"
                except Exception:
                    time = f"{start}-{end}"
            else:
                time = start or end

        topic = sp.get("topic", "")
        name = sp.get("name", "")

        inf = influencers.get(name, {}) if isinstance(influencers, dict) else {}
        title = ""
        org = ""
        if isinstance(inf.get("current_position"), dict):
            title = inf["current_position"].get("title", "")
            org = inf["current_position"].get("organization", "")
        speaker = name
        if title:
            speaker = f"{speaker} {title}".strip()
        if org:
            speaker = f"{speaker}\n{org}"

        if topic == "主持":
            content = f"{topic} {speaker}".strip()
            schedule.append({"type": "host", "content": content})
        elif topic == "休息":
            content = topic if not name or name == topic else f"{topic} {speaker}"
            schedule.append({"type": "break", "time": time, "content": content})
        else:
            schedule.append({
                "type": "talk",
                "time": time,
                "topic": topic,
                "speaker": speaker,
            })

    return schedule


def _format_highest_education(he):
    """Return a single-line string for highest education info."""
    if not isinstance(he, dict):
        return ""
    parts = [
        he.get("school"),
        he.get("department"),
        he.get("degree"),
        he.get("graduation_year"),
    ]
    return " ".join([p for p in parts if p])


def _merge_person(name, influencer_map):
    """Combine base name with influencer details."""
    inf = influencer_map.get(name, {}) if isinstance(influencer_map, dict) else {}
    current = inf.get("current_position") or {}
    photo_url = ""
    # Look for a local static image named after the person
    for ext in [".png", ".jpg", ".jpeg", ".gif", ".webp"]:
        candidate = STATIC_DIR / f"{name}{ext}"
        if candidate.exists():
            photo_url = url_for("static", filename=f"{name}{ext}")
            break
    if not photo_url:
        photo_url = inf.get("photo_url", "")

    return {
        "name": name,
        "title": current.get("title", ""),
        "organization": current.get("organization", ""),
        "highest_education": _format_highest_education(inf.get("highest_education")),
        "experience": inf.get("experience", []) or [],
        "achievements": inf.get("achievements", []) or [],
        "photo_url": photo_url,
    }

def build_safe_context(program, influencer_map=None):
    """Build a simple template context focused on speakers."""
    if not program or not isinstance(program, dict):
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

    context = {}
    for k, v in program.items():
        context[k] = v

    ev = _normalize_event_names(program)
    context["eventNames"] = ev
    context["assets"] = program.get("assets", {}) or {}
    context["program"] = program.get("program", program) or program
    context["title"] = program.get("title", "") or (ev[0] if ev else "")
    context["date"] = program.get("date", "") or ""
    context["locations"] = program.get("locations", []) or []
    context["organizers"] = program.get("organizers", []) or []
    context["contact"] = program.get("contact", "") or ""
    chairs = []
    speakers_list = []
    for sp in program.get("speakers", []) or []:
        if not isinstance(sp, dict):
            continue
        name = sp.get("name", "")
        person = _merge_person(name, influencer_map)
        topic = sp.get("topic", "") or ""
        sp_type = sp.get("type", "") or ""
        if sp_type == "主持人" or topic == "主持":
            chairs.append(person)
        elif ("致詞" not in topic) and ("休息" not in topic) and ("討論" not in topic)  and sp_type not in ("致詞人", "休息"):
            speakers_list.append(person)

    context["speakers"] = speakers_list
    context["chairs"] = chairs
    context["schedule"] = _schedule_from_speakers(program, influencer_map)
    context["_all_keys"] = list(program.keys())
    return context


def get_context_for_event(event_id):
    """Load JSON, select program, and return context dict with all fields expanded."""
    raw = load_json_safe(DATA_FILE)
    program = pick_program_by_id(raw, event_id)

    infl_raw = load_json_safe(INFLUENCER_FILE) or []
    infl_list = flatten_influencers(infl_raw)
    infl_map = {p.get("name"): p for p in infl_list if isinstance(p, dict)}

    ctx = build_safe_context(program, infl_map)
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


@app.route("/pdf")
def pdf_route():
    """Render template as PDF with table of contents using wkhtmltopdf."""
    event_id = request.args.get("event_id", None)
    try:
        event_id = int(event_id) if event_id is not None else None
    except Exception:
        event_id = None

    ctx = get_context_for_event(event_id)
    try:
        html = render_template("template.html", **ctx)
    except TemplateNotFound:
        return f"Template template.html not found in {TEMPLATE_DIR}", 404
    except Exception as e:
        traceback.print_exc()
        return f"Rendering error: {e}", 500

    if pdfkit is None:
        return "pdfkit or wkhtmltopdf not installed", 500

    try:
        options = {"encoding": "utf-8", "toc": ""}
        pdf_bytes = pdfkit.from_string(html, False, options=options)
    except Exception as e:  # pragma: no cover - runtime error
        return f"PDF generation failed: {e}", 500

    headers = {"Content-Disposition": "attachment; filename=report.pdf"}
    return Response(pdf_bytes, mimetype="application/pdf", headers=headers)


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