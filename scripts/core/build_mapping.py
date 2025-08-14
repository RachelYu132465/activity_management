# 放在檔案最前面
from pathlib import Path
import sys
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"

INVALID_WIN = r'[<>:"/\\|?*\x00-\x1F]'

def sanitize_filename(name: str, max_len: int = 100) -> str:
    """Return a filesystem-safe version of *name* truncated to *max_len* characters."""
    s = (name or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(INVALID_WIN, " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:max_len]

def read_json_relaxed(p: Path) -> Any:
    """Load a JSON file allowing trailing commas and UTF-8 BOM."""
    s = p.read_text(encoding="utf-8")
    if s and s[0] == "\ufeff":
        s = s.lstrip("\ufeff")
    s = re.sub(r",\s*(?=[}\]])", "", s)
    return json.loads(s)

def flatten_list(data: Iterable[Any]) -> List[Dict[str, Any]]:
    """Recursively flatten nested lists of dictionaries."""
    out: List[Dict[str, Any]] = []
    def rec(x: Any) -> None:
        if isinstance(x, dict):
            out.append(x)
        elif isinstance(x, list):
            for y in x:
                rec(y)
    rec(data)
    return out

def load_json(name: str) -> Any:
    """Search for *name* under DATA_DIR and return parsed JSON contents."""
    # 常見位置：data/、data/activities/ 等
    candidates = [
        DATA_DIR / name,
        DATA_DIR / "activities" / name,
        DATA_DIR / "shared" / name,
        ]
    # 兜底：遞迴找
    candidates.extend(DATA_DIR.rglob(name))

    seen = set()
    for cand in candidates:
        if cand in seen:
            continue
        seen.add(cand)
        if cand.exists():
            return read_json_relaxed(cand)
    raise FileNotFoundError(f"找不到 {name}")

def compute_times(
        settings: Dict[str, Any],
        speakers: List[Dict[str, Any]],
) -> Dict[Any, Tuple[str, str]]:
    """Compute start/end times for each speaker and return a lookup table."""
    times: Dict[Any, Tuple[str, str]] = {}
    fmt = "%H:%M"
    cur = datetime.strptime(settings["start_time"], fmt)
    per = int(settings.get("speaker_minutes", 30))

    # 開場 special（after_speaker == 0）
    for s in settings.get("special_sessions", []):
        if int(s.get("after_speaker", -1)) == 0:
            cur += timedelta(minutes=int(s.get("duration") or 0))

    for sp in speakers:
        start, end = cur, cur + timedelta(minutes=per)
        no = sp.get("no")
        nm = sp.get("name")
        times[no] = (start.strftime(fmt), end.strftime(fmt))
        if nm:
            times[nm] = (start.strftime(fmt), end.strftime(fmt))
        cur = end
        # 插入 special
        for ss in settings.get("special_sessions", []):
            if int(ss.get("after_speaker", -1)) == no:
                cur += timedelta(minutes=int(ss.get("duration") or 0))
    return times

def get_event_speaker_mappings(event_name: str) -> List[Dict[str, Any]]:
    """Return a list of merged program/activity/influencer info for *event_name*."""
    programs = load_json("program_data.json")
    activities = load_json("activities_data.json")
    influencers_raw = load_json("influencer_data.json")
    influencers = flatten_list(influencers_raw if isinstance(influencers_raw, list) else [influencers_raw])

    program = next((p for p in programs if event_name in (p.get("eventNames") or [])), None)
    if not program:
        raise ValueError(f"找不到 program: {event_name}")

    activity = next((a for a in activities if event_name in (a.get("eventNames") or [])), None)
    if not activity:
        raise ValueError(f"找不到 activities: {event_name}")

    infl_map: Dict[str, Dict[str, Any]] = {i.get("name"): i for i in influencers if i.get("name")}
    # 以 organization 當備援 key
    for i in influencers:
        org = (i.get("current_position") or {}).get("organization")
        if org and org not in infl_map:
            infl_map[org] = i

    speakers = list(activity.get("speakers") or [])
    settings = dict(activity.get("agenda_settings") or {})
    time_map = compute_times(settings, speakers) if settings and speakers else {}

    results: List[Dict[str, Any]] = []
    for sp in speakers:
        name = sp.get("name") or ""
        inf = infl_map.get(name, {})
        st, et = time_map.get(sp.get("no"), ("", ""))
        if st == "" and name in time_map:
            st, et = time_map[name]

        locations = program.get("locations") or ["", ""]
        location_main = locations[0] if len(locations) > 0 else ""
        location_addr = locations[1] if len(locations) > 1 else ""

        mapping: Dict[str, Any] = {
            "no": sp.get("no"),
            "name": name,
            "topic": sp.get("topic", ""),
            "start_time": st,
            "end_time": et,
            "date": program.get("date", ""),
            "location_main": location_main,
            "location_addr": location_addr,
            **inf,  # 展開 influencer 欄位（含 current_position 等）
        }
        mapping["safe_filename"] = sanitize_filename(name or inf.get("name") or "TBD")
        results.append(mapping)
    return results