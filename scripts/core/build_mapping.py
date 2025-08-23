from __future__ import annotations
import re
import unicodedata
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple, Callable, Optional

from scripts.core.data_util import read_json_relaxed

# --- minimal, safe bootstrap ---
_THIS = Path(__file__).resolve()
_PARENTS = _THIS.parents
ROOT = _PARENTS[2] if len(_PARENTS) > 2 else _PARENTS[-1]
root_str = str(ROOT)
import sys
if root_str not in sys.path:
    sys.path.insert(0, root_str)

BASE_DIR = ROOT
DATA_DIR = BASE_DIR / "data"
# --- end bootstrap

INVALID_WIN = r'[<>:"/\\|?*\x00-\x1F]'

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def sanitize_filename(name: str, max_len: int = 100) -> str:
    """Return a filesystem-safe version of *name* truncated to *max_len* characters."""
    s = (name or "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(INVALID_WIN, " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:max_len]


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
    candidates = [
        DATA_DIR / name,
        DATA_DIR / "activities" / name,
        DATA_DIR / "shared" / name,
        ]
    # recursive fallback
    try:
        candidates.extend(list(DATA_DIR.rglob(name)))
    except Exception:
        pass

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

    # opening specials (after_speaker == 0)
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
        # insert special sessions after this speaker if any
        for ss in settings.get("special_sessions", []):
            if int(ss.get("after_speaker", -1)) == no:
                cur += timedelta(minutes=int(ss.get("duration") or 0))
    return times


def get_event_speaker_mappings(event_name: str) -> List[Dict[str, Any]]:
    """Return a list of merged program/influencer info for *event_name*."""
    programs = load_json("program_data.json")
    influencers_raw = load_json("influencer_data.json")
    influencers = flatten_list(influencers_raw if isinstance(influencers_raw, list) else [influencers_raw])

    program = next((p for p in programs if event_name in (p.get("eventNames") or [])), None)
    if not program:
        raise ValueError(f"找不到 program: {event_name}")

    infl_map: Dict[str, Dict[str, Any]] = {i.get("name"): i for i in influencers if i.get("name")}
    # fallback: use organization as key
    for i in influencers:
        org = (i.get("current_position") or {}).get("organization")
        if org and org not in infl_map:
            infl_map[org] = i

    speakers = list(program.get("speakers") or [])
    settings = dict(program.get("agenda_settings") or {})

    # first read explicit start/end times from event speaker entries
    time_map: Dict[Any, Tuple[str, str]] = {}
    for sp in speakers:
        st = sp.get("start_time")
        et = sp.get("end_time")
        no = sp.get("no")
        nm = sp.get("name")
        if st and et:
            time_map[no] = (st, et)
            if nm:
                time_map[nm] = (st, et)

    # fill gaps by computing times if settings available
    if settings and any(sp.get("no") not in time_map for sp in speakers):
        computed = compute_times(settings, speakers)
        for sp in speakers:
            no = sp.get("no")
            nm = sp.get("name")
            if no not in time_map:
                st, et = computed.get(no, ("", ""))
                if not st and nm in computed:
                    st, et = computed[nm]
                time_map[no] = (st, et)
                if nm:
                    time_map[nm] = (st, et)
                sp.setdefault("start_time", st)
                sp.setdefault("end_time", et)

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
            **inf,  # expand influencer fields (e.g. current_position)
            "no": sp.get("no"),
            "name": name,
            "topic": sp.get("topic", ""),
            "start_time": st,
            "end_time": et,
            "date": program.get("date", ""),
            "location_main": location_main,
            "location_addr": location_addr,
        }
        mapping["safe_filename"] = sanitize_filename(name or inf.get("name") or "TBD")
        results.append(mapping)
    return results


def _norm_key(s: Optional[str]) -> str:
    """Normalize a string for stable deduplication (NFKC, strip, lower)."""
    if not s:
        return ""
    return unicodedata.normalize("NFKC", str(s)).strip().lower()


def get_program_speaker_mappings(
        program_id_or_eventname: str,
        attach_email: bool = False,
        email_finder: Optional[Callable[[Dict[str, Any]], Optional[str]]] = None,
) -> List[Dict[str, Any]]:
    """
    Return merged speaker mappings for a program.
    - program_id_or_eventname: program id (e.g. "2") or an event name contained in eventNames.
    - attach_email: if True and email_finder provided, call email_finder(record) and set record['email'].
    - email_finder: optional callable(record) -> email str (keeps this module independent from template_utils).
    """
    programs = load_json("program_data.json")
    key = str(program_id_or_eventname).strip()

    # try match by id first
    program = next((p for p in programs if str(p.get("id", "")).strip() == key), None)
    # fallback: try match where eventNames contains provided string (support passing event name)
    if not program:
        program = next((p for p in programs if key in (p.get("eventNames") or [])), None)

    if not program:
        raise ValueError(f"找不到 program: {program_id_or_eventname}")

    event_names = program.get("eventNames") or []
    if not event_names:
        raise ValueError(f"program {program.get('id')} 缺少 eventNames")

    merged: List[Dict[str, Any]] = []
    seen = set()
    for ev in event_names:
        try:
            ev_maps = get_event_speaker_mappings(ev)
        except Exception:
            logging.warning("get_event_speaker_mappings failed for event: %s", ev)
            continue
        for m in ev_maps:
            key = _norm_key(m.get("name") or m.get("safe_filename") or "")
            if not key:
                key = _norm_key(f"{m.get('topic')}-{m.get('no')}")
            if key in seen:
                continue
            seen.add(key)
            m.setdefault("program_data", program)
            if attach_email and email_finder:
                try:
                    m["email"] = email_finder(m)
                except Exception:
                    logging.warning("email_finder failed for %s", m.get("name"))
                    m["email"] = None
            merged.append(m)
    return merged
