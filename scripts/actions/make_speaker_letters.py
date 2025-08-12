# scripts/actions/make_speaker_letters.py
# pip install python-docx
from __future__ import annotations
import json, re, argparse
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional

from docx import Document

# ---- 專案路徑（優先用你的 bootstrap，失敗則用後備猜測） ----
try:
    from scripts.core.bootstrap import initialize, DATA_DIR, OUTPUT_DIR, TEMPLATE_DIR
except Exception:
    BASE_DIR = Path(__file__).resolve().parents[2]
    DATA_DIR = BASE_DIR / "data"
    OUTPUT_DIR = BASE_DIR / "output"
    TEMPLATE_DIR = BASE_DIR / "templates"
    def initialize(): OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---- JSON 讀取（BOM & 尾逗號 容錯）----
def read_json_relaxed(p: Path) -> Any:
    s = p.read_text(encoding="utf-8")
    if s and s[0] == "\ufeff":  # 去 BOM
        s = s.lstrip("\ufeff")
    s = re.sub(r",\s*(?=[}\]])", "", s)  # 移除尾逗號
    return json.loads(s)

# ---- 找資料檔 ----
def load_programs() -> List[Dict[str, Any]]:
    for cand in [DATA_DIR / "program_data.json", DATA_DIR / "shared" / "program_data.json"]:
        if cand.exists(): return read_json_relaxed(cand)
    raise SystemExit("找不到 program_data.json")

def load_activities() -> List[Dict[str, Any]]:
    for cand in [DATA_DIR / "activities" / "activities_data.json", DATA_DIR / "activities_data.json"]:
        if cand.exists(): return read_json_relaxed(cand)
    raise SystemExit("找不到 activities_data.json")

def flatten_influencer_list(data) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    def rec(x):
        if isinstance(x, dict):
            out.append(x)
        elif isinstance(x, list):
            for y in x: rec(y)
    rec(data)
    return out

def load_influencers() -> List[Dict[str, Any]]:
    for cand in [DATA_DIR / "shared" / "influencer_data.json", DATA_DIR / "influencer_data.json"]:
        if cand.exists():
            raw = read_json_relaxed(cand)
            if isinstance(raw, list):  return flatten_influencer_list(raw)
            if isinstance(raw, dict):  return [raw]
            return []
    raise SystemExit("找不到 influencer_data.json")

# ---- event 對應 ----
def match_event(programs: List[Dict[str, Any]], activities: List[Dict[str, Any]], event_name: str):
    pg = next((x for x in programs if any(event_name == n for n in (x.get("eventNames") or []))), None)
    if not pg:
        # 允許 program 無完全相同但有交集
        for x in programs:
            if set(x.get("eventNames", [])) & {event_name}:
                pg = x; break
    if not pg:
        raise SystemExit(f"program_data 找不到 event：{event_name}")

    act = next((x for x in activities if any(event_name == n for n in (x.get("eventNames") or []))), None)
    if not act:
        for n in pg.get("eventNames", []):
            act = next((x for x in activities if any(n == m for m in (x.get("eventNames") or []))), None)
            if act: break
    if not act:
        raise SystemExit(f"activities_data 找不到 event：{event_name}")

    return pg, act

# ---- 議程時間推算 ----
def compute_speaker_times(activity: Dict[str, Any]) -> Dict[Any, tuple[str, str]]:
    settings = dict(activity.get("agenda_settings") or {})
    speakers = list(activity.get("speakers") or [])
    if not settings or not speakers:
        return {}
    fmt = "%H:%M"
    cur = datetime.strptime(settings["start_time"], fmt)
    per = int(settings.get("speaker_minutes", 30))

    # 開場 special
    for s in (settings.get("special_sessions") or []):
        if int(s.get("after_speaker", -1)) == 0:
            dur = int(s.get("duration") or 0)
            cur = cur + timedelta(minutes=dur)

    times: Dict[Any, tuple[str, str]] = {}
    for sp in speakers:
        start = cur
        end = cur + timedelta(minutes=per)
        key_no = int(sp.get("no", 0))
        key_nm = (sp.get("name") or "").strip()
        times[key_no] = (start.strftime(fmt), end.strftime(fmt))
        if key_nm:
            times[key_nm] = (start.strftime(fmt), end.strftime(fmt))
        cur = end
        # 插入 special
        for s in (settings.get("special_sessions") or []):
            if int(s.get("after_speaker", -1)) == key_no:
                dur = int(s.get("duration") or 0)
                cur = cur + timedelta(minutes=dur)
    return times

# ---- 講師索引：name 為主、organization 為備援 ----
def build_influencer_map(influencers: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    mp: Dict[str, Dict[str, Any]] = {}
    for it in influencers:
        if not isinstance(it, dict): continue
        nm = (it.get("name") or "").strip()
        if nm: mp[nm] = it
        org = ((it.get("current_position") or {}).get("organization") or "").strip()
        if org and org not in mp:
            mp[org] = it
    return mp

# ---- 模板尋找（支援子資料夾）----
def find_template_file(template_filename: str) -> Path:
    p = TEMPLATE_DIR / template_filename
    if p.exists(): return p
    matches = list(TEMPLATE_DIR.rglob(template_filename))
    if matches: return matches[0]
    raise SystemExit(f"找不到模板：{template_filename}（已搜尋 {TEMPLATE_DIR} 子資料夾）")

# ---- DOCX 置換 ----
REPLACERS = [
    ("{{name}}", "name"),
    ("{{topic}}", "topic"),
    ("{{start_time}}", "start_time"),
    ("{{end_time}}", "end_time"),
    ("{{date}}", "date"),
    ("{{location_main}}", "location_main"),
    ("{{location_addr}}", "location_addr"),
    ("{{organization}}", "organization"),
    ("{{title}}", "title"),
    # 相容你的舊模板標記
    ("{{ activities_data.speakers. name}}", "name"),
    ("{{ activities_data.speakers. topic}}", "topic"),
    ("{{ activities_data.speakers. starttime}}", "start_time"),
    ("{{ activities_data.speakers. endtime}}", "end_time"),
    ("{{ program_data.. date}}", "date"),
    ("{{locations[0] }}", "location_main"),
    ("{{locations[1] }}", "location_addr"),
]

def render_docx(template_path: Path, out_path: Path, mapping: Dict[str, str]) -> None:
    doc = Document(str(template_path))

    def _apply(text: str) -> str:
        s = text
        for pat, key in REPLACERS:
            s = s.replace(pat, str(mapping.get(key, "")))
        return s

    # 段落
    for p in doc.paragraphs:
        if "{{" in p.text and "}}" in p.text:
            new = _apply(p.text)
            if new != p.text:
                for r in p.runs: r.text = ""
                p.add_run(new)

    # 表格
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "{{" in p.text and "}}" in p.text:
                        new = _apply(p.text)
                        if new != p.text:
                            for r in p.runs: r.text = ""
                            p.add_run(new)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

# ---- 主流程 ----
def make_letters(event_name: str, template_filename: str,
                 out_dir: Optional[Path]=None,
                 filter_speaker_no: Optional[int]=None,
                 filter_speaker_name: Optional[str]=None) -> List[Path]:
    initialize()
    programs    = load_programs()
    activities  = load_activities()
    influencers = load_influencers()

    program, activity = match_event(programs, activities, event_name)
    time_map = compute_speaker_times(activity)
    infl_map = build_influencer_map(influencers)

    date = program.get("date", "")
    locs = program.get("locations") or []
    location_main = (locs[0] if len(locs) > 0 else "")
    location_addr = (locs[1] if len(locs) > 1 else "")

    template_path = find_template_file(template_filename)

    speakers: List[Dict[str, Any]] = list(activity.get("speakers") or [])
    if filter_speaker_no is not None:
        speakers = [s for s in speakers if int(s.get("no", -1)) == int(filter_speaker_no)]
    if filter_speaker_name:
        target = filter_speaker_name.strip()
        speakers = [s for s in speakers if (s.get("name") or "").strip() == target]

    out_base = out_dir or (OUTPUT_DIR / "letters")
    results: List[Path] = []

    for sp in speakers:
        no    = int(sp.get("no", 0))
        name  = (sp.get("name") or "").strip()
        topic = (sp.get("topic") or "").strip()

        st, et = "", ""
        if no in time_map:   st, et = time_map[no]
        elif name in time_map: st, et = time_map[name]

        org, title = "", ""
        if name in infl_map:
            pos = infl_map[name].get("current_position") or {}
            org   = (pos.get("organization") or "")
            title = (pos.get("title") or "")

        mapping = {
            "name": name,
            "topic": topic,
            "start_time": st,
            "end_time": et,
            "date": date,
            "location_main": location_main,
            "location_addr": location_addr,
            "organization": org,
            "title": title,
        }

        safe_name = name or "TBD"
        out_name  = f"{no:02d}_{safe_name}_敬請協助提供CV與簡報.docx"
        out_path  = out_base / out_name
        render_docx(template_path, out_path, mapping)
        results.append(out_path)

    return results

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="依 eventName 產出每位講者的《敬請協助提供CV與簡報》信件（Word）")
    ap.add_argument("--event", required=True, help="event name（與 program/activities 的 eventNames 任一相符即可）")
    ap.add_argument("--template", default="敬請協助提供CV與簡報.docx", help="模板檔名（放在 templates/ 或其子資料夾）")
    ap.add_argument("--outdir", default=None, help="輸出資料夾（預設 output/letters）")
    ap.add_argument("--speaker-no", type=int, default=None, help="（選配）只產出此講者編號")
    ap.add_argument("--speaker-name", default=None, help="（選配）只產出此講者姓名")
    args = ap.parse_args()

    outdir = Path(args.outdir) if args.outdir else None
    files = make_letters(args.event, args.template, outdir, args.speaker_no, args.speaker_name)
    print("=== DONE ===")
    for p in files:
        try:
            print("-", p.relative_to(Path(__file__).resolve().parents[2]))
        except Exception:
            print("-", p)
