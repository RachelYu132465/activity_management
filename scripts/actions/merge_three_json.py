#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Merge Program/Activities/Influencers into programs with embedded activities.

Usage (from repository root):
  python scripts/actions/merge_three_json.py \
      --program data/shared/program_data.json \
      --activities data/activities/activities_data.json \
      --influencers data/shared/influencer_data.json \
      --out output/merged/program_with_activities.json

If arguments are omitted, the defaults shown above are used.
"""

from __future__ import annotations
import argparse
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List

# ---------- I/O helpers ----------

def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def deep_copy(obj: Any) -> Any:
    return json.loads(json.dumps(obj, ensure_ascii=False))


# ---------- generic utilities ----------

def as_list(x: Any) -> List[Any]:
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]


def norm_event_names(rec: Dict[str, Any]) -> List[str]:
    """Normalize event names within *rec* to a clean list of strings."""
    if not isinstance(rec, dict):
        return []
    if rec.get("eventNames"):
        return [str(s).strip() for s in as_list(rec["eventNames"]) if str(s).strip()]
    if rec.get("eventName"):
        s = str(rec["eventName"]).strip()
        return [s] if s else []
    return []


def trim_person_name(raw: Any) -> str:
    if not isinstance(raw, str):
        return ""
    s = raw.split("\n", 1)[0]
    s = s.split(" ", 1)[0]
    return s.strip()


def drop_empty(x: Any) -> Any:
    """Recursively drop empty strings/collections from *x*."""
    if isinstance(x, dict):
        out: Dict[str, Any] = {}
        for k, v in x.items():
            vv = drop_empty(v)
            if vv not in ("", None) and vv != [] and vv != {}:
                out[k] = vv
        return out
    if isinstance(x, list):
        out_list = [drop_empty(v) for v in x]
        return [v for v in out_list if v not in ("", None) and v != [] and v != {}]
    if isinstance(x, str):
        return x.strip()
    return x


# ---------- Influencer helpers ----------

def iter_influencer_objects(payload: Any) -> Iterable[Dict[str, Any]]:
    """Yield any nested dicts containing a ``name`` field."""
    if payload is None:
        return
    if isinstance(payload, dict):
        if isinstance(payload.get("name"), str) and payload["name"].strip():
            yield payload
        sp = payload.get("speakers")
        if isinstance(sp, list):
            for s in sp:
                if isinstance(s, dict) and "name" in s:
                    yield s
        for v in payload.values():
            if isinstance(v, (dict, list)):
                yield from iter_influencer_objects(v)
        return
    if isinstance(payload, list):
        for item in payload:
            yield from iter_influencer_objects(item)
        return
    # ignore other types


def build_influencer_index(payload: Any) -> Dict[str, Dict[str, Any]]:
    idx: Dict[str, Dict[str, Any]] = {}
    for obj in iter_influencer_objects(payload):
        key = trim_person_name(obj.get("name", ""))
        if key and key not in idx:
            idx[key] = deep_copy(obj)
    return idx


# ---------- Activities helpers ----------

def build_activities_by_event(activities_payload: Any) -> Dict[str, List[Dict[str, Any]]]:
    mapping: Dict[str, List[Dict[str, Any]]] = {}

    def visit(node: Any) -> None:
        if isinstance(node, dict):
            evns = norm_event_names(node)
            if evns:
                for evn in evns:
                    mapping.setdefault(evn, []).append(node)
            for v in node.values():
                if isinstance(v, (dict, list)):
                    visit(v)
        elif isinstance(node, list):
            for it in node:
                visit(it)

    visit(activities_payload)
    return mapping


# ---------- merge logic ----------

def merge_tables(program_payload: Any, activities_payload: Any, influencer_payload: Any) -> List[Dict[str, Any]]:
    infl_idx = build_influencer_index(influencer_payload)
    act_by_event = build_activities_by_event(activities_payload)

    # collect program records (list or nested)
    program_records: List[Dict[str, Any]] = []

    def collect_program(node: Any) -> None:
        if isinstance(node, dict):
            if norm_event_names(node):
                program_records.append(node)
            for v in node.values():
                if isinstance(v, (dict, list)):
                    collect_program(v)
        elif isinstance(node, list):
            for it in node:
                collect_program(it)

    collect_program(program_payload)

    merged: List[Dict[str, Any]] = []

    for prog in program_records:
        prog_copy = deep_copy(prog)
        prog_copy["activities"] = []
        for evn in norm_event_names(prog):
            for act in act_by_event.get(evn, []):
                act_copy = deep_copy(act)

                sp_list = act_copy.get("speakers")
                if isinstance(sp_list, list):
                    new_speakers: List[Dict[str, Any]] = []
                    for sp in sp_list:
                        if not isinstance(sp, dict):
                            continue
                        sp_copy = deep_copy(sp)
                        clean_name = trim_person_name(sp_copy.get("name", ""))
                        sp_copy["name"] = clean_name
                        inf = infl_idx.get(clean_name)
                        if inf:
                            sp_copy["influencer"] = deep_copy(inf)
                        new_speakers.append(sp_copy)
                    act_copy["speakers"] = new_speakers

                prog_copy["activities"].append(act_copy)

        merged.append(drop_empty(prog_copy))

    return merged


# ---------- CLI ----------

def main() -> None:
    BASE_DIR = Path(__file__).resolve().parents[2]

    ap = argparse.ArgumentParser(description="Merge Program/Activities/Influencers into program records with activities.")
    ap.add_argument("--program", default=str(Path("data/shared/program_data.json")))
    ap.add_argument("--activities", default=str(Path("data/activities/activities_data.json")))
    ap.add_argument("--influencers", default=str(Path("data/shared/influencer_data.json")))
    ap.add_argument("--out", default=str(Path("output/merged/program_with_activities.json")))
    args = ap.parse_args()

    def to_abs(p: str) -> Path:
        pth = Path(p)
        return (BASE_DIR / pth) if not pth.is_absolute() else pth

    program_path = to_abs(args.program)
    activities_path = to_abs(args.activities)
    influencers_path = to_abs(args.influencers)
    out_path = to_abs(args.out)

    program_data = read_json(program_path)
    activities_data = read_json(activities_path)
    influencer_data = read_json(influencers_path)

    merged = merge_tables(program_data, activities_data, influencer_data)
    merged = drop_empty(merged)
    write_json(out_path, merged)


if __name__ == "__main__":
    main()
