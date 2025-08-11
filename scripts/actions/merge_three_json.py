#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
三表合併腳本（強化容錯版）
- Program <eventName(s)> JOIN Activities <eventName(s)>
- Activities speakers[*].name（取空白/換行前）JOIN Influencers name（自動攤平/辨識多種結構）
- 輸出單一 JSON 大表，遞迴刪空欄位

執行（從專案根目錄）：

  python scripts\actions\merge_three_json.py ^
    --program data\shared\program_data.json ^
    --activities data\activities\activities_data.json ^
    --influencers data\shared\influencer_data.json ^
    --out output\merged\combined_table.json

不帶參數也可以（使用預設路徑）。
"""

from __future__ import annotations
import json
import argparse
from pathlib import Path
from typing import Any, Dict, List, Iterable

# ---------- I/O ----------
def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))

def write_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

def deep_copy(obj: Any) -> Any:
    return json.loads(json.dumps(obj, ensure_ascii=False))

# ---------- 清理/工具 ----------
def as_list(x: Any) -> List[Any]:
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]

def norm_event_names(rec: Dict[str, Any]) -> List[str]:
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
    if isinstance(x, dict):
        out = {}
        for k, v in x.items():
            vv = drop_empty(v)
            if vv not in ("", None) and vv != [] and vv != {}:
                out[k] = vv
        return out
    if isinstance(x, list):
        out_list = [drop_empty(v) for v in x]
        out_list = [v for v in out_list if v not in ("", None) and v != [] and v != {}]
        return out_list
    if isinstance(x, str):
        return x.strip()
    return x

# ---------- Influencer 解析（強化版） ----------
def iter_influencer_objects(payload: Any) -> Iterable[Dict[str, Any]]:
    """
    遍歷各種可能結構，輸出「含 name 的 dict」：
    - 直接是 dict 且有 name
    - 直接是 dict（SCHEMA 版），雖然可能其他欄位多
    - list 裡面包 dict / list 再往下找
    - {"speakers": [...]} 裡每個 speaker 若有 name 也抓
    - 純字串/其他型態會略過
    """
    if payload is None:
        return
    if isinstance(payload, dict):
        # 1) 直接含 name 的物件
        if "name" in payload and isinstance(payload["name"], str) and payload["name"].strip():
            yield payload
        # 2) 若有 speakers，從 speakers 取 name
        sp = payload.get("speakers")
        if isinstance(sp, list):
            for s in sp:
                if isinstance(s, dict) and "name" in s:
                    yield s
        # 3) 遍歷其他欄位的子結構
        for v in payload.values():
            if isinstance(v, (dict, list)):
                for obj in iter_influencer_objects(v):
                    yield obj
        return
    if isinstance(payload, list):
        for item in payload:
            for obj in iter_influencer_objects(item):
                yield obj
        return
    # 其他型態不處理

def build_influencer_index(influencer_payload: Any) -> Dict[str, Dict[str, Any]]:
    idx: Dict[str, Dict[str, Any]] = {}
    for obj in iter_influencer_objects(influencer_payload):
        raw = obj.get("name", "")
        key = trim_person_name(raw)
        if key and key not in idx:
            idx[key] = deep_copy(obj)
    return idx

# ---------- Activities by event ----------
def build_activities_by_event(activities_payload: Any) -> Dict[str, List[Dict[str, Any]]]:
    mapping: Dict[str, List[Dict[str, Any]]] = {}

    def visit(node: Any):
        if isinstance(node, dict):
            # 視為活動記錄候選
            evns = norm_event_names(node)
            if evns:
                rec = node
                for evn in evns:
                    mapping.setdefault(evn, []).append(rec)
            # 繼續深入
            for v in node.values():
                if isinstance(v, (dict, list)):
                    visit(v)
        elif isinstance(node, list):
            for it in node:
                visit(it)

    visit(activities_payload)
    return mapping

# ---------- 合併 ----------
def merge_tables(program_payload: Any, activities_payload: Any, influencer_payload: Any) -> List[Dict[str, Any]]:
    infl_idx = build_influencer_index(influencer_payload)
    act_by_event = build_activities_by_event(activities_payload)

    # 取出所有 program 記錄（容許是 list 或包在某層裡）
    program_records: List[Dict[str, Any]] = []

    def collect_program(node: Any):
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
        event_names = norm_event_names(prog)
        if not event_names:
            continue

        for evn in event_names:
            out_item: Dict[str, Any] = {
                "eventName": evn,
                "program": deep_copy(prog),
                "activities": []
            }

            for act in act_by_event.get(evn, []):
                act_copy = deep_copy(act)

                # speakers JOIN influencers
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

                out_item["activities"].append(act_copy)

            out_item = drop_empty(out_item)
            merged.append(out_item)

    return merged

# ---------- CLI ----------
def main():
    BASE_DIR = Path(__file__).resolve().parents[2]

    ap = argparse.ArgumentParser(description="Merge Program / Activities / Influencers into one JSON (drop empty fields).")
    ap.add_argument("--program", default=str(Path("data/shared/program_data.json")))
    ap.add_argument("--activities", default=str(Path("data/activities/activities_data.json")))
    ap.add_argument("--influencers", default=str(Path("data/shared/influencer_data.json")))
    ap.add_argument("--out", default=str(Path("output/merged/combined_table.json")))
    args = ap.parse_args()

    def to_abs(p: str) -> Path:
        pth = Path(p)
        return (BASE_DIR / pth) if not pth.is_absolute() else pth

    program_path     = to_abs(args.program)
    activities_path  = to_abs(args.activities)
    influencers_path = to_abs(args.influencers)
    out_path         = to_abs(args.out)

    program_data     = read_json(program_path)
    activities_data  = read_json(activities_path)
    influencer_data  = read_json(influencers_path)

    merged = merge_tables(program_data, activities_data, influencer_data)
    merged = drop_empty(merged)
    write_json(out_path, merged)

if __name__ == "__main__":
    main()
