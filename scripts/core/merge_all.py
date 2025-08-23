# scripts/core/merge_all.py
from __future__ import annotations
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import json, csv, re, argparse
from json import JSONDecodeError
from datetime import datetime
from typing import Any

from scripts.core.bootstrap import BASE_DIR, DATA_DIR, OUTPUT_DIR as BASE_OUTPUT_DIR

CONFIG_SCHEMA_DIR = BASE_DIR / "config" / "schema"
OUTPUT_DIR = BASE_OUTPUT_DIR / "merged"
BACKUP_ROOT = BASE_OUTPUT_DIR / "backups"

def initialize() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_ROOT.mkdir(parents=True, exist_ok=True)

def _read_text(p: Path) -> str:
    return p.read_text(encoding="utf-8")

def _read_json_relaxed(p: Path) -> tuple[object, list[str]]:
    warnings: list[str] = []
    s = _read_text(p)
    try:
        return json.loads(s), warnings
    except JSONDecodeError:
        cleaned = re.sub(r',\s*(?=[}\]])', '', s)
        if cleaned != s:
            try:
                obj = json.loads(cleaned)
                warnings.append("{}: trailing commas removed in-memory; please fix file.".format(p.name))
                return obj, warnings
            except JSONDecodeError:
                raise
        raise

def read_csv(p: Path) -> list[dict]:
    with p.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def try_find_payload(stem: str) -> tuple[str, Path | None]:
    cand = DATA_DIR / ("{}_data.json".format(stem))
    if cand.exists(): return ("json", cand)
    cand = DATA_DIR / ("{}.csv".format(stem))
    if cand.exists(): return ("csv", cand)
    for q in DATA_DIR.rglob("{}_data.json".format(stem)):
        return ("json", q)
    return ("none", None)

def schema_defaults_from(obj: object) -> dict | None:
    # 支援 1) JSON Schema 風格 {properties:{...}} 2) 直接 defaults dict 3) [ {defaults} ]
    if isinstance(obj, dict):
        if "properties" in obj and isinstance(obj["properties"], dict):
            props = obj["properties"]
            return {k: spec.get("default") if isinstance(spec, dict) else None for k, spec in props.items()}
        return obj
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        return obj[0]
    return None

# ====== 新增：遞迴轉型與深合併 ======

def coerce_by_schema(value: Any, schema_default: Any) -> Any:
    # dict：以 schema 為模板遞迴
    if isinstance(schema_default, dict) and isinstance(value, dict):
        out = {}
        for k, sd in schema_default.items():
            v = value.get(k, None)
            out[k] = coerce_by_schema(v, sd)
        # 帶上來源多出的鍵
        for k, v in value.items():
            if k not in out:
                out[k] = v
        return out

    # list：常見情況是來源為 "" 或 逗號分隔字串
    if isinstance(schema_default, list):
        if value is None or (isinstance(value, str) and value.strip() == ""):
            return []
        if isinstance(value, str) and "," in value:
            return [s.strip() for s in value.split(",") if s.strip()]
        # 若 schema_default 提供了元素模板，且 value 是 list of dict，可嘗試遞迴
        if isinstance(value, list) and schema_default and isinstance(schema_default[0], dict):
            return [coerce_by_schema(v, schema_default[0]) if isinstance(v, dict) else v for v in value]
        return value

    # bool
    if isinstance(schema_default, bool):
        if isinstance(value, str):
            lv = value.strip().lower()
            if lv in ("true", "yes", "1"): return True
            if lv in ("false", "no", "0"): return False
        return (bool(value) if value is not None else schema_default)

    # int
    if isinstance(schema_default, int):
        if isinstance(value, str):
            try:
                return int(value.strip() or 0)
            except Exception:
                return schema_default
        return value if isinstance(value, int) else (schema_default if value is None else value)

    # float
    if isinstance(schema_default, float):
        if isinstance(value, str):
            try:
                return float(value.strip() or 0.0)
            except Exception:
                return schema_default
        return value if isinstance(value, (int, float)) else (schema_default if value is None else value)

    # 其餘（含字串 / None / 任意型別）
    return schema_default if value is None else value

def deep_merge(defaults: Any, record: Any) -> Any:
    # 任一方非 dict → 以 record 為準（若 None 則用 defaults）
    if not isinstance(defaults, dict) or not isinstance(record, dict):
        return record if record is not None else defaults
    out = dict(defaults)
    for k, rv in record.items():
        if k in out:
            out[k] = deep_merge(out[k], rv)
        else:
            out[k] = rv
    return out

# ====== 舊函式保留（但不再使用於主流程） ======

def coerce_row_types(row: dict, schema_defaults: dict) -> dict:
    out = dict(row)
    for k, default in schema_defaults.items():
        if k not in out: continue
        v = out[k]
        if v is None or default is None or isinstance(v, type(default)): continue
        try:
            if isinstance(default, bool) and isinstance(v, str):
                lv = v.strip().lower()
                if lv in ("true", "yes", "1"): out[k] = True
                elif lv in ("false", "no", "0"): out[k] = False
            elif isinstance(default, int) and isinstance(v, str):
                out[k] = int(v.strip() or 0)
            elif isinstance(default, float) and isinstance(v, str):
                out[k] = float(v.strip() or 0.0)
            elif isinstance(default, list) and isinstance(v, str) and "," in v:
                out[k] = [s.strip() for s in v.split(",") if s.strip()]
        except Exception:
            pass
    return out

def merge_one(schema_defaults: dict, record: dict) -> dict:
    # 改由深合併實作
    return deep_merge(schema_defaults, record)

def load_records(payload_type: str, payload_path: Path) -> tuple[list[dict], list[str]]:
    warnings: list[str] = []
    if payload_type == "json":
        obj, w = _read_json_relaxed(payload_path)
        warnings += w
        if isinstance(obj, dict):
            for k, val in obj.items():
                if isinstance(val, list) and all(isinstance(x, dict) for x in val):
                    return val, warnings
            return [obj], warnings
        return (obj if isinstance(obj, list) else []), warnings
    elif payload_type == "csv":
        return read_csv(payload_path), warnings
    return [], warnings

def _backup_file(src: Path) -> Path:
    """Backup original file under output/backups/<relative_path>/<name>.<ts>.bak<suffix>"""
    try:
        rel = src.relative_to(BASE_DIR)
    except ValueError:
        rel = src.name  # fallback: flat
    dest_dir = (BACKUP_ROOT / rel).parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = ''.join(src.suffixes)
    stem = src.name[:-len(suffix)] if suffix else src.name
    backup_path = dest_dir / ("{}.{}.bak{}".format(stem, ts, suffix or '.json'))
    backup_path.write_bytes(src.read_bytes())
    return backup_path

def _write_json(p: Path, obj: object) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

def batch_merge(overwrite: bool=False) -> list[tuple[str, str]]:
    initialize()
    report: list[tuple[str, str]] = []
    schema_files = sorted(
        p for p in CONFIG_SCHEMA_DIR.glob("*.json")
        if not (p.name.endswith("_data.json") or p.name.endswith("_merged.json"))
    )
    if not schema_files:
        return [("ALL", "no schema files found")]

    for schema_fp in schema_files:
        name = schema_fp.stem
        try:
            schema_obj, warns = _read_json_relaxed(schema_fp)
            for w in warns: report.append((name, "WARNING: {}".format(w)))
            defaults = schema_defaults_from(schema_obj)
            if not isinstance(defaults, dict):
                report.append((name, "skip (invalid schema format)"))
                continue

            payload_type, payload_path = try_find_payload(name)
            if payload_type == "none":
                out_fp = OUTPUT_DIR / ("{}_merged.json".format(name))
                _write_json(out_fp, [])
                report.append((name, "no payload found -> wrote empty []"))
                continue

            rows, w2 = load_records(payload_type, payload_path)
            for w in w2: report.append((name, "WARNING: {}".format(w)))

            # 新流程：先依 schema 預設遞迴轉型，再深合併
            coerced_rows = [coerce_by_schema(r, defaults) for r in rows]
            merged = [deep_merge(defaults, r) for r in coerced_rows]

            if overwrite:
                if payload_type == "json":
                    b = _backup_file(payload_path)
                    _write_json(payload_path, merged)
                    report.append((name, "OVERWROTE {} [{} rows] (backup: {})".format(payload_path.relative_to(BASE_DIR), len(merged), b.relative_to(BASE_DIR))))
                elif payload_type == "csv":
                    target_json = payload_path.with_name("{}_data.json".format(payload_path.stem))
                    if target_json.exists():
                        b = _backup_file(target_json)
                        note = "(backup: {})".format(b.relative_to(BASE_DIR))
                    else:
                        note = "(new file)"
                    _write_json(target_json, merged)
                    report.append((name, "CSV source -> wrote {} [{} rows] {}".format(target_json.relative_to(BASE_DIR), len(merged), note)))
            else:
                out_fp = OUTPUT_DIR / ("{}_merged.json".format(name))
                _write_json(out_fp, merged)
                report.append((name, "OK ({}) -> {} [{} rows]".format(payload_type, out_fp.relative_to(BASE_DIR), len(merged))))

        except Exception as e:
            report.append((name, "ERROR: {!r}".format(e)))

    return report

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true", help="Overwrite original data file (with auto-backup) instead of writing to output/merged")
    args = parser.parse_args()
    results = batch_merge(overwrite=args.overwrite)
    print("=== Merge Report ===")
    for name, status in results:
        print("- {}: {}".format(name, status))
