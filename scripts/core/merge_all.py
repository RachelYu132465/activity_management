# scripts/core/merge_all.py
from __future__ import annotations
import json, csv, re, argparse
from json import JSONDecodeError
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_SCHEMA_DIR = BASE_DIR / "config" / "schema"
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output" / "merged"
BACKUP_ROOT = BASE_DIR / "output" / "backups"

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
                warnings.append(f"{p.name}: trailing commas removed in-memory; please fix file.")
                return obj, warnings
            except JSONDecodeError:
                raise
        raise

def read_csv(p: Path) -> list[dict]:
    with p.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def try_find_payload(stem: str) -> tuple[str, Path | None]:
    cand = DATA_DIR / f"{stem}_data.json"
    if cand.exists(): return ("json", cand)
    cand = DATA_DIR / f"{stem}.csv"
    if cand.exists(): return ("csv", cand)
    for q in DATA_DIR.rglob(f"{stem}_data.json"):
        return ("json", q)
    return ("none", None)

def schema_defaults_from(obj: object) -> dict | None:
    if isinstance(obj, dict):
        if "properties" in obj and isinstance(obj["properties"], dict):
            props = obj["properties"]
            return {k: spec.get("default") if isinstance(spec, dict) else None for k, spec in props.items()}
        return obj
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        return obj[0]
    return None

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
    merged = dict(schema_defaults)
    merged.update(record)
    return merged

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
    backup_path = dest_dir / f"{stem}.{ts}.bak{suffix or '.json'}"
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
            for w in warns: report.append((name, f"WARNING: {w}"))
            defaults = schema_defaults_from(schema_obj)
            if not isinstance(defaults, dict):
                report.append((name, "skip (invalid schema format)"))
                continue

            payload_type, payload_path = try_find_payload(name)
            if payload_type == "none":
                out_fp = OUTPUT_DIR / f"{name}_merged.json"
                _write_json(out_fp, [])
                report.append((name, "no payload found -> wrote empty []"))
                continue

            rows, w2 = load_records(payload_type, payload_path)
            for w in w2: report.append((name, f"WARNING: {w}"))
            rows = [coerce_row_types(r, defaults) for r in rows]
            merged = [merge_one(defaults, r) for r in rows]

            if overwrite:
                if payload_type == "json":
                    # backup original JSON then overwrite
                    b = _backup_file(payload_path)
                    _write_json(payload_path, merged)
                    report.append((name, f"OVERWROTE {payload_path.relative_to(BASE_DIR)} [{len(merged)} rows] (backup: {b.relative_to(BASE_DIR)})"))
                elif payload_type == "csv":
                    # write/overwrite a sibling _data.json (backup if exists)
                    target_json = payload_path.with_name(f"{payload_path.stem}_data.json")
                    if target_json.exists():
                        b = _backup_file(target_json)
                        note = f"(backup: {b.relative_to(BASE_DIR)})"
                    else:
                        note = "(new file)"
                    _write_json(target_json, merged)
                    report.append((name, f"CSV source -> wrote {target_json.relative_to(BASE_DIR)} [{len(merged)} rows] {note}"))
            else:
                out_fp = OUTPUT_DIR / f"{name}_merged.json"
                _write_json(out_fp, merged)
                report.append((name, f"OK ({payload_type}) -> {out_fp.relative_to(BASE_DIR)} [{len(merged)} rows]"))

        except Exception as e:
            report.append((name, f"ERROR: {e!r}"))

    return report

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true", help="Overwrite original data file (with auto-backup) instead of writing to output/merged")
    args = parser.parse_args()
    results = batch_merge(overwrite=args.overwrite)
    print("=== Merge Report ===")
    for name, status in results:
        print(f"- {name}: {status}")
