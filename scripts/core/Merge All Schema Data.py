# scripts/actions/merge_all_schema_data.py
# 目的：不帶任何指令參數，直接自動掃描 config/schema/ 下的所有 schema，
# 在 data/ 內尋找對應資料檔（*_data.json / *.json / *.csv），
# 以 schema 的預設鍵值補齊每一列資料，並保留 data 的額外欄位（不丟）。
# 產出寫到 output/merged/<name>_filled.json，並生成合併報告。
# 加強版：
#  - 嚴格驗證並在「欄位型別不符、資料形狀不符」時立刻停止執行，輸出錯誤報告（含欄位名與列號）。

from __future__ import annotations
import json, csv, sys
from pathlib import Path
from typing import Any, Dict, List, Iterable, Tuple
from datetime import datetime

# === 全域設定 ===
STRICT_STOP_ON_ERROR = True          # 偵測到錯誤（如型別不符）時，立刻停止整個流程
FILL_EMPTY_DEFAULT = False           # 預設不以 default 覆寫空值
CARRY_EXTRA_DEFAULT = True           # 預設保留 data 額外欄位
ALLOW_NUMERIC_COMPAT = True          # 允許 int/float 互通
ALLOW_NUMERIC_STRING = False         # 不把數字字串視為合法

# --- 路徑設定 ---
BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
SCHEMA_DIR = BASE_DIR / "config" / "schema"
OUTPUT_DIR = BASE_DIR / "output"
MERGED_DIR = OUTPUT_DIR / "merged"
REPORTS_DIR = OUTPUT_DIR / "reports"

# --- 初始化 ---
def initialize() -> None:
    for p in (OUTPUT_DIR, MERGED_DIR, REPORTS_DIR):
        p.mkdir(parents=True, exist_ok=True)

# --- 檔案搜尋 ---
def rglob_one(base: Path, name: str) -> Path | None:
    for p in base.rglob("*"):
        if p.name == name:
            return p
    return None

PAIRING_RULES: Dict[str, List[str]] = {
    "program": ["program_data.json", "programs.json", "program.json", "programs.csv", "program.csv"],
    "influencer": ["influencer_data.json", "influencers.json", "influencer.json", "influencers.csv", "influencer.csv"],
    "follower": ["follower_data.json", "followers.json", "follower.json", "followers.csv", "follower.csv"],
    "event_contacts": ["event_contacts.json", "event_contacts_data.json", "event_contacts.csv"],
}
PATTERNS = [
    "{b}_data.json", "{b}.json", "{b}s.json", "{b}_data.csv", "{b}.csv", "{b}s.csv"
]

# --- 載入器 ---
def load_json_any(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))

def load_json_as_rows(path: Path) -> List[Dict[str, Any]]:
    obj = load_json_any(path)
    if isinstance(obj, list):
        if obj and not isinstance(obj[0], dict):
            raise ValueError(f"JSON array 必須是物件陣列: {path}")
        return obj
    if isinstance(obj, dict):
        return [obj]
    raise ValueError(f"不支援的 JSON 形狀: {path}")

def load_csv_as_rows(path: Path) -> List[Dict[str, Any]]:
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def load_rows_by_ext(path: Path) -> List[Dict[str, Any]]:
    if path.suffix.lower() == ".csv":
        return load_csv_as_rows(path)
    return load_json_as_rows(path)

# --- 小工具 ---
def is_empty(v: Any) -> bool:
    return v is None or v == "" or v == [] or v == {}

def to_json(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2)

# 型別相容性
NUMERIC_TYPES = (int, float)

def same_type_or_compatible(default_val: Any, data_val: Any) -> bool:
    td, tv = type(default_val), type(data_val)
    if td is tv:
        return True
    if ALLOW_NUMERIC_COMPAT and (td in NUMERIC_TYPES) and (tv in NUMERIC_TYPES):
        return True
    if ALLOW_NUMERIC_STRING and (td in NUMERIC_TYPES) and isinstance(data_val, str):
        s = data_val.strip()
        if s.replace("_", "").replace(",", "").replace(".", "", 1).lstrip("+-").isdigit():
            return True
    return False

# --- 驗證錯誤類型 ---
class MergeValidationError(Exception):
    pass

# --- 合併與驗證 ---
def merge_row(schema_defaults: Dict[str, Any], row: Dict[str, Any],
              fill_empty: bool = FILL_EMPTY_DEFAULT,
              carry_extra: bool = CARRY_EXTRA_DEFAULT,
              row_index: int | None = None,
              schema_name: str | None = None,
              data_rel: str | None = None,
              errors: List[str] | None = None) -> Dict[str, Any]:
    if errors is None:
        errors = []
    if not isinstance(row, dict):
        errors.append(f"E000 非物件列 row_index={row_index}")
        return row

    merged: Dict[str, Any] = {}
    for k, default in schema_defaults.items():
        if k in row:
            v = row[k]
            if not is_empty(v) and not same_type_or_compatible(default, v):
                dt, vt = type(default).__name__, type(v).__name__
                errors.append(
                    f"E001 型別不符 field='{k}' row={row_index} expected={dt} got={vt} value={repr(v)}"
                )
            merged[k] = (default if (fill_empty and is_empty(v)) else v)
        else:
            merged[k] = default
    if carry_extra:
        for k, v in row.items():
            if k not in schema_defaults:
                merged[k] = v
    return merged

# --- 匹配資料檔 ---
def gen_candidates(base_name: str) -> List[Path]:
    candidates: List[Path] = []
    for name in PAIRING_RULES.get(base_name, []):
        p = rglob_one(DATA_DIR, name)
        if p:
            candidates.append(p)
    for pat in PATTERNS:
        name = pat.format(b=base_name)
        p = rglob_one(DATA_DIR, name)
        if p and p not in candidates:
            candidates.append(p)
    for p in DATA_DIR.rglob("*"):
        if p.is_file() and base_name in p.stem and p.suffix.lower() in {".json", ".csv"} and p not in candidates:
            candidates.append(p)
    return candidates

def score_candidate(path: Path, schema_keys: Iterable[str]) -> int:
    try:
        rows = load_rows_by_ext(path)
        if not rows:
            return 0
        sample_keys = set(rows[0].keys())
        return len(sample_keys.intersection(schema_keys))
    except Exception:
        return -1

def pick_best_candidate(base_name: str, schema: Dict[str, Any]) -> Path | None:
    cands = gen_candidates(base_name)
    if not cands:
        return None
    schema_keys = set(schema.keys())
    scored: List[Tuple[int, Path]] = [(score_candidate(p, schema_keys), p) for p in cands]
    scored.sort(key=lambda x: x[0], reverse=True)
    best = scored[0]
    return best[1] if best[0] >= 0 else None

# --- 報告工具 ---
def write_error_report(schema_rel: str, data_rel: str | None, messages: List[str]) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = REPORTS_DIR / f"merge_error_{ts}.md"
    lines = [
        f"# Merge Error Report ({ts})",
        f"schema: {schema_rel}",
        f"data:   {data_rel}",
        "",
        "**Errors:**",
    ]
    lines.extend([f"- {m}" for m in messages])
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path

# --- 主流程 ---
def merge_one_schema(schema_path: Path, fill_empty: bool = FILL_EMPTY_DEFAULT, carry_extra: bool = CARRY_EXTRA_DEFAULT) -> Dict[str, Any]:
    base = schema_path.stem
    schema = load_json_any(schema_path)
    if not isinstance(schema, dict):
        raise MergeValidationError(f"E100 Schema 必須是扁平 dict: {schema_path.name}")

    data_path = pick_best_candidate(base, schema)
    result_info: Dict[str, Any] = {
        "schema": str(schema_path.relative_to(BASE_DIR)),
        "schema_keys": sorted(list(schema.keys())),
        "data": None,
        "output": None,
        "rows": 0,
        "carry_extra": carry_extra,
        "fill_empty": fill_empty,
    }

    if not data_path:
        out_path = MERGED_DIR / f"{base}_filled.json"
        rows = [schema]
        out_path.write_text(to_json(rows), encoding="utf-8")
        result_info.update({
            "data": None,
            "output": str(out_path.relative_to(BASE_DIR)),
            "rows": 1,
        })
        return result_info

    rows = load_rows_by_ext(data_path)
    merged: List[Dict[str, Any]] = []
    errors: List[str] = []

    for idx, r in enumerate(rows, start=1):
        if not isinstance(r, dict):
            errors.append(f"E000 非物件列 row_index={idx}")
            break
        m = merge_row(schema, r, fill_empty=fill_empty, carry_extra=carry_extra,
                      row_index=idx, schema_name=schema_path.name,
                      data_rel=str(data_path.relative_to(BASE_DIR)), errors=errors)
        merged.append(m)
        if STRICT_STOP_ON_ERROR and errors:
            break

    if errors:
        report = write_error_report(
            str(schema_path.relative_to(BASE_DIR)),
            str(data_path.relative_to(BASE_DIR)),
            errors,
        )
        print(f"❌ 驗證失敗（已停止）：{schema_path.name} | 報告：{report}")
        raise MergeValidationError("; ".join(errors))

    out_path = MERGED_DIR / f"{base}_filled.json"
    out_path.write_text(to_json(merged), encoding="utf-8")

    result_info.update({
        "data": str(data_path.relative_to(BASE_DIR)),
        "output": str(out_path.relative_to(BASE_DIR)),
        "rows": len(rows),
    })
    return result_info

def run(fill_empty: bool = FILL_EMPTY_DEFAULT, carry_extra: bool = CARRY_EXTRA_DEFAULT) -> None:
    initialize()
    if not SCHEMA_DIR.exists():
        raise FileNotFoundError(f"找不到 schema 目錄: {SCHEMA_DIR}")

    results: List[Dict[str, Any]] = []
    try:
        for schema_path in sorted(SCHEMA_DIR.glob("*.json")):
            info = merge_one_schema(schema_path, fill_empty=fill_empty, carry_extra=carry_extra)
            results.append(info)
            print(f"✅ {schema_path.name} 合併完成 → {info['output']}")
    except MergeValidationError as e:
        print(f"\n⛔ 已停止：{e}")
        sys.exit(2)

    # 成功的情況寫總報告
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = REPORTS_DIR / f"merge_report_{ts}.md"
    lines: List[str] = [
        f"# Merge Report ({ts})",
        f"BASE_DIR: {BASE_DIR}",
        "",
    ]
    for info in results:
        lines.append(f"- {info['schema']} → {info['output']} (rows={info['rows']})")
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"📄 報告：{report_path}")

# --- 入口點 ---
if __name__ == "__main__":
    run(fill_empty=FILL_EMPTY_DEFAULT, carry_extra=CARRY_EXTRA_DEFAULT)
