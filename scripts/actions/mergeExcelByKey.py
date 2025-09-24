#!/usr/bin/env python3
"""
merge_by_registration_prefixed.py

Behavior:
- Reads all Excel files in DATA_DIR
- Detect header row by scanning top DETECT_ROWS rows for KEY_KEYWORD
- Keeps key column (報到編號) as-is
- Renames *other* columns to '<filename>__<original_column>' to avoid merging same-named columns across files
- Merges rows by key value (string)
- Merge policy: FIRST_NON_EMPTY_WINS (default) or last-write-wins
- Outputs Excel file
"""

import os
import re
import pandas as pd
from collections import OrderedDict, defaultdict

# ---------------- CONFIG ----------------
DATA_DIR = r"C:\Users\User\Desktop\922"   # <-- set your folder
OUTPUT_XLSX = os.path.join(DATA_DIR, "merged_by_registration_number_prefixed.xlsx")
KEY_KEYWORD = "報到編號"
DETECT_ROWS = 8
SUPPORTED_EXT = (".xlsx", ".xls", ".xlsm")
FIRST_NON_EMPTY_WINS = True   # True: first non-empty wins; False: last-write-wins
# ----------------------------------------

def normalize_for_matching(s):
    if s is None:
        return ""
    t = str(s)
    t = t.replace("\u00A0", " ").replace("　", " ")
    t = t.strip()
    t = re.sub(r'^[\s\(\)\-\.\d\)]+', '', t)
    t = re.sub(r'\s+', ' ', t)
    return t.strip().lower()

def detect_header_row(path, sheet_name=0, detect_rows=DETECT_ROWS):
    try:
        top = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=detect_rows, engine="openpyxl")
    except Exception:
        top = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=detect_rows)
    for i in range(min(len(top), detect_rows)):
        row = top.iloc[i].astype(str).fillna("").tolist()
        normalized = [normalize_for_matching(c) for c in row]
        for nc in normalized:
            if KEY_KEYWORD in nc:
                return i
    return 0

def read_df_with_detected_header(path):
    hr = detect_header_row(path)
    try:
        df = pd.read_excel(path, header=hr, engine="openpyxl")
    except Exception:
        df = pd.read_excel(path, header=hr)
    df.columns = [str(c) for c in df.columns.tolist()]
    return df, hr

def merge_files_prefixed():
    files = sorted([f for f in os.listdir(DATA_DIR) if f.lower().endswith(SUPPORTED_EXT)])
    if not files:
        raise SystemExit(f"No Excel files in {DATA_DIR}")

    # master list of output headers (prefixed where applicable), keep first-seen order
    master_headers = []
    # normalized -> list of prefixed column names (to allow propagation within normalized groups)
    norm_to_prefixed = defaultdict(list)
    # mapping prefixed_col -> norm (for quick lookup)
    prefixed_to_norm = {}

    file_meta = []  # list of dicts with fname, df, orig_cols (after renaming)

    for fname in files:
        path = os.path.join(DATA_DIR, fname)
        try:
            df, hr = read_df_with_detected_header(path)
        except Exception as e:
            print(f"Skipping {fname} (read error): {e}")
            continue

        orig_cols = [str(c) for c in df.columns.tolist()]

        # find key_col in original columns (normalized)
        key_col = None
        for c in orig_cols:
            if KEY_KEYWORD in normalize_for_matching(c):
                key_col = c
                break

        # build rename map: keep key_col unchanged, prefix other cols
        rename_map = {}
        for c in orig_cols:
            if c == key_col:
                newname = c
            else:
                # safe filename prefix: remove path-unfriendly chars from fname
                safe_fname = re.sub(r'[\\/:*?"<>|]', '_', fname)
                newname = f"{safe_fname}__{c}"
            rename_map[c] = newname

            # compute norm from original column label (without prefix)
            norm = normalize_for_matching(c)
            prefixed_to_norm[newname] = norm
            if newname not in master_headers:
                master_headers.append(newname)
            if newname not in norm_to_prefixed[norm]:
                norm_to_prefixed[norm].append(newname)

        # ensure key_col appears in master_headers (kept original name)
        if key_col is not None and key_col not in master_headers:
            master_headers.insert(0, key_col)  # put key first if newly seen

        # rename dataframe columns
        df = df.rename(columns=rename_map)
        # store meta
        file_meta.append({"fname": fname, "df": df, "orig_cols": [str(c) for c in df.columns.tolist()], "key_col": key_col})

    # ensure sourcefiles column at end
    if "sourcefiles" not in master_headers:
        master_headers.append("sourcefiles")

    # master records: key_val -> dict(prefixed_header -> value)
    master = OrderedDict()

    # merge rows
    for meta in file_meta:
        fname = meta["fname"]
        df = meta["df"]
        orig_cols = meta["orig_cols"]
        key_col = meta["key_col"]  # this is original name (not prefixed) or None

        if key_col is None:
            print(f"No key col detected in {fname}; skipping file.")
            continue

        for idx, row in df.iterrows():
            raw_key = row.get(key_col, None)
            if pd.isna(raw_key) or str(raw_key).strip() == "":
                continue
            key_val = str(raw_key).strip()
            if key_val not in master:
                master[key_val] = {h: "" for h in master_headers}
                master[key_val]["sourcefiles"] = fname
            else:
                existing = master[key_val].get("sourcefiles", "")
                if fname not in existing.split(" | "):
                    master[key_val]["sourcefiles"] = existing + (" | " if existing else "") + fname

            # fill values for all columns (prefixed names)
            for c in orig_cols:
                # skip key_col? No — c could equal key_col (original) so we handle it normally
                val = row.get(c, None)
                if pd.isna(val):
                    continue
                sval = str(val).strip()
                if sval == "":
                    continue
                # ensure header exists in master record
                if c not in master[key_val]:
                    master[key_val][c] = ""
                if FIRST_NON_EMPTY_WINS:
                    if master[key_val].get(c, "") == "":
                        master[key_val][c] = sval
                else:
                    master[key_val][c] = sval

    # propagate values across prefixed headers belonging to the same normalized group
    for key_val, rec in master.items():
        for norm, prefixed_list in norm_to_prefixed.items():
            # find first non-empty among prefixed_list
            first_val = None
            for ph in prefixed_list:
                v = rec.get(ph, "")
                if v not in ("", None):
                    first_val = v
                    break
            if first_val is not None:
                for ph in prefixed_list:
                    if rec.get(ph, "") in ("", None):
                        rec[ph] = first_val

    # build DataFrame rows, ensure key is present (prefer the original key header if any)
    # decide preferred key header (first element in master_headers that contains KEY)
    key_candidates = [h for h in master_headers if KEY_KEYWORD in normalize_for_matching(h)]
    preferred_key_header = key_candidates[0] if key_candidates else None

    rows = []
    for key_val, rec in master.items():
        # ensure all columns exist
        for h in master_headers:
            if h not in rec:
                rec[h] = ""
        # put key value in preferred header (or create registration_number)
        if preferred_key_header:
            rec[preferred_key_header] = key_val
        else:
            if "registration_number" not in master_headers:
                master_headers.insert(0, "registration_number")
            rec["registration_number"] = key_val
        rows.append(rec)

    cols_order = [preferred_key_header] + [h for h in master_headers if h != preferred_key_header] if preferred_key_header else master_headers[:]
    merged_df = pd.DataFrame(rows, columns=cols_order)

    merged_df.to_excel(OUTPUT_XLSX, index=False)
    print(f"WROTE: {OUTPUT_XLSX}  (unique keys: {len(master)})")
    with pd.option_context('display.max_columns', None, 'display.width', 240):
        print(merged_df.head(10).to_string(index=False))

if __name__ == "__main__":
    if not os.path.isdir(DATA_DIR):
        raise SystemExit(f"DATA_DIR does not exist: {DATA_DIR}")
    merge_files_prefixed()
