#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_simple_nameplates_layout2_full.py

完整可執行腳本 — 針對 Windows 環境（亦可在 Linux 上跑但字型需調整）
功能：
- 讀 program_data.json 與 influencer_data.json
- 呼叫 build_people(program, influencers) 取得 chairs / speakers
- 對每位產生一張 1200x700 PNG 桌牌，版面樣式類似「圖二」：
  - 左上：單位（標楷體或備援字型）
  - 中間偏左：姓名（大標、顯著、粗體）
  - 右側：職稱（標楷體或備援字型），右對齊
- 會嘗試載入多種系統字型、並印出載入結果以便 debug
- 若找不到系統字型會提示並使用 pillow 的 fallback（但效果不佳；建議放入專案 fonts/）
"""

from __future__ import annotations
import json
import sys
import platform
from pathlib import Path
from typing import Any, Dict, List, Tuple

from PIL import Image, ImageDraw, ImageFont

# project root
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# import project bootstrap + build_people
try:
    from scripts.core.bootstrap import DATA_DIR, OUTPUT_DIR, initialize
except Exception as e:
    print("[ERROR] 無法 import scripts.core.bootstrap:", e)
    raise

try:
    from scripts.actions.influencer import build_people
except Exception as e:
    print("[ERROR] 無法 import build_people from scripts.actions.influencer:", e)
    raise

# Card constants
CARD_W = 1200
CARD_H = 700
PADDING = 40

NAME_SIZE = 220
TITLE_SIZE = 48
ORG_SIZE = 44

# --- font candidate lists (use absolute paths found on Windows) ---
# These reflect typical files present on Windows (based on your list).
MSJH_BOLD_CANDIDATES = [
    r"C:\Windows\Fonts\msjhbd.ttc",
    r"C:\Windows\Fonts\msjhbd.ttf",
    r"C:\Windows\Fonts\msjhl.ttc",
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjh.ttf",
]

MSJH_REGULAR_CANDIDATES = [
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjh.ttf",
    r"C:\Windows\Fonts\msyh.ttc",    # 微軟雅黑 (備援)
]

BKAI_CANDIDATES = [
    r"C:\Windows\Fonts\kaiu.ttf",       # 你的系統中存在
    r"C:\Windows\Fonts\STKAITI.TTF",
    r"C:\Windows\Fonts\STXINGKA.TTF",
    r"C:\Windows\Fonts\NotoSansTC-VF.ttf",  # fallback
]

# If you prefer to bundle fonts with the project, add:
PROJECT_FONTS_DIR = ROOT / "scripts" / "fonts"
# You can insert these into candidate lists (uncomment if you copied fonts into project)
# Example:
# MSJH_BOLD_CANDIDATES.insert(0, str(PROJECT_FONTS_DIR / "msjhbd.ttf"))
# BKAI_CANDIDATES.insert(0, str(PROJECT_FONTS_DIR / "DFKai-SB.ttf"))

# ---------------- helpers ----------------
def sanitize_filename(s: str) -> str:
    s = str(s or "").strip()
    invalid = '<>:"/\\|?*\n\r\t'
    for ch in invalid:
        s = s.replace(ch, "_")
    return s

def load_font_with_fallback(candidates: List[str], size: int, friendly_name: str) -> Tuple[ImageFont.FreeTypeFont, str | None]:
    """
    Try to load a TTF/collection from given absolute paths (candidates).
    Then try some common system names that freetype might resolve.
    Returns (font_obj, source_path_or_name_or_None).
    Prints debug info.
    """
    for p in candidates:
        if not p:
            continue
        try:
            pth = Path(p)
            if pth.exists():
                f = ImageFont.truetype(str(pth), size)
                print(f"[font] loaded {friendly_name} from path: {pth}")
                return f, str(pth)
        except Exception as e:
            # continue trying next
            # print(f"[font debug] failed to load {p}: {e}")
            continue

    # try a small list of common names (let freetype / OS try to resolve)
    common_names = ["DejaVuSans.ttf", "Arial.ttf", "LiberationSans-Regular.ttf"]
    if platform.system() == "Windows":
        # include MSJH / DFKai names commonly present
        common_names = ["msjhbd.ttf", "msjh.ttf", "msyh.ttc", "DFKai-SB.ttf", "kaiu.ttf"] + common_names

    for name in common_names:
        try:
            f = ImageFont.truetype(name, size)
            print(f"[font] loaded {friendly_name} by system name: {name}")
            return f, name
        except Exception:
            continue

    # last resort: use default bitmap font (not scalable) and warn
    print(f"[font WARNING] Could not find TTF for {friendly_name}. Candidates tried: {candidates + common_names}")
    print(" -> Using ImageFont.load_default() fallback (small bitmap font; not scalable).")
    return ImageFont.load_default(), None

def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> Tuple[int, int]:
    """Return width,height using textbbox for Pillow compatibility."""
    if not text:
        return 0, 0
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    return int(w), int(h)

def draw_right_aligned(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, right_x: int, y: int, fill="black") -> Tuple[int, int, int, int]:
    w, h = text_size(draw, text, font)
    x = right_x - w
    draw.text((x, y), text, font=font, fill=fill)
    return x, y, w, h

# ---------------- layout rendering ----------------
def generate_card_image_layout2(
        name: str,
        title: str,
        org: str,
        out_path: Path,
        name_font: ImageFont.FreeTypeFont,
        title_font: ImageFont.FreeTypeFont,
        org_font: ImageFont.FreeTypeFont,
):
    """Create the card with the requested layout and save to out_path."""
    img = Image.new("RGB", (CARD_W, CARD_H), "white")
    draw = ImageDraw.Draw(img)

    # --- Draw organization top-left ---
    org_text = (org or "").strip()
    if org_text:
        ox = PADDING
        oy = PADDING
        max_org_w = CARD_W - PADDING * 2 - 200
        # naive wrapping
        org_lines: List[str] = []
        if text_size(draw, org_text, org_font)[0] <= max_org_w:
            org_lines = [org_text]
        else:
            words = org_text.split()
            cur = ""
            for w in words:
                test = (cur + " " + w).strip()
                if text_size(draw, test, org_font)[0] <= max_org_w:
                    cur = test
                else:
                    if cur:
                        org_lines.append(cur)
                    cur = w
            if cur:
                org_lines.append(cur)
        oy_cursor = oy
        for ln in org_lines[:3]:
            draw.text((ox, oy_cursor), ln, font=org_font, fill="black")
            _, hln = text_size(draw, ln, org_font)
            oy_cursor += hln + 6

    # --- Draw Name (big, left-biased area) ---
    name_text = (name or "").strip()
    x_name = PADDING + 40
    right_column_width = 360
    right_column_left = CARD_W - PADDING - right_column_width
    max_name_width = right_column_left - x_name - 20

    # attempt to scale down name font if needed
    trial_font = name_font
    w_name, h_name = text_size(draw, name_text, trial_font)
    if w_name > max_name_width and w_name > 0:
        scale = max_name_width / float(w_name)
        new_size = max(36, int(NAME_SIZE * scale))
        try:
            # If original font was truetype, get its path if possible; else just try to load by name
            # We can't easily recover path from ImageFont object, so just attempt to load same style by common filenames:
            # Try msjhbd first from candidates
            trial_font = ImageFont.truetype(MSJH_BOLD_CANDIDATES[0], new_size) if Path(MSJH_BOLD_CANDIDATES[0]).exists() else ImageFont.truetype(MSJH_BOLD_CANDIDATES[-1], new_size)
            w_name, h_name = text_size(draw, name_text, trial_font)
            name_font = trial_font
        except Exception:
            # fallback: keep original font size (may overflow)
            w_name, h_name = text_size(draw, name_text, name_font)

    y_name = (CARD_H // 2) - (h_name // 2)
    draw.text((x_name, y_name), name_text, font=name_font, fill="black")

    # --- Draw Title on right column, right-aligned, vertically centered ---
    title_text = (title or "").strip()
    if title_text:
        right_x = CARD_W - PADDING
        max_title_w = right_column_width - 20
        # naive wrap
        title_lines: List[str] = []
        if text_size(draw, title_text, title_font)[0] <= max_title_w:
            title_lines = [title_text]
        else:
            words = title_text.split()
            cur = ""
            for w in words:
                test = (cur + " " + w).strip()
                if text_size(draw, test, title_font)[0] <= max_title_w:
                    cur = test
                else:
                    if cur:
                        title_lines.append(cur)
                    cur = w
            if cur:
                title_lines.append(cur)
        total_h = sum(text_size(draw, ln, title_font)[1] + 6 for ln in title_lines) - 6
        y_title_start = (CARD_H // 2) - (total_h // 2)
        y_cursor = y_title_start
        for ln in title_lines:
            draw_right_aligned(draw, ln, title_font, right_x, y_cursor, fill="black")
            _, hln = text_size(draw, ln, title_font)
            y_cursor += hln + 6

    # Save
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(out_path), format="PNG", optimize=True)
    return out_path

# ---------------- program loading ----------------
def load_program(program_id: int | None) -> Dict[str, Any]:
    data_file = DATA_DIR / "shared" / "program_data.json"
    programs_raw = json.loads(data_file.read_text(encoding="utf-8"))
    if isinstance(programs_raw, list):
        if program_id is not None:
            for prog in programs_raw:
                try:
                    if int(prog.get("id", -1)) == program_id:
                        return prog
                except (TypeError, ValueError):
                    continue
        return programs_raw[0] if programs_raw else {}
    elif isinstance(programs_raw, dict):
        return programs_raw
    return {}

# ---------------- main flow ----------------
def main(program_id_raw: str):
    try:
        pid = int(program_id_raw)
    except (TypeError, ValueError):
        pid = None

    initialize()

    # load influencers
    infl_file = DATA_DIR / "shared" / "influencer_data.json"
    try:
        influencers = json.loads(infl_file.read_text(encoding="utf-8"))
    except Exception:
        influencers = []

    program = load_program(pid)
    if not program:
        print(f"[ERROR] 找不到 program (id={pid}) / program not found.")
        return

    people_tuple = build_people(program, influencers)
    if isinstance(people_tuple, tuple) and len(people_tuple) == 2:
        chairs, speakers = people_tuple
    elif isinstance(people_tuple, list):
        chairs, speakers = [], people_tuple
    else:
        chairs, speakers = [], []

    all_people: List[Dict[str, Any]] = []
    for c in (chairs or []):
        c["_role"] = "chair"
        all_people.append(c)
    for s in (speakers or []):
        s["_role"] = "speaker"
        all_people.append(s)

    if not all_people:
        print("[INFO] 沒有 chairs 或 speakers 資料 / no people found.")
        return

    # load fonts (with debug)
    name_font_obj, name_src = load_font_with_fallback(MSJH_BOLD_CANDIDATES, NAME_SIZE, "Name-Bold")
    title_font_obj, title_src = load_font_with_fallback(MSJH_REGULAR_CANDIDATES, TITLE_SIZE, "Title-Regular")
    org_font_obj, org_src = load_font_with_fallback(BKAI_CANDIDATES, ORG_SIZE, "Org-BKAI")

    print(f"[font selected] Name -> {name_src}, Title -> {title_src}, Org -> {org_src}")

    out_dir = OUTPUT_DIR / "desk_cards"
    out_dir.mkdir(parents=True, exist_ok=True)
    created = []

    for i, person in enumerate(all_people, start=1):
        name = person.get("name", "N/A")
        title = person.get("title", "") or ""
        # organization resolution: prefer current_position.organization if present
        org = ""
        cur_pos = person.get("current_position")
        if isinstance(cur_pos, dict):
            org = cur_pos.get("organization") or cur_pos.get("org") or ""
        org = org or person.get("organization") or person.get("affiliation") or ""
        fname = f"program_{pid or 'unknown'}_{i}_{sanitize_filename(name)}.png"
        out_path = out_dir / fname
        print(f"[render] {i}/{len(all_people)} -> {name} | title='{title}' | org='{org}'")
        try:
            generate_card_image_layout2(
                name=name,
                title=title,
                org=org,
                out_path=out_path,
                name_font=name_font_obj,
                title_font=title_font_obj,
                org_font=org_font_obj,
            )
            created.append(out_path)
            print(f"[ok] saved: {out_path}")
        except Exception as e:
            print(f"[ERROR] failed render for {name}: {e}")

    print(f"[DONE] 共產生 {len(created)} 張桌牌。 路徑: {out_dir}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pid_arg = sys.argv[1]
    else:
        pid_arg = input("請輸入 Program ID：").strip()
    main(pid_arg)
