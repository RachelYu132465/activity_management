#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_nameplates_foldable_landscape.py

每位講者一張 A4（橫向），上下兩半為相同內容，
但所有文字「字頭朝向中間折線」：上半頁繪製後旋轉 180°，下半頁直接貼上。
列印後對折（top-fold half sheet）— 折疊後文字的頭會靠中線。

此版本：
- 名稱與頭銜共用同一 baseline（底線）
- 半頁寬度分 13 份：name = 10 / title = 3
- 名字大小盡量不變；名字字間距改為「等寬 slot 並在每個 slot 內置中」，同時改善四捨五入誤差分布
- 若 title 超出右側 box，會自動換行（單字換行，若單字仍超出則以字元換行）
"""
from __future__ import annotations
import json
import create_publisher_file
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont

# project root (adjust if needed)
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

# ---------- A4 landscape / layout constants ----------
A4_W = 3508
A4_H = 2480
HALF_H = A4_H // 2

PADDING = 100  # 留白

# base font sizes
BASE_NAME_SIZE = 350
BASE_TITLE_SIZE = 100
BASE_ORG_SIZE = 56

# 微調常數
NAME_OFFSET_X = 30   # 名字水平偏移（整體往右）
NAME_OFFSET_Y = 200   # 名字垂直偏移（往下）
NAME_TITLE_RATIO = (14, 3)  # name : title = 10 : 3 (總份數 13)

# safety limits for negative gap before shrinking font (in fraction of avg char width)
NEGATIVE_GAP_LIMIT_FACTOR = 0.45  # 允許負 gap 至少為 avg_char_width * -0.45
SHOW_ORG = False  # True = 顯示 organization；False = 不顯示（關閉顯示）
LOGO_CAPTION_COLOR = (128, 128, 128)
# font candidates
KAIU_PATH = r"C:\Windows\Fonts\kaiu.ttf"
PROJECT_FONTS_DIR = ROOT / "scripts" / "fonts"
NAME_CANDS = [str(PROJECT_FONTS_DIR / "kaiu.ttf"), KAIU_PATH]
TITLE_CANDS = [str(PROJECT_FONTS_DIR / "kaiu.ttf"), KAIU_PATH]
ORG_CANDS = [str(PROJECT_FONTS_DIR / "kaiu.ttf"), KAIU_PATH]


# logo asset and caption
LOGO_PATH_CANDIDATES = [
    Path(r"C:\\Users\\User\\activity_management\\static logo.png"),
    Path(r"C:\\Users\\User\\activity_management\\static\\logo.png"),
    ROOT / "static" / "logo.png",
    ROOT / "static logo.png",
    ]
LOGO_CAPTION = "台灣醫界聯盟基金會"
LOGO_CAPTION_FONT_SIZE_BASE = 40
LOGO_PADDING = 60
LOGO_MAX_HEIGHT = 220
LOGO_MAX_WIDTH = 320

# ---------- helpers ----------
def load_font_from_src(src: Optional[str], size: int) -> ImageFont.ImageFont:
    if src:
        try:
            return ImageFont.truetype(src, size)
        except Exception:
            pass
    try:
        return ImageFont.truetype(KAIU_PATH, size)
    except Exception:
        return ImageFont.load_default()


def sanitize_filename(s: str) -> str:
    s = str(s or "").strip()
    invalid = '<>:"/\\|?*\n\r\t'
    for ch in invalid:
        s = s.replace(ch, "_")
    s = s.strip(". ").strip()
    return s or "unnamed"

def try_truetype(candidates: List[str], size: int) -> Tuple[ImageFont.ImageFont, Optional[str]]:
    for p in candidates:
        if not p:
            continue
        try:
            f = ImageFont.truetype(p, size)
            print(f"[font] loaded {p} (size={size})")
            return f, p
        except Exception:
            continue
    common = ["DejaVuSans.ttf", "Arial.ttf", "msjh.ttf", "msyh.ttc", "DFKai-SB.ttf", "kaiu.ttf"]
    for n in common:
        try:
            f = ImageFont.truetype(n, size)
            print(f"[font] loaded by name {n} (size={size})")
            return f, n
        except Exception:
            continue
    print("[font WARNING] using ImageFont.load_default() fallback")
    return ImageFont.load_default(), None

def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> Tuple[int,int]:
    if not text:
        return 0,0
    bbox = draw.textbbox((0,0), text, font=font)
    return int(bbox[2]-bbox[0]), int(bbox[3]-bbox[1])

def wrap_text_to_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_w: int) -> List[str]:
    """Wrap text by words to fit max_w; if a word alone is too long, fall back to char wrap."""
    if not text:
        return []
    words = text.split()
    lines: List[str] = []
    cur = ""
    for w in words:
        test = (cur + " " + w).strip()
        if text_size(draw, test, font)[0] <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            # if single word longer than max_w, split by characters
            if text_size(draw, w, font)[0] > max_w:
                # char-split
                part = ""
                for ch in w:
                    if text_size(draw, part + ch, font)[0] <= max_w:
                        part += ch
                    else:
                        if part:
                            lines.append(part)
                        part = ch
                if part:
                    cur = part
                else:
                    cur = ""
            else:
                cur = w
    if cur:
        lines.append(cur)
    return lines

def load_logo_image() -> Tuple[Optional[Image.Image], Optional[Path]]:
    for path in LOGO_PATH_CANDIDATES:
        if not path:
            continue
        try:
            p = Path(path)
        except TypeError:
            continue
        if p.is_file():
            try:
                img = Image.open(p).convert("RGBA")
                print(f"[asset] loaded logo from {p}")
                return img, p
            except Exception as exc:
                print(f"[asset WARNING] 無法讀取 logo ({p}): {exc}")
                continue
    print("[asset WARNING] 找不到 logo 檔案，將略過 logo 與標註")
    return None, None

def add_logo_and_caption(img: Image.Image, logo_rgba: Optional[Image.Image], caption_font_src: Optional[str]):
    if logo_rgba is None:
        return

    draw = ImageDraw.Draw(img)
    _, img_h = img.size

    if logo_rgba.width == 0 or logo_rgba.height == 0:
        return

    scale_factor = min(
        LOGO_MAX_WIDTH / logo_rgba.width,
        LOGO_MAX_HEIGHT / logo_rgba.height,
        1.0,
        )
    new_w = int(logo_rgba.width * scale_factor)
    new_h = int(logo_rgba.height * scale_factor)
    resized_logo = logo_rgba.resize((new_w, new_h), Image.LANCZOS)

    x_logo = LOGO_PADDING
    y_logo = img_h - LOGO_PADDING - new_h

    img.paste(resized_logo, (x_logo, y_logo), resized_logo)

    caption = LOGO_CAPTION
    caption_font_size = max(LOGO_CAPTION_FONT_SIZE_BASE, int(LOGO_CAPTION_FONT_SIZE_BASE * (img_h / 600.0)))
    caption_font = load_font_from_src(caption_font_src, caption_font_size)

    _, text_h = text_size(draw, caption, caption_font)
    x_text = x_logo + new_w + 20
    y_text = y_logo + new_h - text_h -40

    draw.text((x_text, y_text), caption, font=caption_font, fill=LOGO_CAPTION_COLOR)
def draw_name_proportional(draw: ImageDraw.ImageDraw, name: str, font: ImageFont.ImageFont,
                           x_left: int, y_baseline: int, box_w: int) -> Tuple[ImageFont.ImageFont, int, int]:
    """
    改良版：使用等寬 slot，並把每個字在自己的 slot 中水平置中。
    - 若某個字寬超過 slot 寬度，會微幅縮小字型以讓所有字能放入 (盡量不縮放，僅在必要時)
    - 回傳 (used_font, used_total_width, used_height)
    """
    if not name:
        return font, 0, 0

    n = len(name)
    if n == 0:
        return font, 0, 0

    # initial measure
    widths = [text_size(draw, ch, font)[0] for ch in name]
    heights = [text_size(draw, ch, font)[1] for ch in name]
    ch_h = max(heights) if heights else 0
    sum_w = sum(widths)

    # compute slot width (float)
    slot_w = float(box_w) / n

    # if any char wider than slot_w * 0.98, we need to shrink font slightly
    max_char_w = max(widths) if widths else 0
    used_font = font
    if max_char_w > slot_w * 0.98:
        # compute shrink_factor to fit the largest char into 95% of slot
        target = slot_w * 0.95
        shrink_factor = target / max_char_w if max_char_w > 0 else 1.0
        current_size = getattr(font, "size", BASE_NAME_SIZE)
        new_size = max(8, int(current_size * shrink_factor))
        try:
            if hasattr(font, "path"):
                used_font = ImageFont.truetype(font.path, new_size)
            else:
                used_font = ImageFont.truetype(KAIU_PATH, new_size)
        except Exception:
            used_font = font
        # re-measure widths/heights with used_font
        widths = [text_size(draw, ch, used_font)[0] for ch in name]
        heights = [text_size(draw, ch, used_font)[1] for ch in name]
        ch_h = max(heights) if heights else ch_h
        max_char_w = max(widths) if widths else max_char_w

    # Now draw: each char centered in its slot
    # We'll compute float x positions to reduce rounding jitter, and distribute rounding remainder evenly.
    total_slots_w = slot_w * n  # equals box_w (float)
    start_x = float(x_left)

    # Precompute ideal float positions for each char
    positions: List[float] = []
    for i, ch in enumerate(name):
        slot_start = start_x + i * slot_w
        ch_w = widths[i]
        ch_x = slot_start + (slot_w - ch_w) / 2.0
        positions.append(ch_x)

    # To avoid cumulative rounding error, convert to integer positions by distributing remainders:
    int_positions: List[int] = []
    remainders: List[float] = []
    for pos in positions:
        int_pos = int(pos)
        int_positions.append(int_pos)
        remainders.append(pos - int_pos)

    # Distribute leftover pixels based on largest remainders first
    # Compute how many extra pixels we can distribute before exceeding box/right edge
    # We'll compute target_right = x_left + box_w
    target_right = x_left + box_w
    # compute current rightmost if using int_positions
    current_right = int_positions[-1] + widths[-1] if widths else x_left
    # how many pixels we can still push rightwards without exceeding target_right
    available_pixels = int(round(target_right - current_right))
    # If available_pixels > 0, we can add +1 to some int_positions to better center;
    # choose indices with largest remainders (descending)
    if available_pixels > 0:
        idxs = sorted(range(len(remainders)), key=lambda i: remainders[i], reverse=True)
        for j in range(min(available_pixels, len(idxs))):
            k = idxs[j]
            int_positions[k] += 1

    # Final drawing using int_positions
    for i, ch in enumerate(name):
        x_draw = int_positions[i]
        y_draw = y_baseline - ch_h
        draw.text((x_draw, y_draw), ch, font=used_font, fill="black")

    used_total = int(round((int_positions[-1] + widths[-1]) - x_left)) if widths else 0
    return used_font, used_total, ch_h

# Draw a single half content onto an image (half canvas)
# place_near_inner: if True => place content near the inner edge (靠向中線 / fold)
def draw_half_content(img: Image.Image, name: str, org: str,
                      name_font: ImageFont.ImageFont,  org_font: ImageFont.ImageFont,
                      place_near_inner: bool = True):
    draw = ImageDraw.Draw(img)
    w, h = img.size

    # org (top-left)
    ox = PADDING
    oy = PADDING
    org_height_accum = 0
    if SHOW_ORG and org:
        max_w = w - PADDING*2
        words = org.split()
        lines = []
        cur = ""
        for wd in words:
            test = (cur + " " + wd).strip()
            if text_size(draw, test, org_font)[0] <= max_w:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = wd
        if cur:
            lines.append(cur)
        for ln in lines[:3]:
            draw.text((ox, oy), ln, font=org_font, fill="black")
            _, hh = text_size(draw, ln, org_font)
            oy += hh + 8
            org_height_accum += hh + 8

    # top content baseline reference: place below org
    top_content_y = oy + 20

    # content box width (full width minus paddings)
    content_w = w - PADDING * 2
    # split into name / title boxes by ratio
    a, b = NAME_TITLE_RATIO
    total_parts = a + b
    name_box_w = int(content_w * (a / total_parts))
    title_box_w = int(content_w - name_box_w)  # remainder
    # left box x, right box x_right
    x_left = PADDING
    x_right = PADDING + content_w  # right edge of content area

    # baseline location: we choose baseline relative to top_content_y plus a bit (so text near fold)
    name_height = text_size(draw, (name or " "), name_font)[1]
    baseline_y = top_content_y + name_height + NAME_OFFSET_Y

    # draw name in left box (left aligned), using improved equal-slot drawing
    used_name_font, used_name_w, used_name_h = draw_name_proportional(draw, name or "", name_font, x_left, baseline_y, name_box_w)

    # draw title into right box: wrap first to fit title_box_w, then draw lines right-aligned,
    # with last line's bottom aligned to baseline_y (so bottom of last title line shares baseline)
    #
    #     # measure each line height and total height (including small line spacing)
    #     line_heights = [text_size(draw, ln, title_font)[1] for ln in title_lines]
    #     line_spacing = 6
    #     total_title_h = sum(line_heights) + (len(line_heights) - 1) * line_spacing
    #
    #     # y_start so that the bottom of last line aligns to baseline_y
    #     y_start = baseline_y - total_title_h
    #     # draw each line right-aligned inside content_w (x_right is right edge)
    #     for ln_idx, ln in enumerate(title_lines):
    #         tw, th = text_size(draw, ln, title_font)
    #         x = x_right - tw  # right-aligned within content area
    #         draw.text((x, y_start), ln, font=title_font, fill="black")
    #         y_start += th + line_spacing

# ---------- load program ----------
def load_program(program_id: Optional[int]) -> Dict[str,Any]:
    data_file = DATA_DIR / "shared" / "program_data.json"
    programs_raw = json.loads(data_file.read_text(encoding="utf-8"))
    if isinstance(programs_raw, list):
        if program_id is not None:
            for prog in programs_raw:
                try:
                    if int(prog.get("id", -1)) == program_id:
                        return prog
                except Exception:
                    continue
        return programs_raw[0] if programs_raw else {}
    elif isinstance(programs_raw, dict):
        return programs_raw
    return {}

# ---------- main ----------
def main(program_id_raw: str):
    try:
        pid = int(program_id_raw)
    except Exception:
        pid = None

    initialize()

    infl_file = DATA_DIR / "shared" / "influencer_data.json"
    try:
        influencers = json.loads(infl_file.read_text(encoding="utf-8"))
    except Exception:
        influencers = []

    program = load_program(pid)
    if not program:
        print(f"[ERROR] 找不到 program (id={pid})")
        return

    people_tuple = build_people(program, influencers)
    if isinstance(people_tuple, tuple) and len(people_tuple) == 2:
        chairs, speakers = people_tuple
    elif isinstance(people_tuple, list):
        chairs, speakers = [], people_tuple
    else:
        chairs, speakers = [], []

    all_people = []
    for c in (chairs or []):
        c["_role"] = "chair"
        all_people.append(c)
    for s in (speakers or []):
        s["_role"] = "speaker"
        all_people.append(s)

    if not all_people:
        print("[INFO] 沒有任何講者資料")
        return

    out_dir = OUTPUT_DIR / "desk_cards_foldable_landscape"
    out_dir.mkdir(parents=True, exist_ok=True)

    # choose base fonts (nominal sizes) to get usable source names
    _, name_src = try_truetype(NAME_CANDS, BASE_NAME_SIZE)
    _, title_src = try_truetype(TITLE_CANDS, BASE_TITLE_SIZE)
    _, org_src = try_truetype(ORG_CANDS, BASE_ORG_SIZE)

    logo_image, _ = load_logo_image()
    caption_font_src = org_src or title_src or name_src or KAIU_PATH

    created = []
    for idx, person in enumerate(all_people, start=1):
        name = person.get("name", "N/A")

        cur_pos = person.get("short_title")
        org =cur_pos

# canvas A4 landscape
        canvas = Image.new("RGB", (A4_W, A4_H), "white")

        # create two half images (same size)
        half_top = Image.new("RGB", (A4_W, HALF_H), "white")
        half_bottom = Image.new("RGB", (A4_W, HALF_H), "white")

        # scaled font sizes for half
        scale = HALF_H / 700.0
        name_size = max(48, int(BASE_NAME_SIZE * scale))
        title_size = max(18, int(BASE_TITLE_SIZE * scale))
        org_size = max(12, int(BASE_ORG_SIZE * scale))

        def load_font_from_src(src, size):
            if src:
                try:
                    return ImageFont.truetype(src, size)
                except Exception:
                    pass
            try:
                return ImageFont.truetype(KAIU_PATH, size)
            except Exception:
                return ImageFont.load_default()

        name_font = load_font_from_src(name_src, name_size)
        # title_font = load_font_from_src(title_src, title_size)
        org_font = load_font_from_src(org_src, org_size)

        # draw content on both halves
        draw_half_content(half_top, name,  org, name_font,  org_font, place_near_inner=True)
        draw_half_content(half_bottom, name,  org, name_font, org_font, place_near_inner=True)

        # rotate top half 180 degrees so text's head faces the middle fold
        half_top_rot = half_top.rotate(180)
        # add logo + caption to both halves in their final orientation
        add_logo_and_caption(half_top_rot, logo_image, caption_font_src)
        add_logo_and_caption(half_bottom, logo_image, caption_font_src)

        # paste halves into canvas: top = rotated top-half, bottom = (normal) bottom-half
        canvas.paste(half_top_rot, (0,0))
        canvas.paste(half_bottom, (0, HALF_H))

        fname = f"program_{pid or 'unknown'}_card_{idx}_{sanitize_filename(name)}.png"
        out_path = out_dir / fname
        canvas.save(str(out_path), format="PNG", optimize=True)
        created.append(out_path)
        print(f"[ok] saved {out_path}")

    print(f"[DONE] 產生 {len(created)} 張可對折桌牌 (每人一張)。輸出路徑: {out_dir}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pid_arg = sys.argv[1]
    else:
        pid_arg = input("請輸入 Program ID：").strip()
    main(pid_arg)
