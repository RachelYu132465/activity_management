import json
import csv
from datetime import datetime, timedelta

# ====== 設定基礎路徑 ======


from scripts.core.bootstrap import (
    initialize, load_schema, merge_schema, load_json_file,load_csv_file,
    BASE_DIR, DATA_DIR, OUTPUT_DIR, TEMPLATE_DIR
)
# 初始化環境（建立資料夾等）
initialize()



# ====== 載入 JSON 模板（分類用） ======
with open((TEMPLATE_DIR / "activities" / "session_type.json"), encoding="utf-8") as f:
    SESSION_TYPE_TEMPLATE = json.load(f)



# ====== 載入設定檔與 schema 合併 ======
config_data = load_json_file("agenda_settings.json")
schema = load_schema("agenda_settings.json")
config = merge_schema(schema, [config_data])[0]

# ====== 載入講者資料 CSV，並轉為 list ======
raw_speakers = load_csv_file("speakers.csv")
speakers = [
    {
        "index": int(row["序號"]),
        "title": row["主題"].strip(),
        "speaker_name": row["中文姓名"].strip()
    }
    for row in raw_speakers
]

# ====== 自動分配空白時段 ======
def distribute_empty_durations(config, speaker_count):
    fmt = "%H:%M"
    total_minutes = int((datetime.strptime(config["end_time"], fmt) -
                         datetime.strptime(config["start_time"], fmt)).total_seconds() / 60)

    total_known = speaker_count * config["speaker_minutes"]
    empty_items = []

    for s in config["special_sessions"]:
        if not s.get("duration"):
            empty_items.append(s)
        else:
            total_known += s["duration"]

    remaining = total_minutes - total_known
    if remaining < 0:
        print(f"[警告] 總時數不足，缺 {-remaining} 分鐘")
        return config

    if empty_items:
        avg_time = round(remaining / len(empty_items))
        print(f"[INFO] 平均分配 {remaining} 分鐘，每個空值分配 {avg_time} 分鐘")
        for s in empty_items:
            s["duration"] = avg_time

    return config

config = distribute_empty_durations(config, len(speakers))

# ====== 工具函式 ======
def generate_session_id(event_date: str, index: int) -> str:
    return f"session_{event_date}_{index:02d}"

def classify_session_type(title: str) -> str:
    t = title.strip()
    for mapping in SESSION_TYPE_TEMPLATE.values():
        if t in mapping:
            return mapping[t]
    return "lecture"

def insert_special_sessions(agenda, event_date, current_time, current_index, special_list):
    for s in special_list:
        if s["after_speaker"] == current_index:
            session_type = classify_session_type(s["title"])
            end_time = current_time + timedelta(minutes=s["duration"])
            agenda.append({
                "session_id": generate_session_id(event_date, len(agenda) + 1),
                "start_time": current_time.strftime("%H:%M"),
                "end_time": end_time.strftime("%H:%M"),
                "session_title": s["title"],
                "session_type": session_type
            })
            current_time = end_time
    return current_time

# ====== 主議程生成器 ======
def generate_agenda(event_date: str, config: dict, speakers: list):
    agenda = []
    current_time = datetime.strptime(f"{event_date} {config['start_time']}", "%Y%m%d %H:%M")

    current_time = insert_special_sessions(agenda, event_date, current_time, 0, config["special_sessions"])

    for sp in speakers:
        end_time = current_time + timedelta(minutes=config["speaker_minutes"])
        agenda.append({
            "session_id": generate_session_id(event_date, len(agenda) + 1),
            "start_time": current_time.strftime("%H:%M"),
            "end_time": end_time.strftime("%H:%M"),
            "session_title": sp["title"],
            "speaker_name": sp["speaker_name"],
            "session_type": classify_session_type("演講")
        })
        current_time = end_time

        current_time = insert_special_sessions(agenda, event_date, current_time, sp["index"], config["special_sessions"])

    current_time = insert_special_sessions(agenda, event_date, current_time, 999, config["special_sessions"])

    return agenda

# ====== 主執行區 ======
if __name__ == "__main__":
    event_date = "20250922"
    agenda_list = generate_agenda(event_date, config, speakers)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "agenda.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(agenda_list, f, ensure_ascii=False, indent=2)

    print(f"✅ 議程已產生：{output_path}")
