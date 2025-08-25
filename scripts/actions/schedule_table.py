from __future__ import annotations
from typing import Any, Dict, List


def build_table(event: Dict[str, Any]) -> List[Dict[str, str]]:
    """Build table rows from event speakers.

    Each row dictionary contains:
        - type: "host", "talk", or "merge" (others/breaks)
        - time: time string or ""
        - topic / speaker or content depending on type
    """
    rows: List[Dict[str, str]] = []
    speakers = event.get("speakers", []) or []

    for sp in speakers:
        sp_type = sp.get("type", "")
        start = sp.get("start_time") or ""
        end = sp.get("end_time") or ""
        time_str = f"{start}-{end}".strip("-") if start or end else ""
        topic = sp.get("topic", "")
        name = sp.get("name", "")

        if sp_type == "主持人":
            # Host row: merge all columns
            content = "{} {}".format(topic, name).strip()
            rows.append({
                "type": "host",
                "time": time_str,
                "content": content,
            })
        elif sp_type in ("講者", "致詞人"):
            # Normal row: show topic and speaker separately
            rows.append({
                "type": "talk",
                "time": time_str,
                "topic": topic,
                "speaker": name,
            })
        else:
            # Other types treated as break/merged rows
            content = topic
            if name:
                content = f"{topic} {name}".strip()
            rows.append({
                "type": "merge",
                "time": time_str,
                "content": content,
            })
    return rows


if __name__ == "__main__":  # simple debug output
    import json
    from pathlib import Path
    from scripts.core.bootstrap import DATA_DIR

    data_file = DATA_DIR / "shared" / "program_data.json"
    data = json.loads(data_file.read_text(encoding="utf-8"))
    event = data[0] if isinstance(data, list) else data
    tbl = build_table(event)
    print(json.dumps(tbl, ensure_ascii=False, indent=2))
