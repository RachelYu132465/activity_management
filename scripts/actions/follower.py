from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import pandas as pd

__all__ = [
    "FollowerRecord",
    "read_followers_from_excel",
    "followers_to_signin_rows",
]


@dataclass
class FollowerRecord:
    """Data row representing a registered follower/attendee."""

    no: str = ""
    registration_time: str = ""
    registration_no: str = ""
    chinese_name: str = ""
    gender: str = ""
    company_department: str = ""
    title: str = ""
    day_phone: str = ""
    mobile_phone: str = ""
    email: str = ""
    subscribe_latest: str = ""
    service_unit_category: str = ""
    occupation_category: str = ""
    referral_source: str = ""
    order_note: str = ""
    system_note: str = ""

    def to_signin_row(self) -> "SignInRow":
        """Convert this record to a :class:`SignInRow` instance."""

        from scripts.actions.signin_table_render import SignInRow

        topic = (self.company_department or "").strip()
        if not topic:
            topic = (self.registration_no or "").strip()
        return SignInRow(
            topic=topic,
            name=(self.chinese_name or "").strip(),
            title=(self.title or "").strip(),
            organization=(self.service_unit_category or "").strip(),
        )


_COLUMN_TO_ATTR = {
    "No": "no",
    "報名時間": "registration_time",
    "報名序號": "registration_no",
    "中文姓名(必填)": "chinese_name",
    "性別": "gender",
    "公司/部門名稱(中文 必填)": "company_department",
    "職稱(中文)": "title",
    "日間聯絡電話(必填)": "day_phone",
    "行動電話": "mobile_phone",
    "電子信箱": "email",
    "是否願意收到最新活動訊息": "subscribe_latest",
    "服務單位類別": "service_unit_category",
    "職業身分類別(可複選)": "occupation_category",
    "您是如何得知活動辦理訊息(可複選)": "referral_source",
    "訂單備註": "order_note",
    "系統備註": "system_note",
}


def _as_str(value) -> str:
    if value is None:
        return ""
    text = str(value)
    return text.strip()


def read_followers_from_excel(path: Path | str, *, sheet: str | int = 0) -> List[FollowerRecord]:
    """Read follower records from the given Excel file."""

    excel_path = Path(path)
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet, dtype=str)
    except TypeError:
        df = pd.read_excel(excel_path, sheet_name=sheet, dtype=str, engine="openpyxl")

    df = df.fillna("")

    missing = [col for col in _COLUMN_TO_ATTR if col not in df.columns]
    if missing:
        missing_str = ", ".join(missing)
        raise ValueError(f"Excel 檔案缺少必要欄位: {missing_str}")

    records: List[FollowerRecord] = []
    for _, row in df.iterrows():
        data = {attr: _as_str(row.get(column)) for column, attr in _COLUMN_TO_ATTR.items()}
        records.append(FollowerRecord(**data))
    return records


def followers_to_signin_rows(followers: Iterable[FollowerRecord]) -> List["SignInRow"]:
    """Helper to map follower records to :class:`SignInRow` values."""

    from scripts.actions.signin_table_render import SignInRow

    rows: List[SignInRow] = []
    for follower in followers:
        rows.append(follower.to_signin_row())
    return rows
