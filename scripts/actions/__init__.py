from __future__ import annotations
from datetime import datetime
import locale
from typing import Optional

__all__ = ["format_date"]

def _temp_setlocale(name: str):
    """Set locale temporarily and return old locale string (may raise)."""
    old = locale.setlocale(locale.LC_TIME)
    locale.setlocale(locale.LC_TIME, name)
    return old

def format_date(
        date_str: str,
        target_format: str = "%Y-%m-%d",
        chinese_weekday: bool = False,
        no_leading_zero: bool = False,
        locale_name: Optional[str] = None,
        sep: Optional[str] = None,
) -> str:
    """
    Convert an ISO date string ("YYYY-MM-DD") to a formatted string.

    Parameters
    ----------
    date_str : str
        ISO date "YYYY-MM-DD", e.g. "2025-09-03".
    target_format : str
        A strftime format string (e.g. "%Y-%m-%d %A" or "%Y/%m/%d (%a)").
    chinese_weekday : bool
        If True, replace weekday with Chinese "星期X" (stable, does not require locale).
    no_leading_zero : bool
        If True, remove leading zeros from month/day (e.g. "9" instead of "09").
    locale_name : Optional[str]
        If provided, attempt to set locale for strftime (e.g. "zh_TW.UTF-8").
        If locale setting fails, fallback to manual formatting.
    sep : Optional[str]
        If provided, normalize any of the characters "-", "/", "." in the final
        formatted string to this separator. Useful to change "2025-09-03" -> "2025/09/03".

    Returns
    -------
    str
        The formatted date string.
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d")

    # 1) Try system locale if requested (best-effort)
    if locale_name:
        try:
            old = _temp_setlocale(locale_name)
            out = dt.strftime(target_format)
            locale.setlocale(locale.LC_TIME, old)
            # handle chinese_weekday/no_leading_zero below if requested
            if not (chinese_weekday or no_leading_zero or sep):
                return out
            base = out
        except Exception:
            # fallback to manual path
            try:
                locale.setlocale(locale.LC_TIME, old)
            except Exception:
                pass
            base = None
    else:
        base = None

    # 2) Prepare format, avoid %A if chinese_weekday requested
    placeholder = "{WEEKDAY}"
    if chinese_weekday and "%A" in target_format:
        fmt = target_format.replace("%A", placeholder)
    else:
        fmt = target_format

    if base is None:
        base = dt.strftime(fmt)

    # 3) Remove leading zeros if requested (cross-platform)
    if no_leading_zero:
        # replace first occurrence of zero-padded month/day
        base = base.replace(dt.strftime("%m"), str(dt.month), 1)
        base = base.replace(dt.strftime("%d"), str(dt.day), 1)

    # 4) Inject Chinese weekday if requested
    if chinese_weekday:
        cn_map = ["一", "二", "三", "四", "五", "六", "日"]
        weekday_cn = "星期" + cn_map[dt.isoweekday() - 1]
        base = base.replace(placeholder, weekday_cn)

    # 5) Normalize separators if requested (conservative: only replace - / .)
    if sep is not None:
        # replace only the common separators (won't touch other punctuation)
        for ch in ("-", "/", "."):
            if ch != sep:
                base = base.replace(ch, sep)

    return base
