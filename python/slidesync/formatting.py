"""Turn cell values into display strings.

Supports a practical subset of Excel number formats plus ``date:<pattern>``.
"""

from __future__ import annotations

import datetime as _dt
import re
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Optional

_DATE_TOKENS = [
    ("yyyy", "%Y"),
    ("yy", "%y"),
    ("mmmmm", "%b"),   # Excel's single-initial month; closest strftime equivalent
    ("mmmm", "%B"),
    ("mmm", "%b"),
    ("mm", "%m"),
    ("m", "{m}"),      # no zero padding, filled in after strftime
    ("dddd", "%A"),
    ("ddd", "%a"),
    ("dd", "%d"),
    ("d", "{d}"),
    ("hh", "%H"),
    ("h", "{h}"),
    ("ss", "%S"),
    ("s", "{s}"),
    ("am/pm", "%p"),
    ("a/p", "%p"),
]
_MINUTE_TOKENS = {"mm": "%M", "m": "{M}"}


def _excel_date_pattern(pattern: str):
    """Translate an Excel date pattern into (strftime pattern, needs_unpadded).

    Tokenised left to right so single-letter ``m``/``d`` work, and so literal
    text is never partially rewritten. ``m`` means minutes directly after an
    hour token or before a seconds token, and months everywhere else.
    """
    out = []
    i = 0
    n = len(pattern)
    last_was_hour = False
    while i < n:
        ch = pattern[i]
        if ch in "[\"'":  # literal section, copy verbatim
            close = {"[": "]", '"': '"', "'": "'"}[ch]
            j = pattern.find(close, i + 1)
            j = n if j < 0 else j
            out.append(pattern[i + 1:j].replace("%", "%%"))
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            out.append(pattern[i + 1].replace("%", "%%"))
            i += 2
            continue
        low = pattern[i:].lower()
        # minutes: m/mm straight after an hour token, or immediately before seconds
        if low.startswith("m") and (last_was_hour or re.match(r"m{1,2}\s*:?\s*s", low)):
            tok = "mm" if low.startswith("mm") else "m"
            out.append(_MINUTE_TOKENS[tok])
            i += len(tok)
            last_was_hour = False
            continue
        for src, dst in _DATE_TOKENS:
            if low.startswith(src):
                out.append(dst)
                last_was_hour = src.startswith("h")
                i += len(src)
                break
        else:
            out.append(ch.replace("%", "%%"))
            if ch.isalnum():          # a separator keeps hh:mm together
                last_was_hour = False
            i += 1
    return "".join(out)


def _strftime(value, pattern: str) -> str:
    """strftime with Excel's unpadded ``m``/``d``/``h``/``s`` filled in."""
    translated = _excel_date_pattern(pattern)
    # Decide 12- vs 24-hour from the pattern: after strftime, %p is already "PM".
    twelve_hour = "%p" in translated
    text = value.strftime(translated)
    if "{" not in text:
        return text
    hour = getattr(value, "hour", 0)
    return text.format(
        m=value.month, d=value.day,
        h=((hour % 12) or 12) if twelve_hour else hour,
        s=getattr(value, "second", 0), M=getattr(value, "minute", 0),
    )


def _looks_like_date_format(fmt: str) -> bool:
    return bool(re.search(r"[dmyh]", fmt, re.IGNORECASE)) and not re.search(r"[#0]", fmt)


def _format_number(value: float, fmt: str) -> str:
    """Handle #,##0 / 0.00 / 0.0% / $#,##0 / (…) negatives, first section only."""
    section = fmt.split(";")[0]
    percent = "%" in section
    prefix = ""
    m = re.match(r'^([^#0,.]*)', section)
    if m:
        prefix = m.group(1).replace('"', "").replace("\\", "").replace("_", "").strip()
    core = section[len(m.group(0)):] if m else section
    core = core.replace("%", "").replace('"', "")
    decimals = 0
    if "." in core:
        decimals = len(re.sub(r"[^0#]", "", core.split(".", 1)[1]))
    grouping = "," in core
    v = value * 100 if percent else value
    # Excel rounds half away from zero; Python's format() rounds half to even.
    q = Decimal(str(abs(v))).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    body = f"{q:,.{decimals}f}" if grouping else f"{q:.{decimals}f}"
    if percent:
        body += "%"
    sign = "-" if v < 0 else ""
    return f"{sign}{prefix}{body}"


def format_value(value: Any, fmt: Optional[str] = None, cell_format: Optional[str] = None) -> str:
    """Render *value*.

    ``fmt`` is the explicit token format (``{{ref|fmt}}``); ``cell_format`` is the
    Excel number format of the source cell, used when no explicit format is given.
    """
    if value is None:
        return ""

    chosen = fmt.strip() if fmt else None

    if isinstance(value, (_dt.datetime, _dt.date)):
        if chosen and chosen.lower().startswith("date:"):
            return _strftime(value, chosen[5:])
        if chosen and _looks_like_date_format(chosen):
            return _strftime(value, chosen)
        if cell_format and cell_format != "General" and _looks_like_date_format(cell_format):
            return _strftime(value, cell_format)
        return value.strftime("%d-%b-%Y")

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    if isinstance(value, (int, float)):
        if chosen and not chosen.lower().startswith("date:"):
            return _format_number(float(value), chosen)
        if cell_format and cell_format != "General" and re.search(r"[#0]", cell_format):
            return _format_number(float(value), cell_format)
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, float):
            # ``:g`` caps at 6 significant digits (12345678.9 -> 1.23457e+07);
            # 15 significant digits reproduces the value the way Excel stores it.
            return f"{value:.15g}"
        return str(value)

    return str(value)
