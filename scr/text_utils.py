from __future__ import annotations

import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Optional


_ILLEGAL_FILE_CHARS = r"\\/:*?\"<>|"


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def to_half_width(text: str) -> str:
    # 全角转半角
    res = []
    for ch in text:
        code = ord(ch)
        if code == 0x3000:
            res.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            res.append(chr(code - 0xFEE0))
        else:
            res.append(ch)
    return "".join(res)


def canonicalize_header(name: str) -> str:
    if name is None:
        return ""
    s = str(name)
    s = to_half_width(s)
    s = s.replace("\n", " ")
    s = normalize_whitespace(s)
    s = s.lower()
    s = s.replace("(", "").replace(")", "").replace("（", "").replace("）", "")
    s = s.replace(" ", "")
    return s


def sanitize_file_stem(name: str) -> str:
    s = normalize_whitespace(name)
    s = to_half_width(s)
    s = re.sub(f"[{_ILLEGAL_FILE_CHARS}]", "_", s)
    s = s.strip(".")
    return s or "未命名"


def extract_short_name(product_name: str) -> str:
    if not product_name:
        return "未命名"
    m = re.search(r"鼎(.*?)集", product_name)
    if m:
        short = m.group(1)
        short = normalize_whitespace(short)
        if short:
            return short
    return product_name


def to_decimal(value) -> Optional[Decimal]:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return None
    s = to_half_width(s)
    s = s.replace(",", "")
    s = s.replace("元", "")
    s = s.replace(" ", "")
    # 括号负数，如 (123.45)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return None


def quantize_2(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)


def format_amount(value: Decimal) -> str:
    v = quantize_2(value)
    # 千分位格式
    return f"{v:,.2f}"
