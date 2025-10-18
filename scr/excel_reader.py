from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple
import logging

import pandas as pd

from scr.text_utils import canonicalize_header, to_decimal, normalize_whitespace


# 列名同义映射（标准名 -> 备选列名）
CHARITY_REQUIRED = {
    "product_name": {"产品名称", "产品", "产品全称"},
    "counterparty": {"对手方", "对手方名称"},
    "amount": {"发生金额", "金额", "捐赠金额"},
}

HOLDING_REQUIRED = {
    "product_name": {"产品名称", "产品", "产品全称"},
    "customer_name": {"客户名称", "客户", "持有人名称", "投资者名称"},
    "share": {"当前份额", "份额", "持有份额"},
}


def _engine_for_suffix(suffix: str) -> str:
    if suffix.lower() == ".xlsx":
        return "openpyxl"
    if suffix.lower() == ".xls":
        return "xlrd"
    return "openpyxl"


def _canonical_columns(cols: List[str]) -> List[str]:
    return [canonicalize_header(c) for c in cols]


def _build_column_mapping(df: pd.DataFrame, required: Dict[str, set]) -> Optional[Dict[str, str]]:
    """生成列映射：标准字段名 -> 实际列名；若不满足则返回 None。"""
    canonical_cols = {canonicalize_header(c): c for c in df.columns}
    mapping: Dict[str, str] = {}
    for field, candidates in required.items():
        found = None
        for cand in candidates:
            key = canonicalize_header(cand)
            if key in canonical_cols:
                found = canonical_cols[key]
                break
        if not found:
            return None
        mapping[field] = found
    return mapping


def _try_read_with_header(path: Path, header_row: int) -> pd.DataFrame:
    engine = _engine_for_suffix(path.suffix)
    return pd.read_excel(path, engine=engine, header=header_row)


def _detect_header_and_read(path: Path, required: Dict[str, set], logger: logging.Logger) -> Tuple[pd.DataFrame, Dict[str, str]]:
    # 步骤1：先尝试第2行作为表头（0 基下标 1）
    try:
        df = _try_read_with_header(path, 1)
        mapping = _build_column_mapping(df, required)
        if mapping is not None:
            return df, mapping
    except Exception as e:
        logger.warning("默认第2行作为表头解析失败，将尝试探测表头: %s", e)

    # 步骤2：不带表头读取，扫描前 6 行探测表头
    engine = _engine_for_suffix(path.suffix)
    raw = pd.read_excel(path, engine=engine, header=None)
    best_row = None
    best_score = -1
    for r in range(min(6, len(raw))):
        row_values = raw.iloc[r].astype(str).tolist()
        canon = [canonicalize_header(v) for v in row_values]
        score = 0
        for _, cands in required.items():
            if any(canonicalize_header(c) in canon for c in cands):
                score += 1
        if score > best_score:
            best_score = score
            best_row = r

    if best_row is None or best_score <= 0:
        raise ValueError("无法定位表头行，请检查源文件格式")

    df = pd.read_excel(path, engine=engine, header=best_row)
    mapping = _build_column_mapping(df, required)
    if mapping is None:
        raise ValueError("表头已探测，但未能匹配所需列")
    logger.info("表头探测成功，使用第%s行作为表头(1-based)", best_row + 1)
    return df, mapping


def read_charity_records(path: Path, logger: logging.Logger) -> List[Dict[str, object]]:
    df, mapping = _detect_header_and_read(path, CHARITY_REQUIRED, logger)

    records: List[Dict[str, object]] = []
    for _, row in df.iterrows():
        product = normalize_whitespace(str(row.get(mapping["product_name"], "")).strip())
        counterparty = normalize_whitespace(str(row.get(mapping["counterparty"], "")).strip())
        amount = to_decimal(row.get(mapping["amount"]))

        if not product or amount is None:
            continue
        records.append(
            {
                "product_name": product,
                "counterparty": counterparty,
                "amount": amount,
            }
        )
    logger.info("读取慈善台账 %s 条有效记录", len(records))
    return records


def read_holding_records(path: Path, logger: logging.Logger) -> List[Dict[str, object]]:
    df, mapping = _detect_header_and_read(path, HOLDING_REQUIRED, logger)

    records: List[Dict[str, object]] = []
    for _, row in df.iterrows():
        product = normalize_whitespace(str(row.get(mapping["product_name"], "")).strip())
        customer = normalize_whitespace(str(row.get(mapping["customer_name"], "")).strip())
        share = to_decimal(row.get(mapping["share"]))

        if not product or not customer:
            continue
        records.append(
            {
                "product_name": product,
                "customer_name": customer,
                "share": share,
            }
        )
    logger.info("读取份额文件 %s 条记录（含空份额）", len(records))
    return records
