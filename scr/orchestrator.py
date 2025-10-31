from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple
from decimal import Decimal
import logging

from scr.io_utils import get_project_paths, ensure_directories, scan_source_files, build_nonconflict_path
from scr.text_utils import extract_short_name, sanitize_file_stem, format_amount
from scr.excel_reader import read_charity_records, read_holding_records
from scr.allocation import allocate_proportional
from scr.docx_utils import render_docx_from_template, attach_detail_table


class DuplicateProductError(Exception):
    pass


def run(root: Path, overwrite: bool, logger: logging.Logger) -> Dict[str, int]:
    paths = get_project_paths(root)
    ensure_directories(paths)

    charity_files, holding_files = scan_source_files(paths["source_dir"])
    logger.info("检测到慈善台账文件: %s 个, 份额文件: %s 个", len(charity_files), len(holding_files))

    if not paths["template_path"].exists():
        raise FileNotFoundError(f"模板文件不存在: {paths['template_path']}")

    product_to_doc: Dict[str, Path] = {}
    product_to_amount: Dict[str, Decimal] = {}
    product_to_short: Dict[str, str] = {}

    # 统计
    stats = {
        "charity_files": len(charity_files),
        "holding_files": len(holding_files),
        "charity_rows": 0,
        "docs_success": 0,
        "docs_failed": 0,
        "detail_files": 0,
        "detail_attach_success": 0,
        "detail_attach_failed": 0,
        "holding_records": 0,
    }

    # 按台账逐行生成代捐说明
    for f in charity_files:
        records = read_charity_records(f, logger)
        stats["charity_rows"] += len(records)
        for rec in records:
            product_name: str = rec["product_name"]
            counterparty: str = rec["counterparty"]
            amount: Decimal = rec["amount"]

            if product_name in product_to_amount:
                raise DuplicateProductError(f"慈善台账发现重复产品: {product_name}")

            short_name = extract_short_name(product_name)
            safe_short = sanitize_file_stem(short_name)
            safe_counterparty = sanitize_file_stem(counterparty) if counterparty else ""
            if safe_counterparty:
                out_doc_stem = f"{safe_counterparty}_代捐说明_{safe_short}"
            else:
                out_doc_stem = f"代捐说明_{safe_short}"
            out_doc_name = f"{out_doc_stem}.docx"
            out_doc_path = build_nonconflict_path(paths["out_doc_dir"] / out_doc_name, overwrite=overwrite)

            placeholders = {
                "#产品名称#": product_name,
                "#对手方#": counterparty,
                "#发生金额#": format_amount(amount),
            }
            try:
                render_docx_from_template(paths["template_path"], out_doc_path, placeholders, logger)
                stats["docs_success"] += 1
            except Exception:
                stats["docs_failed"] += 1
                raise

            product_to_amount[product_name] = amount
            product_to_doc[product_name] = out_doc_path
            product_to_short[product_name] = short_name

    if not product_to_amount:
        logger.warning("未从慈善台账读取到任何产品记录")

    # 计算分摊并生成明细表，同时插入到文档
    import pandas as pd
    detail_headers_docx = [
        "序号",
        "票据抬头\n（实际捐赠人姓名）",
        "票据金额\n（实际捐赠金额（元））",
    ]
    detail_headers_xlsx = [
        "序号",
        "票据抬头（实际捐赠人姓名）",
        "票据金额（实际捐赠金额（元））",
    ]

    for f in holding_files:
        records = read_holding_records(f, logger)
        stats["holding_records"] += len(records)
        if not records:
            continue

        by_product: Dict[str, List[dict]] = {}
        for r in records:
            by_product.setdefault(r["product_name"], []).append(r)

        for product_name, rows in by_product.items():
            amount = product_to_amount.get(product_name)
            if amount is None:
                logger.warning("份额文件中的产品在慈善台账中不存在: %s，该产品分摊金额记为0", product_name)
                amount = Decimal("0")

            shares = [r.get("share") or Decimal("0") for r in rows]
            allocated = allocate_proportional(amount, shares)

            data = []
            total_calc = Decimal("0.00")
            for idx, (r, amt) in enumerate(zip(rows, allocated), start=1):
                total_calc += amt
                data.append([idx, r["customer_name"], format_amount(amt)])
            data.append(["合计", "", format_amount(total_calc)])

            short_name = product_to_short.get(product_name, product_name)
            safe_short = sanitize_file_stem(short_name)
            out_xlsx_name = f"票据明细_{safe_short}.xlsx"
            out_xlsx_path = build_nonconflict_path(paths["out_xlsx_dir"] / out_xlsx_name, overwrite=overwrite)

            df = pd.DataFrame(data, columns=detail_headers_xlsx)
            out_xlsx_path.parent.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(out_xlsx_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            stats["detail_files"] += 1

            docx_path = product_to_doc.get(product_name)
            if docx_path is None:
                logger.warning("未找到对应的代捐说明文档，无法插入明细表: %s", product_name)
                stats["detail_attach_failed"] += 1
            else:
                attach_detail_table(
                    docx_path,
                    "#明细表#",
                    detail_headers_docx,
                    [list(map(str, row)) for row in data],
                    logger,
                )
                stats["detail_attach_success"] += 1

    return stats
