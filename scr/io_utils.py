from __future__ import annotations

from pathlib import Path
from datetime import datetime
import logging
from typing import List, Tuple, Dict


def get_project_paths(root: Path) -> Dict[str, Path]:
    """基于根目录构建并返回项目内各路径。"""
    template_path = root / "Template" / "代捐说明模版文件.docx"
    source_dir = root / "数据源"
    output_root = root / "输出"
    out_doc_dir = output_root / "代捐说明"
    out_xlsx_dir = output_root / "明细表"
    logs_dir = root / "logs"
    return {
        "root": root,
        "template_path": template_path,
        "source_dir": source_dir,
        "output_root": output_root,
        "out_doc_dir": out_doc_dir,
        "out_xlsx_dir": out_xlsx_dir,
        "logs_dir": logs_dir,
    }


def ensure_directories(paths: Dict[str, Path]) -> None:
    paths["output_root"].mkdir(parents=True, exist_ok=True)
    paths["out_doc_dir"].mkdir(parents=True, exist_ok=True)
    paths["out_xlsx_dir"].mkdir(parents=True, exist_ok=True)
    paths["logs_dir"].mkdir(parents=True, exist_ok=True)


def scan_source_files(source_dir: Path) -> Tuple[List[Path], List[Path]]:
    """扫描源目录，返回 (慈善台账文件, 份额文件)。

    - 慈善台账：文件名含“慈善”，后缀 .xls/.xlsx
    - 份额文件：文件名含“持有人份额汇总信息查询”，后缀 .xlsx
    """
    charity_files: List[Path] = []
    holding_files: List[Path] = []
    if not source_dir.exists():
        return charity_files, holding_files

    for p in sorted(source_dir.rglob("*")):
        if not p.is_file():
            continue
        name = p.name
        lower_suffix = p.suffix.lower()
        if ("慈善" in name) and (lower_suffix in [".xls", ".xlsx"]):
            charity_files.append(p)
        if ("持有人份额汇总信息查询" in name) and (lower_suffix == ".xlsx"):
            holding_files.append(p)
    return charity_files, holding_files


def build_nonconflict_path(target_path: Path, overwrite: bool = False) -> Path:
    """目标存在且不覆盖时，追加数字后缀以避免冲突。"""
    if overwrite or not target_path.exists():
        return target_path
    stem = target_path.stem
    suffix = target_path.suffix
    parent = target_path.parent
    idx = 2
    while True:
        candidate = parent / f"{stem}_{idx}{suffix}"
        if not candidate.exists():
            return candidate
        idx += 1


def setup_logging(logs_dir: Path, level: int = logging.INFO) -> logging.Logger:
    logs_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    log_file = logs_dir / f"run-{ts}.log"

    logger = logging.getLogger("donation-script")
    logger.setLevel(level)
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.ERROR)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    logger.info("日志初始化完成: %s", log_file)
    return logger
