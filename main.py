from __future__ import annotations

import sys
from pathlib import Path
import logging
import warnings

from scr.io_utils import get_project_paths, ensure_directories, setup_logging
from scr.orchestrator import run, DuplicateProductError


def main() -> int:
    # 隐藏 openpyxl 控制台警告
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    root = Path(__file__).resolve().parent
    paths = get_project_paths(root)
    ensure_directories(paths)
    logger = setup_logging(paths["logs_dir"], level=logging.INFO)

    try:
        stats = run(root=root, overwrite=False, logger=logger)
        # 控制台仅输出简要汇总
        print("\n" + "="*50)
        print("处理完成")
        print("="*50)
        print(f"数据源文件:")
        print(f"  - 捐赠台账: {stats.get('charity_files', 0)} 个")
        print(f"  - 份额文件: {stats.get('holding_files', 0)} 个")
        print(f"  - 已捐项目: {stats.get('charity_rows', 0)} 个")
        print(f"  - 份额记录: {stats.get('holding_records', 0)} 条")
        print(f"\n生成结果:")
        print(f"  - 代捐说明: 成功 {stats.get('docs_success', 0)} 个 / 失败 {stats.get('docs_failed', 0)} 个")
        print(f"  - 明细表:   生成 {stats.get('detail_files', 0)} 个")
        print(f"  - 表插入:   成功 {stats.get('detail_attach_success', 0)} 个 / 失败 {stats.get('detail_attach_failed', 0)} 个")
        print("="*50 + "\n")
        return 0
    except FileNotFoundError as e:
        print(f"必要文件缺失: {e}")
        return 2
    except DuplicateProductError as e:
        print(f"数据错误: {e}")
        return 3
    except Exception as e:
        # 日志记录完整堆栈
        logger.exception("运行过程中发生未预期错误: %s", e)
        print("运行失败，详情见日志。")
        return 1


if __name__ == "__main__":
    sys.exit(main())


