#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""sync_scheme_names.py -- 把《方案清单.xlsx》的方案名同步进已生成的评分表。

场景：评分表在首跑就出了（空白待填），用户之后才在《方案清单.xlsx》里填好方案名。
此时要更新评分表里的方案名，但**不能重建整张表**（用户可能已经填了得分）。

本脚本做两件事：
  1. 改「方案信息」页的方案名称与说明；
  2. 改「方案评分表」页的方案列表头与「依据-xxx」列表头。
得分格、依据格一律不动。

方案数量变了（增删行）时不做列结构调整 —— 那需要重建表，
交由 pipeline 判断（没填分就重建，填了分就报错让人处理）。

用法：
  python sync_scheme_names.py --xlsx "<评分表.xlsx>" --scheme-list "<方案清单.xlsx>"
退出码：0 成功；2 参数/文件错；3 方案数量不一致（需重建表）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import load_workbook                      # noqa: E402
from openpyxl.styles import Font                       # noqa: E402

from scheme_lib import (SHEET_INFO, SCORESHEET_SHEET, HEADER_ROW, SCORE_COL,  # noqa: E402
                        find_header_row, read_scheme_list, read_sheet_schemes)


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="把方案清单里的方案名同步进评分表")
    ap.add_argument("--xlsx", required=True, help="评分表 xlsx 路径")
    ap.add_argument("--scheme-list", required=True, help="方案清单 xlsx 路径")
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.is_file():
        print(f"[AHP-E-INPUT] 评分表不存在：{xlsx}", file=sys.stderr)
        return 2

    listed = read_scheme_list(Path(args.scheme_list))
    if not listed:
        print("[AHP-E-INPUT] 方案清单为空或未填方案名，跳过同步", file=sys.stderr)
        return 2

    wb = load_workbook(xlsx)
    existing = read_sheet_schemes(xlsx)
    if existing and len(existing) != len(listed):
        print(f"[AHP-E-SCHEME] 方案数量不一致：评分表 {len(existing)} 个、清单 {len(listed)} 个", file=sys.stderr)
        return 3

    changed = []

    # ---- 1. 方案信息页 ----
    if SHEET_INFO in wb.sheetnames:
        ws = wb[SHEET_INFO]
        for idx, s in enumerate(listed, start=2):
            old_name = (ws.cell(row=idx, column=3).value or "").strip()
            if old_name != s["name"]:
                changed.append(f"{old_name or '(空)'} → {s['name']}")
            c = ws.cell(row=idx, column=2, value=s["id"])
            c.font = Font(size=11)
            ws.cell(row=idx, column=3, value=s["name"])
            ws.cell(row=idx, column=4, value=s["note"])

    # ---- 2. 方案评分表页表头 ----
    if SCORESHEET_SHEET in wb.sheetnames:
        ws = wb[SCORESHEET_SHEET]
        hr = find_header_row(ws)
        if hr:
            n = len(listed)
            for k, s in enumerate(listed):
                ws.cell(row=hr, column=SCORE_COL + k, value=s["name"])
                ws.cell(row=hr, column=SCORE_COL + n + k, value=f"依据-{s['name']}")

    wb.save(xlsx)
    if changed:
        print("已同步方案名到评分表：" + "；".join(changed))
    else:
        print("评分表方案名与清单一致，无需改动")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
