#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""专家名单导出.py -- 按回收的问卷生成《专家名单.xlsx》（输入文件侧的人读登记表）。

边界约定：输入文件只放常规 Office；原先的 问卷清单.json 由此 xlsx 取代，
机读清单仍由 gen_questionnaire.py 写在过程文件里（过程文件/02-问卷生成/问卷空白/问卷清单.json）。

用法：
  python 专家名单导出.py --dir "输入文件/02-专家问卷（已回收）" [--force]
  已存在且未加 --force 时保留人工编辑内容，仅打印现状。
退出码：0 成功（含已存在跳过）；2 目录错。
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

ROSTER_NAME = "专家名单.xlsx"
DEFAULT_ROLES = ["项目负责人", "投资专家", "投资专家", "规划专家", "财务专家", "结构专家", "造价专家"]


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="生成专家名单 xlsx（输入侧人读登记表）")
    ap.add_argument("--dir", required=True, help="已回收问卷目录")
    ap.add_argument("--force", action="store_true", help="覆盖已存在的专家名单")
    args = ap.parse_args()

    d = Path(args.dir)
    if not d.is_dir():
        print(f"[AHP-E-INPUT] 目录不存在：{args.dir}", file=sys.stderr)
        return 2

    files = sorted(p for p in d.glob("AHP专家问卷-专家*.xlsx") if p.is_file())
    roster = d / ROSTER_NAME
    if roster.is_file() and not args.force:
        wb = load_workbook(roster, data_only=True)
        ws = wb.active
        print(f"专家名单已存在，保留人工编辑：{roster}（{ws.max_row - 1} 行；覆盖请加 --force）")
        return 0
    if not files:
        print(f"[AHP-E-INPUT] 目录内没有 AHP专家问卷-专家*.xlsx：{d}", file=sys.stderr)
        return 2

    wb = Workbook()
    ws = wb.active
    ws.title = "专家名单"
    headers = ["专家编号", "专家角色", "问卷文件", "准则层标度", "指标层标度", "回收状态", "备注"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = WRAP
        c.border = BORDER

    for i, f in enumerate(files, start=1):
        role = DEFAULT_ROLES[i - 1] if i <= len(DEFAULT_ROLES) else f"专家{i}"
        vals = [str(i), role, f.name, "1-9", "0.1-1", "已回收", ""]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=i + 1, column=col, value=v)
            c.alignment = LEFT if col in (3, 7) else WRAP
            c.border = BORDER

    note_row = len(files) + 3
    ws.cell(row=note_row, column=1,
            value=("说明：本表是专家构成的登记表（人读）；专家等权，角色仅用于报告，不参与计算。"
                   "回收问卷放入本目录即可被第 03 步校验、第 04 步聚合；"
                   "机读清单在 过程文件/02-问卷生成/。")).alignment = LEFT
    ws.cell(row=note_row + 1, column=1, value=f"登记日期：{date.today().isoformat()}").alignment = LEFT

    for col, w in enumerate([10, 14, 30, 12, 12, 10, 30], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    wb.save(roster)
    print(f"已生成专家名单：{roster}（{len(files)} 位专家）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
