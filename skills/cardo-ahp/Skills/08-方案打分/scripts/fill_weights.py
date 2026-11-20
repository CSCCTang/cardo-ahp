#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fill_weights.py -- 把算出的综合权重回填进评分表的 F 列。

背景：评分表现在是**首跑就铺出来的空白表**（用户手上什么都没有时，
系统先把表给全）。但权重要等专家问卷回收后第 05 步才算出来，
所以首跑出的表里综合权重列是空的。本脚本在第 08 步把真实权重写回 F 列。

只写 F 列（综合权重），得分格、依据格、其它列一律不动。

用法：
  python fill_weights.py --xlsx "<评分表.xlsx>" --weights "<权重表.json>"
退出码：0 成功（含「表已由系统生成、无需回填」的情况）；2 参数/文件错。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import load_workbook                      # noqa: E402
from openpyxl.styles import PatternFill                 # noqa: E402

from scheme_lib import (SCORESHEET_SHEET, ID_COL, find_header_row)   # noqa: E402

WEIGHT_COL = 6      # F 列
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="回填评分表的综合权重列")
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--weights", required=True)
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    wpath = Path(args.weights)
    if not xlsx.is_file():
        print(f"[AHP-E-INPUT] 评分表不存在：{xlsx}", file=sys.stderr)
        return 2
    if not wpath.is_file():
        print(f"[AHP-E-INPUT] 权重文件不存在：{wpath}", file=sys.stderr)
        return 2

    try:
        weights = json.loads(wpath.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        print(f"[AHP-E-INPUT] 权重文件不可读：{exc}", file=sys.stderr)
        return 2
    gw = {}
    for item in weights.get("indicators", []) or []:
        iid = item.get("id")
        if iid:
            gw[iid] = float(item.get("globalWeight") or 0.0)
    if not gw:
        print("[AHP-E-INPUT] 权重 JSON 里没有 indicators[]，无法回填", file=sys.stderr)
        return 2

    wb = load_workbook(xlsx)
    if SCORESHEET_SHEET not in wb.sheetnames:
        print(f"[AHP-E-FORMAT] 评分表缺「{SCORESHEET_SHEET}」页", file=sys.stderr)
        return 2
    ws = wb[SCORESHEET_SHEET]
    hr = find_header_row(ws)
    if not hr:
        print("[AHP-E-FORMAT] 未找到表头行，无法回填", file=sys.stderr)
        return 2

    filled, r = 0, hr + 1
    while r <= ws.max_row:
        iid = str(ws.cell(row=r, column=ID_COL).value or "").strip()
        if not iid:
            break
        if iid in gw:
            cell = ws.cell(row=r, column=WEIGHT_COL, value=gw[iid])
            cell.number_format = "0.00%"
            # 清掉首跑时打的「待回填」灰底（用默认无填充，不能给 cell.fill 赋 StyleProxy）
            cell.fill = PatternFill()
            filled += 1
        r += 1

    # 综合得分行的展示公式：权重补齐后才能算对
    last_data = r - 1
    n = 0
    col = 8
    while ws.cell(row=hr, column=col).value and not str(ws.cell(row=hr, column=col).value).startswith("依据"):
        n += 1
        col += 1
    from openpyxl.utils import get_column_letter
    for k in range(n):
        cl = get_column_letter(8 + k)
        ws.cell(row=r, column=8 + k,
                value=f"=SUMPRODUCT($F${hr + 1}:$F${last_data},{cl}{hr + 1}:{cl}{last_data})").number_format = "0.00"

    wb.save(xlsx)
    print(f"已回填综合权重 {filled} 项到评分表")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
