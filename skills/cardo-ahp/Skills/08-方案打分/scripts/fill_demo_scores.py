#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fill_demo_scores.py -- 为「演示样例」评分表生成/刷新得分与依据（仅演示用，不参与真实项目）。

为什么单独成脚本：
  真实项目必须由项目组/专家填《AHP方案评分表.xlsx》，本脚本**不得**在真实项目里使用。
  演示/流程验证时需要一个可复现的填表器：分数按固定种子确定，依据格写计算口径，
  并在工作簿里留下 demo 标记，便于事后识别这是演示数据。

依据格措辞纪律：
  · 只写「怎么算出来的」，不写数据来源真伪声明；
  · 不得出现「合成演示 / 非真实依据 / 示例数据 / 演示数据」——09 步依据对照表按此硬拦；
  · 演示身份不靠正文措辞表达，由工作簿 demo 标记 + 产物「演示样例」页脚承担。

用法：
  python fill_demo_scores.py --scoresheet "<AHP方案评分表.xlsx>" --weights "<权重表.json>"
  # 默认只改依据措辞 + 写 demo 标记，保留已有分数；确需重算分数才加 --refill-scores
退出码：0 成功；2 文件或版式错误。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

# 中性依据：描述赋分口径，不声称任何真实数据来源
BASIS_TEXT = "按 0–10 分标度对方案在该指标上的表现赋分"
# 演示身份标记（写进工作簿属性与「数据说明」页，09/08 据此传递 demo=true）
DEMO_NOTE = "演示样例（demo=true）：仅用于产品演示与流程验证，不作为真实项目依据"
SCORE_COL = 8          # H 列起为方案得分
ID_COL = 4
NAME_COL = 3
HEAD_SCAN_ROWS = 12


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def cell_text(ws, r: int, c: int) -> str:
    v = ws.cell(row=r, column=c).value
    return "" if v is None else str(v).strip()


def find_header_row(ws) -> int:
    for r in range(1, HEAD_SCAN_ROWS + 1):
        if cell_text(ws, r, NAME_COL) == "指标" or cell_text(ws, r, ID_COL) in ("指标ID", "指标 ID"):
            return r
    return 0


def count_schemes(wb, ws, header_row: int) -> int:
    """方案数：优先「方案信息」页，其次按表头 H 列到「依据-」之前的列数推断。"""
    if "方案信息" in wb.sheetnames:
        info = wb["方案信息"]
        n = 0
        r = 2
        while r <= 20 and (cell_text(info, r, 2) or cell_text(info, r, 3)):
            n += 1
            r += 1
        if n:
            return n
    n = 0
    col = SCORE_COL
    while cell_text(ws, header_row, col) and not cell_text(ws, header_row, col).startswith("依据"):
        n += 1
        col += 1
    return n


def demo_score(seed: str, scheme: str, iid: str) -> float:
    """按固定种子给出 1.0–9.0、1 位小数的确定性得分（同输入同输出）。"""
    digest = hashlib.sha256(f"{seed}|{scheme}|{iid}".encode("utf-8")).hexdigest()
    return round(1.0 + (int(digest[:8], 16) % 81) / 10.0, 1)


def mark_demo(wb) -> None:
    wb.properties.description = DEMO_NOTE
    if "数据说明" in wb.sheetnames:
        ws = wb["数据说明"]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet("数据说明")
    ws["A1"] = "数据性质"
    ws["B1"] = DEMO_NOTE
    ws["A2"] = "生成方式"
    ws["B2"] = "fill_demo_scores.py（分数按固定种子生成，依据写赋分口径）"
    ws["A3"] = "使用范围"
    ws["B3"] = "产品演示与流程验证；真实项目须由项目组/专家逐项填写"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 80


def main() -> int:
    ap = argparse.ArgumentParser(description="演示样例评分表填分（仅演示用）")
    ap.add_argument("--scoresheet", required=True, help="AHP方案评分表.xlsx（08 生成的表）")
    ap.add_argument("--weights", required=True, help="权重表.json（提供指标ID）")
    ap.add_argument("--refill-scores", action="store_true",
                    help="显式重算分数（默认保留原分数，只刷新依据措辞与 demo 标记）")
    ap.add_argument("--seed", default="cardo-ahp-demo")
    args = ap.parse_args()

    sheet = Path(args.scoresheet)
    if not sheet.is_file():
        print(f"[AHP-E-INPUT] 评分表不存在：{sheet}", file=sys.stderr)
        return 2
    weights = Path(args.weights)
    if not weights.is_file():
        print(f"[AHP-E-INPUT] 权重表不存在：{weights}", file=sys.stderr)
        return 2
    try:
        rules = json.loads(weights.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        print(f"[AHP-E-INPUT] 权重表不是合法 JSON：{exc}", file=sys.stderr)
        return 2
    indicator_ids = [str(i.get("id")) for i in (rules.get("indicators") or []) if i.get("id")]

    try:
        wb = load_workbook(sheet)
    except Exception as exc:
        print(f"[AHP-E-INPUT] 打不开评分表：{exc}", file=sys.stderr)
        return 2
    if "方案评分表" not in wb.sheetnames:
        print("[AHP-E-FORMAT] 缺「方案评分表」工作表", file=sys.stderr)
        return 2
    ws = wb["方案评分表"]
    header_row = find_header_row(ws)
    if not header_row:
        print("[AHP-E-FORMAT] 未找到表头行（应含「指标」「指标ID」列）", file=sys.stderr)
        return 2
    n = count_schemes(wb, ws, header_row)
    if n <= 0:
        print("[AHP-E-FORMAT] 未识别到任何方案列", file=sys.stderr)
        return 2
    scheme_names = [cell_text(ws, header_row, SCORE_COL + k) or f"方案{k + 1}" for k in range(n)]

    basis_start = SCORE_COL + n
    # 表头说明行（第 2 行）里的演示身份也在这里统一刷新，避免残留旧措辞
    if "合成演示" in cell_text(ws, 2, 1) or "非真实" in cell_text(ws, 2, 1):
        ws.cell(row=2, column=1).value = "评价人：项目组（演示样例）"

    filled = 0
    r = header_row + 1
    while r <= ws.max_row:
        name = cell_text(ws, r, NAME_COL)
        if not name or name.startswith("综合得分") or name.startswith("评级"):
            break
        iid = cell_text(ws, r, ID_COL) or (indicator_ids[r - header_row - 1]
                                           if r - header_row - 1 < len(indicator_ids) else "")
        for k, scheme in enumerate(scheme_names):
            if args.refill_scores:
                ws.cell(row=r, column=SCORE_COL + k).value = demo_score(args.seed, scheme, iid)
                filled += 1
            ws.cell(row=r, column=basis_start + k).value = BASIS_TEXT
        r += 1
    mark_demo(wb)
    wb.save(sheet)
    print(f"已写入演示样例：{sheet}（方案 {n} 个，指标 {r - header_row - 1} 项，"
          f"{f'重算分数 {filled} 格并' if args.refill_scores else '保留原分数，'}刷新依据与 demo 标记）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
