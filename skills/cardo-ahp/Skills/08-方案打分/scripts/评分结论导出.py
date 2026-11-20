#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评分结论导出.py -- 把评分 JSON 导出为交付用《AHP方案评分结论.xlsx》（成果文件只放 Office）。

三个工作表：
  1. 排序与评级：名次 / 方案 / 综合得分 / 评级 / 投资建议 / 权重覆盖率
  2. 指标得分明细：准则 / 指标 / 综合权重 / 各方案得分与加权贡献（贡献 = 得分 × 权重）
  3. 说明：口径、评级区间、数据来源

用法：
  python 评分结论导出.py --scores <评分表.json> --out <AHP方案评分结论.xlsx> [--project-name 名称]
退出码：0 成功；2 输入错。
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from _yaml_lite import load_file

RULES_DIR = HERE.parent / "规则"
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
ALT_FILL = PatternFill("solid", fgColor="EDF2F9")
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
# 演示口径：动作列从「投资建议」降为「方法输出」，只描述相对排序位置
DEMO_ACTION = {"优秀": "优秀档（[9,10]）", "良好": "良好档（[7,9)）",
               "一般": "一般档（[6,7)）", "不可行": "不可行档（[0,6)）"}


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def criterion_names() -> dict[str, str]:
    try:
        return {c["id"]: c["name"] for c in load_file(RULES_DIR / "指标体系.yaml")["criteria"]}
    except Exception:
        return {}


def header_row(ws, row: int, headers: list[str]) -> None:
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = WRAP
        c.border = BORDER


def fit_height(ws, row: int, width_sum: float, text, line_h: float = 15,
               pad: float = 4, min_h: float = 18) -> None:
    """按合并区总宽估算折行数并设行高，避免长文本被页边截断（不改内容）。"""
    units = sum(2 if ord(ch) > 127 else 1 for ch in str(text or ""))
    lines = max(1, math.ceil(units / max(10.0, width_sum - 2)))
    ws.row_dimensions[row].height = max(min_h, line_h * lines + pad)


def main() -> int:
    ap = argparse.ArgumentParser(description="评分 JSON → AHP方案评分结论.xlsx")
    ap.add_argument("--scores", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--project-name", default="某某片区改造项目")
    ap.add_argument("--demo", action="store_true",
                    help="演示样例：页脚标注「演示样例」，不改变任何数值")
    args = ap.parse_args()

    p = Path(args.scores)
    if not p.is_file():
        print(f"[AHP-E-INPUT] 评分结果不存在：{args.scores}", file=sys.stderr)
        return 2
    d = json.loads(p.read_text(encoding="utf-8-sig"))
    demo = bool(args.demo or d.get("demo"))
    schemes = d.get("schemes") or []
    if not schemes:
        print("[AHP-E-INPUT] 评分结果无 schemes", file=sys.stderr)
        return 2
    cnames = criterion_names()

    wb = Workbook()

    # ---------- 排序与评级 ----------
    ws = wb.active
    ws.title = "排序与评级"
    ncol = 7
    widths = [7, 16, 11, 9, 24, 11, 14]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    t = ws.cell(row=1, column=1, value=f"{args.project_name} · AHP 方案评分与排序")
    t.font = Font(size=14, bold=True, color="FFFFFF")
    t.fill = HEAD_FILL
    t.alignment = WRAP
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    note = (f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}　"
            f"评分口径：10 分制　权重来源：AHP 层次分析法（几何平均 + 特征值双算法）")
    ws.cell(row=2, column=1, value=note).alignment = WRAP
    fit_height(ws, 2, sum(widths), note)

    action_head = "方法输出" if demo else "投资建议"
    header_row(ws, 4, ["名次", "方案", "综合得分", "评级", action_head, "权重覆盖率", "备注"])
    name_rating = {s["name"]: s for s in schemes}
    row = 5
    for k, r in enumerate(d.get("ranking") or [], start=1):
        sc = name_rating.get(r.get("name"), {})
        cov = (sc.get("coverage") or {})
        covered = float(cov.get("weightCovered") or 0.0)
        # 0 分母不写 0%：一个指标都没得分时覆盖率无意义，按口径记「不适用」
        cover = f"{covered:.0%}" if covered > 0 else "不适用"
        vals = [r.get("order", k), r.get("name"), r.get("composite"),
                (r.get("rating") or {}).get("label"),
                (DEMO_ACTION.get((r.get("rating") or {}).get("label"), "方法输出")
                 if demo else (r.get("rating") or {}).get("action")),
                cover, "未覆盖指标重归一" if r.get("partial") else "全覆盖"]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.border = BORDER
            c.alignment = WRAP if col != 2 else LEFT
            if col == 3:
                c.number_format = "0.00"
                c.font = Font(bold=True)
            if k == 1:
                c.fill = ALT_FILL
        fit_height(ws, row, widths[4], vals[4], min_h=20)
        row += 1

    concl = (d.get("conclusion") or {}).get("summary")
    if concl:
        label = ws.cell(row=row + 1, column=1, value="结论")
        label.font = Font(bold=True)
        label.alignment = Alignment(vertical="top", horizontal="left")
        ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=ncol)
        c = ws.cell(row=row + 1, column=2, value=concl.replace("\n", " "))
        c.alignment = LEFT_TOP
        fit_height(ws, row + 1, sum(widths[1:]), concl, min_h=46)

    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ---------- 指标得分明细 ----------
    ws2 = wb.create_sheet("指标得分明细")
    base = ["准则", "指标", "指标ID", "综合权重"]
    heads = base[:]
    for s in schemes:
        heads += [f"{s['name']}得分", f"{s['name']}贡献"]
    for col, h in enumerate(heads, start=1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = WRAP
        c.border = BORDER

    first = schemes[0]
    by_scheme = {s["id"]: {x["id"]: x for x in s["scores"]} for s in schemes}
    row = 2
    for item in first["scores"]:
        iid = item["id"]
        vals = [cnames.get(item["criterionId"], item["criterionId"]), item["name"], iid,
                item["globalWeight"]]
        for s in schemes:
            x = by_scheme[s["id"]].get(iid, {})
            sc = x.get("score")
            vals += [sc, (sc * x["globalWeight"]) if sc is not None else None]
        for col, v in enumerate(vals, start=1):
            c = ws2.cell(row=row, column=col, value=v)
            c.border = BORDER
            c.alignment = LEFT if col == 2 else WRAP
            if col == 4:
                c.number_format = "0.00%"
            elif col > 4:
                c.number_format = "0.00"
        row += 1
    total_row = row
    ws2.cell(row=total_row, column=3, value="合计").font = Font(bold=True)
    ws2.cell(row=total_row, column=4, value=sum(i["globalWeight"] for i in first["scores"])).number_format = "0.00%"
    ranking = {r.get("name"): r for r in (d.get("ranking") or [])}
    for k in range(len(schemes)):
        col = 5 + k * 2 + 1
        # 成品不写公式（避免 Excel 打开二次计算）：合计直接取排序页同一 composite
        total = float((ranking.get(schemes[k]["name"]) or {}).get("composite") or 0.0)
        cell = ws2.cell(row=total_row, column=col, value=round(total, 2))
        cell.number_format = "0.00"
        cell.font = Font(bold=True)
        cell.fill = ALT_FILL
        cell.border = BORDER

    for col, w in enumerate([14, 20, 22, 11] + [10] * (len(schemes) * 2), start=1):
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.freeze_panes = "E2"

    # ---------- 说明 ----------
    ws3 = wb.create_sheet("说明")
    notes = [
        "AHP 方案评分结论 · 说明",
        "",
        "1. 综合得分 = Σ（各指标得分 × 综合权重）；贡献列 = 得分 × 权重，用于定位拉分/拖分项。",
        "2. 评级区间（左闭右开）：[9,10] 优秀 / [7,9) 良好 / [6,7) 一般 / [0,6) 不可行。",
        "3. 权重来自 AHP：专家成对比较 → 逐元素几何平均 → 几何平均行积法与特征值法双算 → CR<0.1 一致性检验。",
        "4. 评分由项目组/专家按 10 分制锚点逐项给分并注明依据；本表不再人工干预计算。",
        "5. 若某方案覆盖率为 100%，表示 16 项指标全部有分；未覆盖指标按已覆盖权重重归一，备注标为「部分覆盖」。",
        "6. 覆盖率为 0（无任何指标得分）时不做重归一，覆盖率栏记「不适用」。",
        "7. 「方法输出」列由评级档位直接推出，只说明方案落在哪一档。",
        ("8. 本表为方法链路验证输出：判断矩阵与指标得分来自本次合成问卷（5 份问卷、3 个候选方案），"
         "排序与评级只表示该组输入数据下的相对位置。" if demo else
         "8. 排序与评级为评分规则的方法输出，投资决策还需结合现状勘察、财务测算与实施条件。"),
    ]
    for r, text in enumerate(notes, start=1):
        c = ws3.cell(row=r, column=1, value=text)
        c.alignment = WRAP
        fit_height(ws3, r, 110, text, min_h=18)
    ws3.cell(row=1, column=1).font = Font(size=14, bold=True)
    ws3.column_dimensions["A"].width = 110

    # ---------- 封面（五项：项目、编制单位、版本、生成时间、页码） ----------
    cover = wb.create_sheet("封面", 0)
    cover.sheet_view.showGridLines = False
    cover.column_dimensions["A"].width = 16
    cover.column_dimensions["B"].width = 56
    cover["A2"] = "AHP 方案评分结论"
    cover["A2"].font = Font(size=26, bold=True)
    cover.merge_cells("A2:B2")
    cover.row_dimensions[2].height = 46
    cover_rows = [("项目名称", args.project_name), ("成果名称", "AHP 方案评分结论（成品版）"),
                  ("编制单位", "Cardo 城市更新 AHP 项目组"), ("版本", "v1.0"),
                  ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M"))]
    for i, (k, v) in enumerate(cover_rows, start=5):
        cover.cell(row=i, column=1, value=k).font = Font(size=12, bold=True)
        c = cover.cell(row=i, column=2, value=v)
        c.font = Font(size=12)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cover.row_dimensions[i].height = 22
    cover.page_setup.orientation = "portrait"
    cover.page_setup.paperSize = cover.PAPERSIZE_A4
    # 封面五项之一「页码」：页脚写 PAGE/NUMPAGES 域，打印即有页码
    cover.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    cover.oddFooter.center.size = 11

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    from openpyxl.worksheet.properties import PageSetupProperties
    for sh, titles in ((ws, "1:4"), (ws2, "1:1"), (ws3, None)):
        sh.page_setup.orientation = "landscape"
        sh.page_setup.paperSize = sh.PAPERSIZE_A4
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        if titles:
            sh.print_title_rows = titles
        sh.page_margins.bottom = 0.75
        sh.page_margins.footer = 0.3
        # 演示身份写在正文数据说明处，不在每页页脚撒角标（成品版对外交付）
        sh.oddFooter.center.text = "第 &P 页 / 共 &N 页"
        sh.oddFooter.center.size = 11
    try:
        wb.calculation.fullCalcOnLoad = False
    except Exception:
        pass
    wb.save(out)
    print(f"已导出评分结论：{out}（{len(schemes)} 个方案，{len(first['scores'])} 项指标）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
