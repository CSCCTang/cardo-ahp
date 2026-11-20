#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""scheme_lib.py -- 方案清单与评分表的共享读法（08 步内部复用）。

《方案清单.xlsx》：序号 | 方案编号 | 方案名称 | 方案说明 | 是否参与评分
《AHP方案评分表.xlsx》：
    Sheet「方案信息」    序号 | 方案编号 | 方案名称 | 方案说明
    Sheet「方案评分表」  第 3 行表头，第 4 行起数据：
        A序号 B准则层 C指标 D指标ID E指标含义 F综合权重 G评分标准
        H.. 各方案得分（0-10） 其后 依据-各方案
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

SCHEME_LIST_SHEET = "方案清单"
SCORESHEET_SHEET = "方案评分表"
SHEET_INFO = "方案信息"
HEADER_ROW = 3
FIRST_DATA_ROW = 4
SCORE_COL = 8          # H 列：第一个方案得分列
ID_COL = 4             # D 列：指标ID
NAME_COL = 3           # C 列：指标名
HEAD_SCAN_ROWS = 12

# 方案名为空时评分表用的占位写法（gen_scoresheet 生成）
PLACEHOLDER_MARK = "（待填"


def _text(cell) -> str:
    v = cell.value
    if v is None:
        return ""
    return str(v).strip()


def read_scheme_list(path: Path) -> list[dict]:
    """读《方案清单.xlsx》：返回参与评分的方案 [{id, name, note}]。

    「是否参与评分」标「否」的行剔除。
    容错：表不存在 / 没有「方案清单」sheet / 方案名为空 → 返回空列表（调用方自行回落）。
    """
    if not path or not Path(path).is_file():
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    if SCHEME_LIST_SHEET not in wb.sheetnames:
        return []
    ws = wb[SCHEME_LIST_SHEET]
    out: list[dict] = []
    for r in range(2, ws.max_row + 1):
        sid = _text(ws.cell(row=r, column=2))
        name = _text(ws.cell(row=r, column=3))
        if not name:
            continue
        note = _text(ws.cell(row=r, column=4))
        flag = _text(ws.cell(row=r, column=5))
        if flag and flag not in ("是", "参与", "Y", "y", "yes", "True"):
            continue                      # 明确标「否」的方案不参与评分
        out.append({"id": sid or f"S{len(out) + 1}", "name": name, "note": note})
    return out


def read_project_name(path: Path) -> str:
    """读《方案清单.xlsx》B1 格的项目名称（用户自己填的）。

    项目名放在用户拿得到的表里，而不是只放在命令行参数里 ——
    用户打开模板填一格就行，不必去问 AI 或改参数文件。
    """
    if not path or not Path(path).is_file():
        return ""
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return ""
    if SCHEME_LIST_SHEET not in wb.sheetnames:
        return ""
    ws = wb[SCHEME_LIST_SHEET]
    return _text(ws.cell(row=1, column=2))


def find_header_row(ws) -> int:
    for r in range(1, HEAD_SCAN_ROWS + 1):
        c3 = _text(ws.cell(row=r, column=NAME_COL))
        c4 = _text(ws.cell(row=r, column=ID_COL))
        if c3 == "指标" or c4 in ("指标ID", "指标 ID"):
            return r
    return 0


def read_sheet_schemes(path: Path) -> list[dict]:
    """读已生成的评分表里登记的方案（用于与清单比对）。"""
    if not path or not Path(path).is_file():
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []
    out: list[dict] = []
    if SHEET_INFO in wb.sheetnames:
        ws = wb[SHEET_INFO]
        r = 2
        while r <= ws.max_row:
            sid = _text(ws.cell(row=r, column=2))
            name = _text(ws.cell(row=r, column=3))
            if name:
                out.append({"id": sid, "name": name, "note": _text(ws.cell(row=r, column=4))})
            r += 1
    if out:
        return out
    # 没有「方案信息」表时，从「方案评分表」表头推断
    if SCORESHEET_SHEET in wb.sheetnames:
        ws = wb[SCORESHEET_SHEET]
        hr = find_header_row(ws)
        if hr:
            col = SCORE_COL
            while _text(ws.cell(row=hr, column=col)) and not _text(ws.cell(row=hr, column=col)).startswith("依据"):
                out.append({"id": f"S{len(out) + 1}", "name": _text(ws.cell(row=hr, column=col)), "note": ""})
                col += 1
    return out


def is_placeholder_name(name: str) -> bool:
    """方案名是否还是占位（未填真实方案名）。"""
    return (not name) or PLACEHOLDER_MARK in name


def has_filled_scores(path: Path) -> bool:
    """评分表里是否已经填了得分（有分就不能随便重建表）。"""
    if not path or not Path(path).is_file():
        return False
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return False
    if SCORESHEET_SHEET not in wb.sheetnames:
        return False
    ws = wb[SCORESHEET_SHEET]
    hr = find_header_row(ws)
    if not hr:
        return False
    col = SCORE_COL
    while _text(ws.cell(row=hr, column=col)) and not _text(ws.cell(row=hr, column=col)).startswith("依据"):
        r = hr + 1
        while r <= ws.max_row:
            nm = _text(ws.cell(row=r, column=NAME_COL))
            if not nm or nm.startswith("综合得分") or nm.startswith("评级"):
                break
            v = ws.cell(row=r, column=col).value
            if v is not None and str(v).strip() != "":
                return True
            r += 1
        col += 1
    return False
