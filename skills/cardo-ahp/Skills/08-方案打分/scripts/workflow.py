#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""workflow.py -- 逐步引导：判断「用户现在到第几步、每步做完没有」。

**为什么需要**：给了空白表还不够 —— 用户要知道「拿到这张表第一步干什么、干完长什么样、
怎么确认做对了」。所以每一步都要能自动判定状态，而不是让人自己回忆。

判定靠**读文件实际内容**，不靠文件名猜：
  · 方案清单：B1 有项目名？方案名称列有值？
  · 问卷：领料区有空白卷？回收区收齐 ≥5 份？
  · 评分表：有得分格？
  · 成果：成果文件出齐了？

状态三态：done（做完）/ doing（做了但没做完）/ todo（还没开始）
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

PENDING_REL = "输入文件/02-专家问卷（待发放）"
RECV_REL = "输入文件/02-专家问卷（已回收）"
SCORE_DIR_REL = "输入文件/03-方案打分表"
SCHEME_LIST_REL = SCORE_DIR_REL + "/方案清单.xlsx"
SCORESHEET_REL = SCORE_DIR_REL + "/AHP方案评分表.xlsx"
OUT_REL = "成果文件"

MIN_QUESTIONNAIRES = 5


def _text(cell) -> str:
    v = cell.value
    return "" if v is None else str(v).strip()


def _survey_count(ws_dir: Path) -> int:
    if not ws_dir.is_dir():
        return 0
    return len([p for p in ws_dir.glob("AHP专家问卷-专家*.xlsx")
                if not p.name.startswith("~$")])


def scheme_state(ws: Path) -> dict:
    """方案清单填到什么程度。"""
    path = ws / SCHEME_LIST_REL
    if not path.is_file():
        return {"status": "todo", "project": "", "schemes": [], "note": "文件还没生成"}
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return {"status": "todo", "project": "", "schemes": [], "note": "文件打不开"}
    if "方案清单" not in wb.sheetnames:
        return {"status": "todo", "project": "", "schemes": [], "note": "缺「方案清单」页"}
    sh = wb["方案清单"]
    project = _text(sh.cell(row=1, column=2))
    schemes = []
    for r in range(4, sh.max_row + 1):
        name = _text(sh.cell(row=r, column=3))
        if name:
            schemes.append(name)
    wb.close()
    if project and schemes:
        return {"status": "done", "project": project, "schemes": schemes,
                "note": "项目名与方案名都填好了"}
    if schemes:
        return {"status": "doing", "project": project, "schemes": schemes,
                "note": "方案名填好了，还差 B1 行的项目名称"}
    if project:
        return {"status": "doing", "project": project, "schemes": [],
                "note": "项目名填好了，还差方案名称"}
    return {"status": "doing", "project": "", "schemes": [],
            "note": "表已铺好，还没填内容"}


def survey_state(ws: Path) -> dict:
    pending = _survey_count(ws / PENDING_REL)
    recv = _survey_count(ws / RECV_REL)
    return {"pending": pending, "recv": recv,
            "status": "done" if recv >= MIN_QUESTIONNAIRES else ("doing" if recv else "todo")}


def score_state(ws: Path) -> dict:
    path = ws / SCORESHEET_REL
    if not path.is_file():
        # 评分表在第 08 步生成（要等权重算完），不在第 04 步 —— 说明原因，别只说"还没生成"
        return {"status": "todo", "filled": 0,
                "note": "还没生成（跑完 --to 07 算完权重后，08 步会按方案数生成）"}
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return {"status": "todo", "filled": 0, "note": "文件打不开"}
    if "方案评分表" not in wb.sheetnames:
        return {"status": "todo", "filled": 0, "note": "缺「方案评分表」页"}
    sh = wb["方案评分表"]
    hdr = None
    for r in range(1, 12):
        if _text(sh.cell(row=r, column=3)) == "指标":
            hdr = r
            break
    if not hdr:
        return {"status": "todo", "filled": 0, "note": "找不到表头"}
    filled = 0
    col = 8
    while _text(sh.cell(row=hdr, column=col)) and not _text(sh.cell(row=hdr, column=col)).startswith("依据"):
        r = hdr + 1
        while r <= sh.max_row:
            nm = _text(sh.cell(row=r, column=3))
            if not nm or nm.startswith("综合得分") or nm.startswith("评级"):
                break
            v = sh.cell(row=r, column=col).value
            if v is not None and str(v).strip() != "":
                filled += 1
            r += 1
        col += 1
    wb.close()
    total = 16 * max(0, col - 8)
    if filled == 0:
        return {"status": "doing", "filled": 0, "note": "表已生成，还没打分"}
    if filled < total:
        return {"status": "doing", "filled": filled,
                "note": f"已填 {filled}/{total} 格，还有没填的"}
    return {"status": "done", "filled": filled, "note": f"{filled} 格全部填好"}


def result_state(ws: Path) -> dict:
    out = ws / OUT_REL
    need = ["AHP权重表.xlsx", "AHP方案评分结论.xlsx", "AHP权重与评分结论.docx"]
    have = [n for n in need if (out / n).is_file()]
    return {"status": "done" if len(have) == len(need) else ("doing" if have else "todo"),
            "have": have, "need": need}


def build(ws: Path) -> list[dict]:
    """组装完整工作流（每步带状态与操作指引）。"""
    sch = scheme_state(ws)
    sur = survey_state(ws)
    sco = score_state(ws)
    res = result_state(ws)

    steps = [
        {
            "no": "①", "title": "定方案", "who": "你", "state": sch["status"],
            "what": "打开《方案清单.xlsx》填两样：",
            "detail": [
                "B1 格（黄色）填**项目名称** —— 会印在《AHP权重表》《AHP方案评分结论》"
                "《AHP权重与评分结论》三份成果文件上",
                "「方案清单」页填**方案名称**与**方案说明**（可增删行；不知道定哪几个方案，"
                "先翻「候选方案参考」页看思路）",
            ],
            "where": norm_path(ws / SCHEME_LIST_REL),
            "done_when": "B1 有项目名 + 方案名称列有 2~4 行有字",
            "how_check": "重跑一次 `python CardoAHP\\开工.py`，本文件的状态会变成「已完成」",
            "progress": sch["note"],
        },
        {
            "no": "②", "title": "发问卷", "who": "你 → 专家",
            "state": ("done" if sur["recv"] >= MIN_QUESTIONNAIRES
                      else ("doing" if sur["recv"] else "todo")),
            "what": "把领料区的空白卷发给专家（一位专家一份，不要合并、不要改名）：",
            "detail": [
                f"共 {sur['pending']} 份，在「待发放」目录",
                "问卷封面有完整标度说明，直接转发即可，专家不用再问",
                "「指标层」每个矩阵右侧有「指标说明」列 —— 提醒专家填前先看一眼",
            ],
            "where": norm_path(ws / PENDING_REL),
            "done_when": "每份问卷都已发出，专家手上有一份",
            "how_check": "发完不需要检查，下一步等专家交回即可",
            "progress": (f"已回收 {sur['recv']} 份（专家开始交了）" if sur["recv"]
                         else f"待发 {sur['pending']} 份"),
        },
        {
            "no": "③", "title": "回收问卷", "who": "专家 → 你", "state": sur["status"],
            "what": "专家填完的问卷**原样**放回回收区（文件名不要改）：",
            "detail": [
                f"已回收 {sur['recv']} 份，至少要 {MIN_QUESTIONNAIRES} 份才能算权重",
                "收齐后跑 `python CardoAHP\\开工.py --to 07` 算权重与一致性",
                "**如果一致性 CR 超 0.1**，看 `过程文件/03-回收校验/回收校验报告.md`，"
                "里面有「怎么改」和标度对照，转给专家本人调整",
            ],
            "where": norm_path(ws / RECV_REL),
            "done_when": f"回收区里 ≥ {MIN_QUESTIONNAIRES} 份问卷",
            "how_check": "跑 `--to 07`，通过就说明份数够且一致性达标",
            "progress": f"已回收 {sur['recv']}/{MIN_QUESTIONNAIRES} 份",
        },
        {
            "no": "④", "title": "打分", "who": "你", "state": sco["status"],
            "what": "在评分表里给每个方案、每个指标打分：",
            "detail": [
                "评分表在方案定好、权重算完后自动生成（列数 = 你的方案数）",
                "**「参考分档」列**写好了每个指标打 1/3/5/7/9 分分别代表什么 —— 照着它给分",
                "**黄色格**填 0–10 分（1 位小数），**依据列**写这一分怎么来的",
                "「综合权重」列现在是空的，等权重算出来会自动回填，不影响你先填",
            ],
            "where": norm_path(ws / SCORESHEET_REL),
            "done_when": "所有方案、所有指标的黄色格都填了，依据列也写了出处",
            "how_check": "跑 `python CardoAHP\\开工.py --from 08`，缺项会被点名",
            "progress": sco["note"],
        },
        {
            "no": "⑤", "title": "出成果", "who": "系统", "state": res["status"],
            "what": "跑最后一步，系统算完并出成果文件：",
            "detail": [
                "命令：`python CardoAHP\\开工.py --from 08`",
                "产出：`成果文件/AHP权重表.xlsx`、`AHP方案评分结论.xlsx`、"
                "`AHP权重与评分结论.docx`",
                "第 09 步会独立复算一遍（不信任前面的结果，自己再算一次对账）",
            ],
            "where": norm_path(ws / OUT_REL),
            "done_when": "三份成果文件都在 成果文件/ 里",
            "how_check": "看 成果文件/ 目录；或看终端最后一行「全链通过」",
            "progress": ("已出 " + "、".join(res["have"])) if res["have"] else "还没跑",
        },
    ]
    return steps


def norm_path(p: Path) -> str:
    return str(p).replace("\\", "/")


STATUS_TEXT = {
    "done": "已完成",
    "doing": "进行中",
    "todo": "未开始",
}
STATUS_MARK = {
    "done": "✅",
    "doing": "▶",
    "todo": "○",
}
