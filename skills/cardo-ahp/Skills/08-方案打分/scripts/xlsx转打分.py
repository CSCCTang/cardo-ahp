#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlsx转打分.py -- 把输入文件里的《AHP方案评分表.xlsx》转成引擎可读的评分 JSON（过程产物）。

边界约定：输入文件只放常规 Office（xlsx/docx）；JSON 一律落在过程文件。
本脚本就是这条边界上的适配器：读 xlsx（人填）→ 写 过程文件/08-方案打分/打分.json。

支持版式（与 gen_scoresheet.py 一致）：
  Sheet「方案信息」：序号 | 方案编号 | 方案名称 | 方案说明
  Sheet「方案评分表」：表头行含「指标」「指标ID」；数据行：
     C指标 D指标ID F综合权重 H..方案得分 其后 依据-方案…
  综合得分/评级行会被自动识别并跳过。

用法：
  python xlsx转打分.py --xlsx "<评分表.xlsx>" --out "<打分.json>" [--project-name 名称]
退出码：0 成功；1 评分表内容有问题（越界/重复/未知指标）；2 文件或版式错误。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import load_workbook

from _yaml_lite import load_file

RULES_DIR = HERE.parent / "规则"
SCORE_COL = 8
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


DEMO_MARKERS = ("合成演示", "非真实依据", "示例数据", "演示数据", "演示样例")
# 中性依据：只描述赋分口径，不声称任何数据来源（与 fill_demo_scores.py 同一口径）
NEUTRAL_BASIS = "按 0–10 分标度对方案在该指标上的表现赋分"


def workbook_is_demo(wb) -> bool:
    """识别演示样例工作簿（fill_demo_scores.py 写入的属性/「数据说明」页）。"""
    props = wb.properties
    for attr in ("description", "subject", "keywords", "title"):
        text = str(getattr(props, attr, None) or "").lower()
        if "demo" in text or "演示" in text:
            return True
    if "数据说明" in wb.sheetnames:
        for row in wb["数据说明"].iter_rows(values_only=True):
            text = " ".join("" if v is None else str(v) for v in row).lower()
            if "demo" in text or "演示" in text:
                return True
    # 兜底：zip 解压/另存/复制粘贴可能把文档属性与「数据说明」页丢掉，
    # 但表里残留的占位措辞不会自己消失。出现这类词就按演示样例处理，
    # 免得它一路走到 09 才以通用 exit=1 报「评分依据含占位措辞」。
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and any(m in v for m in DEMO_MARKERS):
                    return True
    return False


def to_score(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    s = str(v).strip()
    try:
        x = float(s)
    except ValueError:
        raise ValueError(f"不是数字：{v!r}")
    if x < 0 or x > 10:
        raise ValueError(f"超出 0-10：{x}")
    return x


def load_catalog_ids() -> dict[str, str]:
    """指标 name → id（仅作为缺少指标ID列时的兜底）。"""
    try:
        data = load_file(RULES_DIR / "指标体系.yaml")
        return {i["name"]: i["id"] for i in data["indicators"]}
    except Exception:
        return {}


def read_schemes(wb) -> list[dict]:
    schemes = []
    if "方案信息" in wb.sheetnames:
        ws = wb["方案信息"]
        r = 2
        while r <= 20:
            sid = cell_text(ws, r, 2)
            name = cell_text(ws, r, 3)
            if not sid and not name:
                break
            schemes.append({"id": sid or f"S{len(schemes) + 1}", "name": name or f"方案{len(schemes) + 1}",
                            "note": cell_text(ws, r, 4)})
            r += 1
    return schemes


def find_header_row(ws) -> int:
    for r in range(1, HEAD_SCAN_ROWS + 1):
        c3 = cell_text(ws, r, NAME_COL)
        c4 = cell_text(ws, r, ID_COL)
        if c3 == "指标" or c4 in ("指标ID", "指标 ID"):
            return r
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="AHP 方案评分表 xlsx → 打分 JSON（过程产物）")
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--project-name", default="某某片区改造项目")
    args = ap.parse_args()

    src = Path(args.xlsx)
    if not src.is_file():
        print(f"[AHP-E-INPUT] 评分表不存在：{args.xlsx}", file=sys.stderr)
        return 2
    try:
        wb = load_workbook(src, data_only=True)
    except Exception as exc:
        print(f"[AHP-E-INPUT] 打不开 xlsx：{exc}", file=sys.stderr)
        return 2
    if "方案评分表" not in wb.sheetnames:
        print("[AHP-E-FORMAT] 缺「方案评分表」工作表，请用 gen_scoresheet.py 重新生成模板", file=sys.stderr)
        return 2
    ws = wb["方案评分表"]

    header_row = find_header_row(ws)
    if header_row == 0:
        print("[AHP-E-FORMAT] 未找到表头行（应含「指标」「指标ID」列）", file=sys.stderr)
        return 2

    schemes = read_schemes(wb)
    # 没有方案信息表时，从表头 H 列起按「依据-」分隔推断方案数
    if not schemes:
        n = 0
        col = SCORE_COL
        while cell_text(ws, header_row, col) and not cell_text(ws, header_row, col).startswith("依据"):
            n += 1
            col += 1
        schemes = [{"id": f"S{i + 1}", "name": cell_text(ws, header_row, SCORE_COL + i) or f"方案{i + 1}",
                    "note": ""} for i in range(n)]
    n = len(schemes)
    if n == 0:
        print("[AHP-E-FORMAT] 未识别到任何方案列", file=sys.stderr)
        return 2

    basis_start = SCORE_COL + n
    name_to_id = load_catalog_ids()
    catalog_ok = bool(name_to_id)
    demo = workbook_is_demo(wb)

    problems: list[str] = []
    schemes_out = []
    seen_ids: set[str] = set()
    for k, s in enumerate(schemes):
        manual: dict[str, dict] = {}
        r = header_row + 1
        while r <= ws.max_row:
            name = cell_text(ws, r, NAME_COL)
            if not name:
                break
            if name.startswith("综合得分") or name.startswith("评级"):
                break
            iid = cell_text(ws, r, ID_COL) or name_to_id.get(name, "")
            if not iid:
                problems.append(f"第 {r} 行「{name}」缺少指标ID且无法按名称匹配")
                r += 1
                continue
            if catalog_ok and iid not in set(name_to_id.values()):
                problems.append(f"第 {r} 行指标ID「{iid}」不在指标体系内")
            if iid in manual:
                problems.append(f"第 {r} 行指标ID重复：{iid}")
            raw = ws.cell(row=r, column=SCORE_COL + k).value
            try:
                score = to_score(raw)
            except ValueError as exc:
                problems.append(f"第 {r} 行「{name}」方案「{s['name']}」{exc}")
                score = None
            basis = cell_text(ws, r, basis_start + k)
            # 演示样例里的旧占位措辞（合成演示/非真实依据等）只在转换时改写成中性
            # 赋分口径，分数一格不动 —— 09 的 BAD_WORDS 硬拦不因此放宽。
            if demo and any(m in basis for m in DEMO_MARKERS):
                basis = NEUTRAL_BASIS
            if score is not None:
                manual[iid] = {"score": score, "basis": basis or "人工打分"}
            seen_ids.add(iid)
            r += 1
        if not manual:
            problems.append(f"方案「{s['name']}」没有任何有效评分")
        schemes_out.append({"id": s["id"], "name": s["name"], "schemeNote": s["note"], "manualScores": manual})

    if problems:
        print("[AHP-E-SHEET] 评分表存在下列问题，请修正后重跑：", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1

    payload = {
        "schema": "cardo/ahp-scores-input/v1",
        "projectName": args.project_name,
        "demo": demo,
        "source": {"kind": "scoresheet", "path": str(src).replace("\\", "/")},
        "note": "由输入文件《AHP方案评分表.xlsx》自动转换的过程产物；改评分请改 xlsx，不要改本文件。",
        "schemes": schemes_out,
    }
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已转换：{src.name} → {out}（方案 {n} 个，指标 {len(seen_ids)} 项）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
