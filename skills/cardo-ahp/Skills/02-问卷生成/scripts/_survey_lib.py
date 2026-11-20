#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_survey_lib.py -- cardo-ahp 问卷版式与回收读取（openpyxl，零第三方额外依赖）。

版式（每人一份 xlsx）：
  · 「封面」：项目名占位、标度说明、填写人、日期、填写规则
  · 「准则层」：4×4 两两比较；只填灰色上三角（6 格）；对角线=1；下三角=1/上三角（公式）
  · 「指标层」：4 个 4×4（现状/投资/规划/风险各一块），同上
  · 「自动校验」：5 个矩阵块各自给 列和/行几何平均/权重/λmax/CI/RI/CR/判定 + 缺项/越界/互反异常计数（公式）
本模块只负责读写与结构校验；CR 计算在 CLI 用 `_ahp_lib`（同一口径）。
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.formula import ArrayFormula  # noqa: E402

import _ahp_lib as lib  # noqa: E402

GRAY = PatternFill("solid", fgColor="D9D9D9")
HEAD = PatternFill("solid", fgColor="DDEBF7")
CALC = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
COLS = ["B", "C", "D", "E"]
RI_CHOOSE = "CHOOSE(MIN({n},10),0,0,0.58,0.9,1.12,1.24,1.32,1.41,1.45,1.49)"
SCALE_TEXT = {
    "scale_1_9": "1-9 标度：填 a_ij（1/9~9；1=同等，3=稍重要，5=明显重要，7=强烈重要，9=极端重要，2/4/6/8 中间值；<1 表示 j 更重要，如 1/3）。",
    "scale_0_1": "比值标度 0.1~10（两两比值）：填 i 是 j 的几倍（如 3 表示 i 比 j 重要 3 倍；0.5 表示一半）。",
    "ratio": "比值标度：直接填 a_ij（i 相对 j 的重要倍数，正数；如 2、1.5、0.5）。",
}


def catalog() -> dict:
    return lib.indicator_catalog()


def criteria_names() -> list[str]:
    return [c["name"] for c in catalog()["criteria"]]


def blocks() -> list[dict]:
    cat = catalog()
    out = []
    for c in cat["criteria"]:
        inds = [i for i in cat["indicators"] if i["criterionId"] == c["id"]]
        names = [i["name"] for i in inds]
        defs = {i["name"]: (i.get("definition") or "").strip() for i in inds}
        out.append({"id": c["id"], "name": f"{c['name']}（{len(names)} 项）",
                    "indicatorNames": names, "defs": defs})
    return out


def _style(cell, fill=None, bold=False):
    cell.border = BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if fill is not None:
        cell.fill = fill
    if bold:
        cell.font = Font(bold=True)


def _matrix_block(ws, top: int, names: list[str], scale_id: str) -> dict:
    """在 ws 的 top 行写一个 n×n 矩阵块；返回输入格/公式格坐标。"""
    n = len(names)
    ws.cell(row=top, column=1, value="项目").font = Font(bold=True)
    for j, name in enumerate(names):
        c = ws.cell(row=top, column=2 + j, value=name)
        _style(c, HEAD, bold=True)
    inputs, formulas, diagonals = [], [], []
    for i, row_name in enumerate(names):
        r = top + 1 + i
        _style(ws.cell(row=r, column=1, value=row_name), HEAD, bold=True)
        for j in range(n):
            col = 2 + j
            cell = ws.cell(row=r, column=col)
            if i == j:
                cell.value = 1
                diagonals.append(cell.coordinate)
                _style(cell, CALC)
            elif j > i:
                cell.value = None
                inputs.append(cell.coordinate)
                _style(cell, GRAY)
            else:  # 下三角：公式互反
                upper = f"{COLS[i]}{top + 1 + j}"
                cell.value = f'=IF(OR(N({upper})=0,{upper}<0),"",1/{upper})'
                formulas.append({"cell": cell.coordinate, "upper": upper})
                _style(cell, CALC)
    return {"top": top, "n": n, "inputs": inputs, "formulas": formulas, "diagonals": diagonals}


def _add_definitions(ws, top: int, names: list[str], defs: dict, n: int) -> None:
    """在矩阵块右侧写一列「指标说明」，让专家知道每个指标到底指什么。

    为什么需要：专家判断「老化程度 vs 安全风险」时若不知道指标定义，
    判断会很模糊 → 权重不可靠。定义取自指标体系（标准口径），
    **写在矩阵外面**（右侧新列），不改动 A~E 的矩阵结构，
    所以第 03 步的读取逻辑零改动。
    """
    col = 2 + n          # 矩阵最后一列是 2+n-1，说明写在 2+n
    c = ws.cell(row=top, column=col, value="指标说明（判断依据）")
    _style(c, HEAD, bold=True)
    for i, name in enumerate(names):
        cell = ws.cell(row=top + 1 + i, column=col, value=defs.get(name, ""))
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.font = Font(size=9)
        _style(cell, CALC)
    ws.column_dimensions[get_column_letter(col)].width = 46


def write_questionnaire(path: Path, expert_id: str, indicator_scale: str, criteria_scale: str,
                        project_name: str) -> dict:
    wb = Workbook()
    description = []
    ws = wb.active
    ws.title = "封面"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 76
    rows = [
        ("AHP 专家判断问卷", ""),
        ("项目", project_name),
        ("填写人", f"专家{expert_id}"),
        ("专家角色", "项目负责人 / 投资专家 / 规划专家 / 财务专家（请勾选或填写）"),
        ("日期", "____-__-__"),
        ("标度（准则层）", SCALE_TEXT.get(criteria_scale, "")),
        ("标度（指标层）", SCALE_TEXT.get(indicator_scale, "")),
        ("填写规则", "只填灰色格（矩阵上三角）；对角线固定为 1；下三角自动取倒数（请勿手改公式）。"),
        ("回收方式", "填完后整份文件交回；系统会校验缺项/越界/互反并预估 CR。"),
        ("说明", "本问卷不含真实项目名（占位为「某某片区改造项目」），产物匿名。"),
    ]
    for r, (key, value) in enumerate(rows, start=1):
        a = ws.cell(row=r, column=1, value=key)
        b = ws.cell(row=r, column=2, value=value)
        _style(a, HEAD, bold=True)
        _style(b)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)

    ws_c = wb.create_sheet("准则层")
    ws_c.cell(row=1, column=1, value="一、准则层判断矩阵（只填灰色上三角）").font = Font(bold=True)
    ws_c.cell(row=2, column=1, value=SCALE_TEXT.get(criteria_scale, "")).alignment = Alignment(wrap_text=True)
    ws_c.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    c_block = _matrix_block(ws_c, 4, criteria_names(), criteria_scale)
    ws_c.column_dimensions["A"].width = 16
    for col in "BCDE":
        ws_c.column_dimensions[col].width = 16
    for i in range(4):
        ws_c.row_dimensions[5 + i].height = 30

    ws_i = wb.create_sheet("指标层")
    ws_i.cell(row=1, column=1, value="二、指标层判断矩阵（各准则 4 项，只填灰色上三角）").font = Font(bold=True)
    ws_i.cell(row=2, column=1, value=SCALE_TEXT.get(indicator_scale, "")).alignment = Alignment(wrap_text=True)
    ws_i.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    indicator_blocks = []
    top = 4
    for idx, block in enumerate(blocks()):
        ws_i.cell(row=top, column=1,
                  value=f"（{idx + 1}）{block['name']}").font = Font(bold=True)
        info = _matrix_block(ws_i, top + 1, block["indicatorNames"], indicator_scale)
        # 矩阵右侧写指标说明（专家判断依据）；不改动矩阵本身结构
        _add_definitions(ws_i, top + 1, block["indicatorNames"], block.get("defs", {}), info["n"])
        info["criterionId"] = block["id"]
        info["sheet"] = "指标层"
        indicator_blocks.append(info)
        top += 7
    ws_i.column_dimensions["A"].width = 20
    for col in "BCDE":
        ws_i.column_dimensions[col].width = 16

    ws_v = wb.create_sheet("自动校验")
    ws_v.cell(row=1, column=1, value="三、自动校验（公式自动计算；灰色/蓝底为只读）").font = Font(bold=True)
    ws_v.cell(row=2, column=1, value="缺项/越界/互反异常 由公式统计；λmax/CI/RI/CR 为按几何平均权重的预估值，最终以系统复算为准。").alignment = Alignment(wrap_text=True)
    ws_v.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws_v.cell(row=3, column=1, value="说明：未填完或填写非法（空/0/负数/文本）时判定显示「待填」属正常；ratio 标度下「极端比值」仅提示（>9 或 <1/9），不阻断提交。").alignment = Alignment(wrap_text=True)
    ws_v.merge_cells(start_row=3, start_column=1, end_row=3, end_column=8)
    matrix_info = [{"id": "criteria", "name": "准则层", "sheet": "准则层", "scale": criteria_scale, **c_block}]
    for info in indicator_blocks:
        matrix_info.append({"id": info["criterionId"], "name": f"指标层-{info['criterionId']}",
                            "sheet": "指标层", "scale": indicator_scale, **info})
    row = 4
    for m in matrix_info:
        ws_v.cell(row=row, column=1, value=f"【{m['name']}】").font = Font(bold=True)
        row += 1
        third = "极端比值" if m["scale"] == "ratio" else "越界"
        labels = ["列和", "行几何平均", "权重", "λmax(近似)", "CI(近似)", "RI", "CR(预估)",
                  "判定", "缺项", third, "互反异常"]
        for j, label in enumerate(labels):
            _style(ws_v.cell(row=row + j, column=1, value=label), HEAD, bold=True)
        sheet = m["sheet"]
        for j in range(m["n"]):
            col = 2 + j
            bottom = m["top"] + m["n"]
            colsum = f"=IFERROR(SUM({sheet}!{COLS[j]}{m['top'] + 1}:{COLS[j]}{bottom}),\"\")"
            geo = (f"=IFERROR(GEOMEAN({sheet}!{COLS[0]}{m['top'] + 1 + j}:"
                   f"{sheet}!{COLS[m['n'] - 1]}{m['top'] + 1 + j}),\"\")")
            _style(ws_v.cell(row=row, column=col, value=colsum), CALC)
            _style(ws_v.cell(row=row + 1, column=col, value=geo), CALC)
            _style(ws_v.cell(row=row + 2, column=col,
                             value=f"=IFERROR({COLS[j]}{row + 1}/SUM($B${row + 1}:${COLS[m['n'] - 1]}${row + 1}),\"\")"), CALC)
        # 权重与列和都在本页同一行 → 直接 SUMPRODUCT（不加 TRANSPOSE，避免维度不符）
        # 门控：缺项>0 或 非法值>0 → 待填（ratio 的极端比值仅警告，不参与门控）
        if m["scale"] == "ratio":
            invalid = "+".join(
                f"IF(AND(ISNUMBER({sheet}!{c}),{sheet}!{c}<=0),1,0)" for c in m["inputs"])
            cells = ",".join(f"{sheet}!{c}" for c in m["inputs"])
            gate = f'OR(COUNT({cells})<{len(m["inputs"])},({invalid})>0)'
        else:
            gate = f'OR(B{row + 8}>0,B{row + 9}>0)'
        lam = (f'=IF({gate},"待填",'
               f'IFERROR(SUMPRODUCT(B{row + 2}:{COLS[m["n"] - 1]}{row + 2},B{row}:{COLS[m["n"] - 1]}{row}),""))')
        ws_v.cell(row=row + 3, column=2).value = lam
        _style(ws_v.cell(row=row + 3, column=2), CALC)
        _style(ws_v.cell(row=row + 4, column=2,
                         value=f'=IF(NOT(ISNUMBER(B{row + 3})),"待填",IFERROR((B{row + 3}-{m["n"]})/({m["n"]}-1),""))'), CALC)
        _style(ws_v.cell(row=row + 5, column=2,
                         value="=" + RI_CHOOSE.format(n=m["n"])), CALC)
        _style(ws_v.cell(row=row + 6, column=2,
                         value=f'=IF(NOT(ISNUMBER(B{row + 4})),"待填",IFERROR(IF(B{row + 4}=0,0,B{row + 4}/B{row + 5}),""))'), CALC)
        _style(ws_v.cell(row=row + 7, column=2,
                         value=f'=IF(NOT(ISNUMBER(B{row + 6})),"待填",IF(B{row + 6}<0.1,"✔ 可接受","✘ 需调整"))'), CALC)
        # 缺项/越界/互反异常（按输入格逐格统计；全部引用矩阵 sheet）
        low, high = ALLOWED_BY_SCALE[m['scale']]
        if m["scale"] == "ratio":
            cells_all = ",".join(f"{sheet}!{c}" for c in m["inputs"])
            missing = f"{len(m['inputs'])}-COUNT({cells_all})"
        else:
            missing = "+".join(f"COUNTBLANK({sheet}!{c})" for c in m["inputs"])
        if m["scale"] == "ratio":
            out_range = "+".join(
                f"IF(NOT(ISNUMBER({sheet}!{c})),0,IF(OR({sheet}!{c}>9,{sheet}!{c}<1/9),1,0))"
                for c in m["inputs"])
        else:
            out_range = "+".join(
                f"IF({sheet}!{c}=\"\",0,IF(NOT(ISNUMBER({sheet}!{c})),1,"
                f"IF(OR({sheet}!{c}<{low},{sheet}!{c}>{high}),1,0)))"
                for c in m["inputs"])
        recip = "+".join(
            f"IF(COUNT({sheet}!{f['upper']},{sheet}!{f['cell']})<2,0,"
            f"IF(ABS({sheet}!{f['upper']}*{sheet}!{f['cell']}-1)>0.000000001,1,0))"
            for f in m["formulas"])
        _style(ws_v.cell(row=row + 8, column=2, value=f"={missing}" if missing else "=0"), CALC)
        _style(ws_v.cell(row=row + 9, column=2, value=f"={out_range}"), CALC)
        _style(ws_v.cell(row=row + 10, column=2, value=f"={recip}" if recip else "=0"), CALC)
        description.append({"id": m["id"], "sheet": m["sheet"], "top": m["top"], "inputs": m["inputs"],
                            "formulas": m["formulas"], "diagonals": m["diagonals"]})
        row += 12

    for col in "BCDEFGH":
        ws_v.column_dimensions[col].width = 14
    ws_v.column_dimensions["A"].width = 14

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return {"path": path, "expertId": expert_id, "indicatorScale": indicator_scale,
            "criteriaScale": criteria_scale, "matrices": description}


def m_scale(m: dict) -> str:
    return m.get("scale") or "scale_1_9"


ALLOWED_BY_SCALE = {"scale_1_9": (1.0 / 9.0, 9), "scale_0_1": (0.1, 10), "ratio": (0, None)}


def read_questionnaire(path: Path, indicator_scale: str, criteria_scale: str) -> dict:
    """回收读取：返回每个矩阵的上三角输入 + 结构问题列表。"""
    wb = load_workbook(path, data_only=False)
    problems: list[dict] = []
    matrices: dict[str, dict] = {}
    spec = {"criteria": {"sheet": "准则层", "top": 4, "scale": criteria_scale}}
    cat = catalog()
    for idx, c in enumerate(cat["criteria"]):
        spec[c["id"]] = {"sheet": "指标层", "top": 5 + 7 * idx, "scale": indicator_scale}
    for mid, item in spec.items():
        ws = wb[item["sheet"]] if item["sheet"] in wb.sheetnames else None
        if ws is None:
            problems.append({"code": "AHP-E-SURVEY-SHEET", "what": f"缺 sheet {item['sheet']}"})
            continue
        n = 4
        top = item["top"]
        values = [[None] * n for _ in range(n)]
        upper_ok = True
        for i in range(n):
            for j in range(n):
                cell = ws[f"{COLS[j]}{top + 1 + i}"]
                raw = cell.value
                if i == j:
                    values[i][j] = 1.0
                    if raw not in (1, 1.0, None):
                        problems.append({"code": "AHP-E-DIAGONAL", "matrix": mid, "cell": cell.coordinate,
                                         "what": f"对角线应为 1，实际 {raw!r}"})
                    continue
                if j > i:
                    if raw is None or str(raw).strip() == "":
                        problems.append({"code": "AHP-E-SURVEY-MISSING", "matrix": mid,
                                         "cell": cell.coordinate, "what": "缺项"})
                        upper_ok = False
                        continue
                    try:
                        number = float(raw)
                    except (TypeError, ValueError):
                        problems.append({"code": "AHP-E-SURVEY-VALUE", "matrix": mid,
                                         "cell": cell.coordinate, "what": f"非数值 {raw!r}"})
                        upper_ok = False
                        continue
                    if not math.isfinite(number):
                        problems.append({"code": "AHP-E-SURVEY-VALUE", "matrix": mid,
                                         "cell": cell.coordinate, "what": f"非有限数值 {raw!r}"})
                        upper_ok = False
                        continue
                    low, high = ALLOWED_BY_SCALE[item["scale"]]
                    if number <= 0:
                        problems.append({"code": "AHP-E-SURVEY-RANGE", "matrix": mid,
                                         "cell": cell.coordinate,
                                         "what": f"必须为正数（实际 {number}）"})
                        upper_ok = False
                        continue
                    if number < low or (high is not None and number > high):
                        problems.append({"code": "AHP-E-SURVEY-RANGE", "matrix": mid,
                                         "cell": cell.coordinate,
                                         "what": f"越界 {number}（合法 {low}..{high}）"})
                    values[i][j] = number
                else:
                    if isinstance(raw, str) and raw.startswith("="):
                        values[i][j] = None  # 公式：由上三角取倒数
                    elif raw is None:
                        problems.append({"code": "AHP-E-SURVEY-MISSING", "matrix": mid,
                                         "cell": cell.coordinate, "what": "下三角为空"})
                    else:
                        try:
                            lower_value = float(raw)
                        except (TypeError, ValueError):
                            problems.append({"code": "AHP-E-SURVEY-VALUE", "matrix": mid,
                                             "cell": cell.coordinate, "what": f"非数值 {raw!r}"})
                            continue
                        upper_cell = f"{COLS[i]}{top + 1 + j}"
                        upper_value = values[j][i]
                        if upper_value is None:
                            continue
                        if lower_value <= 0:
                            problems.append({"code": "AHP-E-SURVEY-RANGE", "matrix": mid,
                                             "cell": cell.coordinate,
                                             "what": f"下三角必须为正数（实际 {lower_value}）"})
                            continue
                        if abs(upper_value * lower_value - 1.0) > 1e-9:
                            problems.append({
                                "code": "AHP-E-NONRECIPROCAL", "matrix": mid, "cell": cell.coordinate,
                                "what": f"{upper_cell}×{cell.coordinate}={upper_value * lower_value}≠1",
                            })
        if upper_ok:
            matrix = [[0.0] * n for _ in range(n)]
            for i in range(n):
                for j in range(n):
                    if i == j:
                        matrix[i][j] = 1.0
                    elif j > i:
                        matrix[i][j] = values[i][j]
                        matrix[j][i] = 1.0 / values[i][j]
                    # j<i 由上行填充
            for i in range(n):
                for j in range(n):
                    if matrix[i][j] == 0.0 and i != j:
                        problems.append({"code": "AHP-E-SURVEY-MISSING", "matrix": mid,
                                         "what": f"缺格 ({i + 1},{j + 1})"})
            matrices[mid] = {"matrix": matrix, "scale": item["scale"]}
    warnings: list[dict] = []
    for mid, item in matrices.items():
        for i in range(4):
            for j in range(i + 1, 4):
                value = item["matrix"][i][j]
                if item["scale"] in ("ratio", "scale_0_1") and (value > 9 + 1e-12 or value < 1 / 9 - 1e-12):
                    warnings.append({
                        "code": "AHP-W-SCALE-EXTREME", "matrix": mid,
                        "cell": f"{COLS[j]}{4 + 1 + i}" if mid == "criteria" else f"{COLS[j]}{5 + 7 * _block_index(mid) + 1 + i}",
                        "what": f"极端比值 {value}（超出 1/9~9，仅警告不阻断）",
                    })
    return {"matrices": matrices, "problems": problems, "warnings": warnings}


def _block_index(criterion_id: str) -> int:
    for index, c in enumerate(catalog()["criteria"]):
        if c["id"] == criterion_id:
            return index
    return 0