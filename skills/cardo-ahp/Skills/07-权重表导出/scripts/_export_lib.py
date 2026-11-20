#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_export_lib.py -- cardo-ahp 权重表导出算法（模板写入/回读/契约组装）；CLI 见 ahp_export.py。"""
from __future__ import annotations

import hashlib
import json
import math
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.pagebreak import Break

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402
import _yaml_lite as yaml_lite  # noqa: E402

TZ = timezone(timedelta(hours=8))
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"
EXPORTER_VERSION = "0.1.0"
SCHEMA = "cardo/ahp-weights/v1"
ITEM_TOL_FACTOR = 4  # 逐项比较留量化余量（12 位小数舍入 + 归一化漂移 ≈1e-12，×4 覆盖）
# 2026-10-23 根目录整理：AHP 四准则模板原件从工作根迁入「数据/标准模板/」；
# 保留旧根目录路径作兼容回退，插件独立分发时请显式传 --template。
# 2026-10-24 独立调优副本修复：原插件路径为 6 层（.../plugins/cardo-ahp/skills/cardo-ahp/scripts），
# parents[4]=仓库根；本副本只有 4 层（cardo-ahp/skill/scripts），parents[4] 直接 IndexError，
# 导致 import _export_lib 即崩（ahp_export.py / score_project.py 全挂）。改为：
# ① parents[4] 不可用时置 None 不抛；② 新增「本 skill 自带模板」候选（skill/数据/标准模板/，
# 独立副本的等价版式落点），原工作根候选保持原序优先，行为向后兼容。
_TEMPLATE_NAME = "4 层次分析法AHP四准则模板.xlsx"
_LOCAL_TEMPLATE = HERE.parent / "数据" / "标准模板" / _TEMPLATE_NAME
try:
    _WORKSPACE = HERE.parents[4]
except IndexError:
    _WORKSPACE = None
_WORKSPACE_CANDIDATES = (
    (_WORKSPACE / "数据" / "标准模板" / _TEMPLATE_NAME,
     _WORKSPACE / _TEMPLATE_NAME) if _WORKSPACE is not None else ()
)
try:
    # 包内自带模板：<包根>/assets/templates/ —— 独立分发时唯一一定存在的落点
    _PACKAGE_TEMPLATE = HERE.parents[2] / "assets" / "templates" / _TEMPLATE_NAME
except IndexError:
    _PACKAGE_TEMPLATE = None
_TEMPLATE_CANDIDATES = (*_WORKSPACE_CANDIDATES, _LOCAL_TEMPLATE,
                        *((_PACKAGE_TEMPLATE,) if _PACKAGE_TEMPLATE else ()))
ROOT_TEMPLATE = next((p for p in _TEMPLATE_CANDIDATES if p.exists()), _TEMPLATE_CANDIDATES[0])
RULES_DIR = HERE.parent / "规则"
QUANT_FILE = RULES_DIR / "指标量化分档.yaml"

CRITERIA_CELLS = ["B5", "B6", "B7", "B8"]
INDICATOR_BLOCKS = {"current": 37, "economy": 45, "compliance": 53, "risk": 61}
# 模板里每个判断矩阵块的起始行（列固定 B..E）：准则层 5，指标层四块同 INDICATOR_BLOCKS。
MATRIX_BLOCKS = {"criteria": 5, **INDICATOR_BLOCKS}
# 成品表里判断矩阵的中文名（与报告「四、一致性检验」同一套叫法）
MATRIX_LABELS = {"criteria": "准则层", "current": "现状情况", "economy": "投资经济",
                 "compliance": "规划合规", "risk": "风险可控"}
RI_TABLE = {1: 0.0, 2: 0.0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32,
            8: 1.41, 9: 1.45, 10: 1.49}


def to_matrix(values) -> list:
    """矩阵集既可能是完整矩阵，也可能只有首列（column）；统一还原成 n×n。"""
    if values and isinstance(values[0], (list, tuple)):
        return [[float(x) for x in row] for row in values]
    v = [float(x) for x in values]
    return [[a / b for b in v] for a in v]
CRITERIA_ORDER = ["current", "economy", "compliance", "risk"]
CRITERIA_SCORE_RANGE = ("B30", "E32")
INDICATOR_SCORE_RANGE = ("B95", "Q97")
RESULT_CRITERIA = ("F30", "F32")
RESULT_INDICATORS = ("R95", "R97")

def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
def now_iso() -> str: return datetime.now(TZ).isoformat(timespec="seconds")
def sha256_file(path: Path) -> str: return hashlib.sha256(path.read_bytes()).hexdigest().upper()

def sanitize_path(path_text: str) -> str:
    path = Path(path_text)
    if path.is_absolute():
        return f"<已脱敏绝对路径>/{path.name}"
    return path.as_posix()
def read_json(path_text: str) -> tuple[dict, Path]:
    path = Path(path_text)
    if not path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"文件不存在：{path_text}")
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise lib.AhpError("AHP-E-INPUT", f"不是合法 JSON：{path_text}（{exc}）") from exc
    if not isinstance(payload, dict):
        raise lib.AhpError("AHP-E-INPUT", "权重 JSON 应为对象")
    return payload, path
def normalize_weights(payload: dict) -> dict:
    """兼容 T3 weights / T7 aggregated 两种输入，返回准则与 16 项有序权重。"""
    catalog = lib.indicator_catalog()
    raw_indicators = {i.get("id"): i for i in (lib.load_indicators().get("indicators") or [])}
    criteria = payload.get("criteria")
    indicators = payload.get("indicators")
    if not isinstance(criteria, list) or not isinstance(indicators, list):
        raise lib.AhpError("AHP-E-INPUT", "权重 JSON 缺 criteria[]/indicators[]（需 T3 或 T7 产物）")
    criteria_by_id = {}
    for item in criteria:
        if not isinstance(item, dict) or "id" not in item:
            raise lib.AhpError("AHP-E-INPUT", "criteria[] 条目缺 id")
        weight = item.get("weight", item.get("localWeight"))
        if weight is None:
            raise lib.AhpError("AHP-E-INPUT", f"准则 {item['id']} 缺 weight")
        criteria_by_id[item["id"]] = float(weight)
    expected_ids = [c["id"] for c in catalog["criteria"]]
    if set(criteria_by_id) != set(expected_ids):
        raise lib.AhpError("AHP-E-INPUT", f"准则 id 与规则不一致：{sorted(criteria_by_id)} vs {expected_ids}")
    indicator_by_id = {}
    for item in indicators:
        if not isinstance(item, dict) or "id" not in item:
            raise lib.AhpError("AHP-E-INPUT", "indicators[] 条目缺 id")
        local = item.get("localWeight")
        global_weight = item.get("globalWeight")
        if local is None:
            raise lib.AhpError("AHP-E-INPUT", f"指标 {item['id']} 缺 localWeight")
        indicator_by_id[item["id"]] = (float(local), float(global_weight) if global_weight is not None else None)
    expected_indicator_ids = [i["id"] for i in catalog["indicators"]]
    if set(indicator_by_id) != set(expected_indicator_ids):
        raise lib.AhpError("AHP-E-INPUT", "指标 id 与规则不一致（应有 16 项）")
    tol = lib.rules().compare_tol
    for cid, weight in criteria_by_id.items():
        if weight <= 0:
            raise lib.AhpError("AHP-E-INPUT", f"准则 {cid} 权重必须为正（实际 {weight}；0 权重无法用比值列向量表达）")
    for iid, (local, global_weight) in indicator_by_id.items():
        if local <= 0:
            raise lib.AhpError("AHP-E-INPUT", f"指标 {iid} 局部权重必须为正（实际 {local}）")
        expected = criteria_by_id.get(iid.split(".")[0], 0.0) * local
        if global_weight is not None and abs(global_weight - expected) > tol * ITEM_TOL_FACTOR:
            raise lib.AhpError( "AHP-E-INPUT",
                f"指标 {iid} 的 globalWeight={global_weight} ≠ 准则×局部={expected}（拒绝自相矛盾的输入）",
            )
    if abs(sum(criteria_by_id.values()) - 1.0) > tol * len(criteria_by_id):
        raise lib.AhpError("AHP-E-INPUT", f"准则权重和 {sum(criteria_by_id.values())} ≠ 1（超量化余量）")
    total_global = sum((g if g is not None else criteria_by_id[i.split(".")[0]] * l)
                       for i, (l, g) in indicator_by_id.items())
    if abs(total_global - 1.0) > tol * len(indicator_by_id):
        raise lib.AhpError("AHP-E-INPUT", f"综合权重和 {total_global} ≠ 1（超量化余量）")
    return {
        "criteria": criteria_by_id,
        "indicators": indicator_by_id,
        "criteriaOrder": expected_ids,
        "indicatorOrder": expected_indicator_ids,
        "catalog": catalog,
        "rawIndicators": raw_indicators,
        "source": payload,
    }
def build_full_payload(args, data: dict, source_path: Path, template_path: Path, out_path: Path, out_json: Path | None) -> dict:
    catalog = data["catalog"]
    rules = lib.load_scale_rules()
    quant = yaml_lite.load_file(QUANT_FILE)
    quant_items = {item.get("id"): item for item in quant.get("items", [])}
    source_payload = data["source"]
    criteria_out = []
    for c in catalog["criteria"]:
        criteria_out.append({
            "id": c["id"], "name": c["name"], "weight": data["criteria"][c["id"]],
            "lambdaMax": None, "ci": None, "cr": None, "pass": None,
        })
    consistency_in = dict(source_payload.get("consistency") or {})
    if not consistency_in.get("perMatrix") and isinstance(consistency_in.get("matrices"), list):
        consistency_in["perMatrix"] = {
            m.get("id"): {"lambdaMax": m.get("lambdaMax"), "ci": m.get("ci"), "ri": m.get("ri"), "cr": m.get("cr"), "pass": m.get("pass"),
                          "n": m.get("n"), "threshold": consistency_in.get("threshold")}
            for m in consistency_in["matrices"] if isinstance(m, dict) and m.get("id")
        }
    per_matrix = dict(consistency_in.get("perMatrix") or {})
    for entry in (consistency_in.get("matrices") or []):
        if isinstance(entry, dict) and entry.get("id"):
            per_matrix[entry["id"]] = entry
    source_matrices = source_payload.get("matrices") or {}
    criteria_level = dict(per_matrix.get("criteria") or {})
    fallback = source_matrices.get("criteria") or {}
    for key in ("lambdaMax", "ci", "cr", "pass"):
        criteria_level[key] = criteria_level.get(key) if criteria_level.get(key) is not None else fallback.get(key)
    for item in criteria_out:
        item.update({k: criteria_level.get(k) for k in ("lambdaMax", "ci", "cr", "pass")})
    indicators_out = []
    for ind in catalog["indicators"]:
        raw = data["rawIndicators"].get(ind["id"]) or {}
        item = quant_items.get(ind["id"]) or {}
        local, global_weight = data["indicators"][ind["id"]]
        indicators_out.append({
            "id": ind["id"], "criterionId": ind["criterionId"], "name": ind["name"],
            "definition": raw.get("definition"), "direction": raw.get("direction"),
            "directionNote": raw.get("directionNote"),
            "source": raw.get("source"), "required": True,
            "templateCell": ind.get("templateCell"),
            "anchors": ind.get("anchors") or [],
            "quantification": {
                "metric": item.get("metric"), "unit": item.get("unit"),
                "formula": item.get("formula"), "metricDirection": item.get("metricDirection"),
                "interpolation": item.get("interpolation"),
            },
            "quantBands": {"bands": item.get("bands") or []},
            "localWeight": local,
            "globalWeight": global_weight if global_weight is not None else
                            data["criteria"][ind["criterionId"]] * local,
        })
    sum_criteria = sum(criteria_out[i]["weight"] for i in range(len(criteria_out)))
    sum_global = sum(i["globalWeight"] for i in indicators_out)
    gaps = []
    if getattr(args, "direction", None) is None:
        gaps.append({"code": "AHP-GAP-DIRECTION", "what": "导出时未指定项目方向（direction=generic）", "need": "由 cardo-status 提供方向或导出时 --direction 指定", "owner": "cardo-status / 用户"})
    if abs(sum_global - 1.0) > lib.rules().compare_tol:
        gaps.append({"code": "AHP-GAP-WEIGHT-SUM", "what": f"综合权重和 {sum_global} ≠ 1"})
    payload = {
        "schema": SCHEMA, "schemaVersion": 1, "producer": PRODUCER,
        "source": [{"kind": "weights", "id": "weights", "path": sanitize_path(str(source_path)),
                    "sha256": sha256_file(source_path)}],
        "gaps": gaps,
        "ruleVersion": int(rules.get("version") or 0),
        "engineVersion": ENGINE_VERSION,
        "exporterVersion": EXPORTER_VERSION,
        "direction": {"type": args.direction or "generic", "basis": "导出参数" if args.direction else "未指定（v1 导出占位）"},
        "directionApplicability": {"criterionRenormalized": []},
        "scaleProfile": {"criteria": "scale_1_9", "indicators": "scale_0_1", "aggregation": "geometric-elementwise"},
        "experts": source_payload.get("experts") if isinstance(source_payload.get("experts"), list) else [],
        "calibration": {"applied": bool((source_payload.get("calibrationApplied"))), "records": []},
        "display": {"format": "xlsx", "template": template_path.name, "templateSha256": sha256_file(template_path), "sheet": "层次分析模板"},
        "outputs": {"xlsx": sanitize_path(str(out_path)),
                    "json": sanitize_path(str(out_json)) if out_json else None},
        "criteria": criteria_out,
        "indicators": indicators_out,
        "weights": {
            "unit": "fraction", "tolerance": lib.rules().compare_tol,
            "sumCriteria": sum_criteria, "sumGlobal": sum_global,
            "sumLocalPerCriterion": {c["id"]: sum(i["localWeight"] for i in indicators_out
                                                  if i["criterionId"] == c["id"])
                                     for c in catalog["criteria"]},
        },
        "methods": source_payload.get("methods") or {"primary": "geometric", "compare": "eigen"},
        "consistency": consistency_in or {},
        "scores": {},
        "demo": bool(getattr(args, "demo", False)),
    }
    if not getattr(args, "no_timestamp", False):
        payload["generatedAt"] = now_iso()
    return payload
def apply_print_layout(wb) -> None:
    """权重表打印版式：标题收进 A:F、浮点 6 位、A4 横向 1 页宽、表头跨页重复。

    只动版式，不动任何数值与公式；模板结构（矩阵格位、合并关系）保持可回读。
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    tpl = wb["层次分析模板"] if "层次分析模板" in wb.sheetnames else wb.worksheets[0]
    for rng in list(tpl.merged_cells.ranges):
        if rng.min_col == 1 and rng.max_col > 6:
            tpl.unmerge_cells(str(rng))
            tpl.merge_cells(start_row=rng.min_row, start_column=1,
                            end_row=rng.max_row, end_column=6)
    for col in range(7, 19):
        tpl.column_dimensions[get_column_letter(col)].width = 3
    for row in tpl.iter_rows(min_row=1, max_row=tpl.max_row, max_col=6):
        for c in row:
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                if float(c.value) != int(c.value):
                    c.number_format = "0.000000"   # 16 位小数会撑成 ###
    # A2「使用方法」整句（56 字，约 108–112 宽度单位）必须一行放下：
    # 宋体 11pt 一个汉字实际占 ~2.1 单位，模板原总宽 104.8 略窄，会把句号「。」
    # 挤到第二行行首。这里按内容补足宽度（不改字、不删字）。
    for col, width in (("A", 21.5), ("B", 18.0), ("C", 25.0),
                       ("D", 16.5), ("E", 17.0), ("F", 14.0)):
        tpl.column_dimensions[col].width = width
    tpl.row_dimensions[1].height = 34
    tpl.row_dimensions[2].height = 36
    # CR 判定格是两行文字（B26「✔ 一致性可接受」、F88:F91「0.0000 ✔ 可接受」），
    # 默认行高只有一行，打印时会压到上下相邻行 —— 机检重叠 5 处的来源。
    tpl.row_dimensions[26].height = 30
    for r in range(88, 92):
        tpl.row_dimensions[r].height = 30
    tpl.print_area = f"A1:F{tpl.max_row}"
    # 只在子块边界分页：p1 = 一~四（自然分页，行 33 正好满页）；
    # p2 = 五 + 现状/经济/交付三组矩阵；p3 = 风险矩阵 + 六 + 一致性。自动分页会切在
    # 矩阵中间（p2 从「工程延误风险」半截开始、尾页只剩十几行），跨页表格又没有表头可读。
    for rid in (58, 85):
        tpl.row_breaks.append(Break(id=rid))

    sheets = [(tpl, "1:2", 0)]
    if "导出值" in wb.sheetnames:
        out = wb["导出值"]
        # 键名最长 31.5 单位（指标综合.economy.appreciation），26 会被 B 列值截掉尾巴；
        # 按内容给到 32.5，并允许折行兜底：宁可行高增加，不许砍字。
        out.column_dimensions["A"].width = 32.5
        out.column_dimensions["B"].width = 22
        for row in out.iter_rows(min_row=3, min_col=1, max_col=1):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="center")
        for row in out.iter_rows(min_row=2, min_col=2, max_col=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "0.000000"
        # 键名 37 行：按 1 页硬压会把 11pt 压到 9.5pt；改为分两页、在中点断
        # （不留「尾页只剩几行」的碎片），字号保持 11pt。
        sheets.append((out, "1:2", 0))
        out.row_breaks.append(Break(id=2 + (out.max_row - 2 + 1) // 2))
    for sh, titles, tall in sheets:
        sh.page_setup.orientation = "landscape"
        sh.page_setup.paperSize = sh.PAPERSIZE_A4
        sh.page_setup.fitToWidth = 1
        # 模板不按高度硬压（压到 3 页会把正文压到 9.5pt 以下）；导出值仍按 1 页收口
        sh.page_setup.fitToHeight = tall
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sh.print_title_rows = titles
        sh.page_margins.left = sh.page_margins.right = 0.4
        sh.page_margins.top = 0.5
        sh.page_margins.bottom = 0.75
        sh.page_margins.footer = 0.3


COVER_ORG = "Cardo 城市更新 AHP 项目组"
_MAT_ORDER = ["criteria", "current", "economy", "compliance", "risk"]
_SCALE_TEXT = {
    "criteria": "准则层 1–9 标度",
    "current": "指标层 0.1–10 比值标度",
    "economy": "指标层 0.1–10 比值标度",
    "compliance": "指标层 0.1–10 比值标度",
    "risk": "指标层 0.1–10 比值标度",
}


def _principal_lambda(mat: list) -> float:
    """幂迭代求最大特征值（不依赖 numpy；与引擎的特征值法同解）。"""
    n = len(mat)
    x = [1.0 / n] * n
    for _ in range(500):
        y = [sum(mat[i][j] * x[j] for j in range(n)) for i in range(n)]
        s = sum(y)
        if s <= 0:
            break
        nx = [v / s for v in y]
        if max(abs(a - b) for a, b in zip(nx, x)) < 1e-15:
            x = nx
            break
        x = nx
    ax = [sum(mat[i][j] * x[j] for j in range(n)) for i in range(n)]
    den = sum(v * v for v in x)
    return sum(x[i] * ax[i] for i in range(n)) / den if den else float(n)


def matrix_stats(mat: list) -> dict:
    n = len(mat)
    lam = _principal_lambda(mat)
    ci = (lam - n) / (n - 1) if n > 2 else 0.0
    ri = RI_TABLE.get(n, 0.0)
    cr = ci / ri if ri else 0.0
    return {"n": n, "lambdaMax": lam, "ci": ci, "ri": ri, "cr": cr,
            "pass": cr < 0.1}


def _cover_sheet(wb, title: str, project: str, org: str, version: str, generated: str):
    from openpyxl.styles import Alignment, Font
    ws = wb.active
    ws.title = "封面"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 56
    ws["A2"] = title
    ws["A2"].font = Font(size=26, bold=True)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("A2:B2")
    ws.row_dimensions[2].height = 46
    rows = [("项目名称", project), ("成果名称", title + "（成品版）"),
            ("编制单位", org), ("版本", version), ("生成时间", generated)]
    for i, (k, v) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=k).font = Font(size=12, bold=True)
        cell = ws.cell(row=i, column=2, value=v)
        cell.font = Font(size=12)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[i].height = 22
    return ws


def _page_setup(wb, landscape: bool = True) -> None:
    for ws in wb.worksheets:
        portrait = ws.title == "封面"
        ws.page_setup.orientation = "portrait" if portrait else ("landscape" if landscape else "portrait")
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.6
        ws.page_margins.bottom = 0.6
        ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
        ws.oddFooter.center.size = 9


def _delta_note(data) -> str:
    """方法说明里报本次两法偏差（取自 权重表.json 的 methods.delta），说明与算法同口径。"""
    items = (data.get("methods") or {}).get("delta") or []
    vals = []
    for item in items:
        if isinstance(item, dict) and isinstance(item.get("maxAbsDelta"), (int, float)):
            vals.append(abs(float(item["maxAbsDelta"])))
    if not vals:
        return ""
    return f"本次 {len(vals)} 个矩阵实测最大 Δ = {max(vals):.2e}，均满足判据。"


def write_clean_workbook(out_path: Path, data: dict, matrix_set: dict | None, *,
                         project: str, generated: str, org: str = COVER_ORG,
                         version: str = "v1.0") -> None:
    """成品《AHP权重表》：静态值、零公式、无模板残留，含封面与可自证的方法说明。

    R1 P0-1：不再拷贝模板，示例方案区（方案一/二/三 + 示例得分）与
    fullCalcOnLoad 二次计算一并断根；所有数字由本次 05 产物直接写入。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_head = PatternFill("solid", fgColor="1F4E79")
    fill_alt = PatternFill("solid", fgColor="EAF1F8")

    def head(ws, row, values, width_from=1):
        for col, v in enumerate(values, start=width_from):
            c = ws.cell(row=row, column=col, value=v)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = fill_head
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border

    def body(ws, row, values, fmts=None, start=1):
        for i, v in enumerate(values):
            c = ws.cell(row=row, column=start + i, value=v)
            c.border = border
            c.font = Font(size=10)
            c.alignment = Alignment(vertical="center", wrap_text=True,
                                    horizontal="center" if i else "left")
            if fmts and i < len(fmts) and fmts[i]:
                c.number_format = fmts[i]

    catalog = data.get("catalog") or {}
    crit_names = {c["id"]: c.get("name") or c["id"] for c in catalog.get("criteria", [])}
    ind_names = {i["id"]: i.get("name") or i["id"] for i in catalog.get("indicators", [])}
    mats = {m.get("id"): to_matrix(m.get("values") or [])
            for m in (matrix_set or {}).get("matrices", []) if m.get("id")}
    stats = {mid: matrix_stats(mat) for mid, mat in mats.items()}

    wb = Workbook()
    _cover_sheet(wb, "AHP 权重表", project, org, version, generated)

    # ---------- 权重表 ----------
    ws = wb.create_sheet("权重表")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:D1")
    ws["A1"] = "AHP 权重表"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    ws["A3"] = "一、准则层权重"
    ws["A3"].font = Font(bold=True, size=12)
    head(ws, 4, ["准则", "权重"])
    r = 5
    for cid in data["criteriaOrder"]:
        body(ws, r, [crit_names.get(cid, cid), round(data["criteria"][cid], 6)],
             ["", "0.000000"])
        if (r - 5) % 2 == 1:
            for c in range(1, 3):
                ws.cell(row=r, column=c).fill = fill_alt
        r += 1

    ws.cell(row=r + 1, column=1, value="二、指标层权重（局部权重 × 准则权重 = 综合权重）").font = Font(bold=True, size=12)
    head(ws, r + 2, ["准则", "指标", "局部权重", "综合权重"])
    r = r + 3
    for iid in data["indicatorOrder"]:
        cid = iid.split(".")[0]
        local, glob = data["indicators"][iid]
        body(ws, r, [crit_names.get(cid, cid), ind_names.get(iid, iid),
                     round(local, 6), round(glob, 6)],
             ["", "", "0.000000", "0.000000"])
        if (r % 2) == 0:
            for c in range(1, 5):
                ws.cell(row=r, column=c).fill = fill_alt
        r += 1

    ws.cell(row=r + 1, column=1, value="三、一致性检验（CR < 0.1 判定通过）").font = Font(bold=True, size=12)
    head(ws, r + 2, ["判断矩阵", "阶数 n", "λmax", "CI", "RI", "CR", "判定"])
    r = r + 3
    for mid in _MAT_ORDER:
        st = stats.get(mid)
        if not st:
            continue
        body(ws, r, [MATRIX_LABELS.get(mid, mid), st["n"], round(st["lambdaMax"], 4),
                     round(st["ci"], 6), st["ri"], round(st["cr"], 6),
                     "通过" if st["pass"] else "不通过"],
             ["", "0", "0.0000", "0.000000", "0.00", "0.000000", ""])
        r += 1
    ws.cell(row=r + 1, column=1,
            value="判断矩阵与标度含义见「判断矩阵」「方法说明」两页；λmax 取最大特征值，CR<0.1 视为一致性可接受。"
            ).font = Font(size=9, italic=True)
    for col, w in enumerate([16, 30, 14, 14, 8, 12, 10], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ---------- 判断矩阵 ----------
    ws2 = wb.create_sheet("判断矩阵")
    ws2.sheet_view.showGridLines = False
    ws2["A1"] = "判断矩阵（群体矩阵：多位专家判断逐元素几何平均；a_ji = 1 / a_ij）"
    ws2["A1"].font = Font(size=14, bold=True)
    row = 3
    for k, mid in enumerate(_MAT_ORDER, start=1):
        mat = mats.get(mid)
        if not mat:
            continue
        st = stats[mid]
        if mid == "criteria":
            labels = [crit_names.get(c, c) for c in data["criteriaOrder"]]
        else:
            labels = [ind_names.get(i, i) for i in data["indicatorOrder"] if i.startswith(mid + ".")]
        ws2.cell(row=row, column=1, value=f"（{k}）{MATRIX_LABELS.get(mid, mid)}判断矩阵（{_SCALE_TEXT[mid]}）"
                 ).font = Font(bold=True, size=11)
        row += 1
        head(ws2, row, [""] + labels)
        row += 1
        for i, line in enumerate(labels):
            c = ws2.cell(row=row, column=1, value=line)
            c.font = Font(size=9, bold=True)
            c.border = border
            c.alignment = Alignment(vertical="center")
            for j in range(len(labels)):
                d = ws2.cell(row=row, column=2 + j, value=round(float(mat[i][j]), 4))
                d.number_format = "0.0000"
                d.border = border
                d.font = Font(size=9)
                d.alignment = Alignment(horizontal="center")
            row += 1
        ws2.cell(row=row, column=1,
                 value=f"λmax = {st['lambdaMax']:.4f}　CI = {st['ci']:.6f}　"
                       f"RI = {st['ri']:.2f}　CR = {st['cr']:.6f}　判定：{'通过' if st['pass'] else '不通过'}"
                 ).font = Font(size=9)
        ws2.column_dimensions["A"].width = 26
        for j in range(len(labels)):
            ws2.column_dimensions[get_column_letter(2 + j)].width = 11
        row += 2

    # ---------- 方法说明 ----------
    ws3 = wb.create_sheet("方法说明")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 120
    lines = [
        "方法说明",
        "",
        "1. 标度。准则层采用 1–9 标度：1 表示两两同等重要，3 略重要，5 明显重要，7 强烈重要，9 极端重要，"
        "2/4/6/8 为相邻判断的中间值；a_ij 表示 i 相对 j 的重要度，a_ji = 1 / a_ij。"
        "指标层采用 0.1–10 比值标度：输入为 i 相对 j 的重要性倍数（0.1 表示 i 远不如 j，1 表示同等，"
        "10 表示 i 远强于 j），a_ji = 1 / a_ij；取值粒度 0.1，适配 4 项指标内的比较。",
        "2. 权重。对 n 阶判断矩阵按行求几何平均 g_i = (Π_j a_ij)^(1/n)，归一化得 w_i = g_i / Σ g_j；"
        "同时用最大特征值对应的特征向量复核，两法逐元素偏差 Δ = max|w_i(几何平均) − w_i(特征向量)|，"
        "判据为 Δ ≤ max(1e-12, CR)，满足即取几何平均结果。" + _delta_note(data),
        "3. 一致性。λmax 为判断矩阵的最大特征值；CI = (λmax − n) / (n − 1)；"
        "RI 为同阶随机矩阵平均一致性指标（n=3 取 0.58，n=4 取 0.90）；CR = CI / RI。"
        "CR < 0.1 视为一致性可接受，否则该矩阵需要专家重判。",
        "4. 聚合。多位专家对同一矩阵逐元素取几何平均后再按上述步骤计算，"
        "以降低个别极端判断对整体权重的影响。",
        "5. 阅读顺序。准则层权重给出四个维度的相对重要性；指标层局部权重 × 所属准则权重 = 综合权重；"
        "综合权重直接进入《AHP方案评分结论》与每个方案得分的加权汇总。",
    ]
    for i, text in enumerate(lines, start=1):
        c = ws3.cell(row=i, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(size=11 if i == 1 else 10, bold=(i == 1))
        ws3.row_dimensions[i].height = 24 if i == 1 else (48 if len(text) > 60 else 22)

    _page_setup(wb, landscape=True)
    # 成品不允许 Excel 打开时二次计算：零公式 + 关闭强制重算
    try:
        wb.calculation.fullCalcOnLoad = False
    except Exception:
        pass
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def check_consistency(workbook_path: Path, json_path: Path,
                      matrix_set_path: Path) -> tuple[int, dict]:
    """成品表内 CR 与报告 JSON、与矩阵集复算逐项比对；不一致 exit 3。"""
    from openpyxl import load_workbook as _lw
    payload, _ = read_json(json_path)
    ms = json.loads(Path(matrix_set_path).read_text(encoding="utf-8-sig"))
    mats = {m.get("id"): to_matrix(m.get("values") or [])
            for m in ms.get("matrices", []) if m.get("id")}
    per = ((payload.get("consistency") or {}).get("perMatrix") or {})
    wb = _lw(workbook_path, data_only=True)
    ws = wb["权重表"]
    label2id = {v: k for k, v in MATRIX_LABELS.items()}
    items, bad = [], 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        name = row[0].value
        if not isinstance(name, str) or name not in label2id:
            continue
        mid = label2id[name]
        table_lambda = row[2].value
        table_cr = row[5].value
        if not isinstance(table_cr, (int, float)):
            continue
        calc = matrix_stats(mats[mid]) if mid in mats else None
        rep = per.get(mid) or {}
        rep_cr = rep.get("cr")
        d1 = abs(float(table_cr) - float(rep_cr)) if isinstance(rep_cr, (int, float)) else None
        d2 = abs(float(table_cr) - calc["cr"]) if calc else None
        ok = (d1 is not None and d1 <= 5e-6) and (d2 is not None and d2 <= 5e-6)
        bad += 0 if ok else 1
        items.append({"matrix": name, "id": mid,
                      "tableLambda": table_lambda, "calcLambda": round(calc["lambdaMax"], 6) if calc else None,
                      "tableCR": table_cr, "reportCR": rep_cr,
                      "recomputedCR": round(calc["cr"], 6) if calc else None,
                      "deltaTableVsReport": d1, "deltaTableVsRecompute": d2, "pass": ok})
    ok = bad == 0 and len(items) == len(_MAT_ORDER)
    # 人读表走 stderr：stdout 只留 JSON，pipeline 直接把它落成 CR一致性核对.json
    print("矩阵\t表内CR\t报告CR\t复算CR\t判定", file=sys.stderr)
    for it in items:
        print(f"{it['matrix']}\t{it['tableCR']:.6f}\t"
              f"{it['reportCR']:.6f}\t{it['recomputedCR']:.6f}\t{'OK' if it['pass'] else 'NG'}",
              file=sys.stderr)
    return (0 if ok else 3), {"ok": ok, "items": items}


def write_workbook(template_path: Path, out_path: Path, data: dict, scores: dict | None,
                   demo: bool = False, matrix_set: dict | None = None) -> None:
    """模板底稿（溯源版）：真实完整矩阵写入 + 清掉模板自带的示例方案区。

    2026-11-14 修（R1 P0-2）：原先只写首列 w_i/w_1，其余格仍是「=B5/B6」这类
    还原公式，指标层四块又各自成比例 —— 打开后 λmax 全 4.0000、CR=0，与引擎
    算出的 0.0023 对不上。现按 05 的 full 矩阵逐格写入，表内公式算出的
    λmax/CI/CR 与报告同源。
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(template_path.read_bytes())
    wb = load_workbook(out_path)
    ws = wb["层次分析模板"]
    criteria_values = [data["criteria"][cid] for cid in data["criteriaOrder"]]
    if matrix_set and matrix_set.get("matrices"):
        by_id = {m.get("id"): to_matrix(m.get("values") or []) for m in matrix_set["matrices"]}
        for mid, top in MATRIX_BLOCKS.items():
            mat = by_id.get(mid)
            if not mat:
                continue
            for r, row in enumerate(mat):
                for c, value in enumerate(row, start=2):      # B 列起
                    ws.cell(row=top + r, column=c).value = float(value)
    else:
        # 兼容旧输入：没有矩阵集时退回列向量口径（首列 w_i/w_1，首格必须为 1）
        base = criteria_values[0]
        for cell, value in zip(CRITERIA_CELLS, [v / base for v in criteria_values]):
            ws[cell].value = value
        for cid in data["criteriaOrder"]:
            top = INDICATOR_BLOCKS[cid]
            locals_ = [data["indicators"][iid][0] for iid in data["indicatorOrder"]
                       if iid.startswith(cid + ".")]
            base = locals_[0]
            for offset, value in enumerate([v / base for v in locals_]):
                ws[f"B{top + offset}"].value = value

    # 模板自带的「方案一/二/三 + 示例得分」两段（四、七）与本次数据无关：
    # 留着会被 fullCalcOnLoad 重算成另一套排序（实测 7.1521/6.5129/6.3991，
    # 与报告的 7.26/6.09/5.11 相反）。整块清空（含标题行下的表头行 29/94），
    # 标题改成指向评分结论。
    for title_row, first_row, last_row, last_col in ((28, 29, 32, 6), (93, 94, 97, 18)):
        ws.cell(row=title_row, column=1).value = (
            "四、方案层综合评估：见《AHP方案评分结论》（本表不再重复列示方案得分）"
            if title_row == 28 else
            "七、方案层综合评估：见《AHP方案评分结论》（本表不再重复列示方案得分）")
        for r in range(first_row, last_row + 1):
            for c in range(1, last_col + 1):
                ws.cell(row=r, column=c).value = None
    if "导出值" not in wb.sheetnames:
        ws_out = wb.create_sheet("导出值")
        ws_out.append(["AHP 权重表 · 导出值（与计算区逐项对照）"])
        ws_out.merge_cells("A1:B1")
        title = ws_out["A1"]
        title.font = Font(bold=True, color="FFFFFF")
        title.fill = PatternFill("solid", fgColor="1F4E79")
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws_out.row_dimensions[1].height = 26
        ws_out.append(["项", "值"])
        for item in data["criteriaOrder"]:
            ws_out.append([f"准则权重.{item}", data["criteria"][item]])
        for item in data["indicatorOrder"]:
            ws_out.append([f"指标局部.{item}", data["indicators"][item][0]])
            ws_out.append([f"指标综合.{item}", data["indicators"][item][1]])
    # 两个 sheet 都要页码（页脚居中）。演示身份由 07/09 的过程 JSON 携带、
    # 由结论稿的方法说明交代一次；交付 xlsx 不写「demo=true」这类内部标记，
    # 也不在每页页脚撒角标（成品版对外交付）。
    footer_text = "第 &P 页 / 共 &N 页"
    for sh in wb.worksheets:
        sh.oddFooter.center.text = footer_text
        sh.oddFooter.center.size = 11
    apply_print_layout(wb)
    wb.save(out_path)
def export_weights(args) -> int:
    if args.direction is not None:
        allowed = direction_ids()
        if allowed and args.direction not in allowed and args.direction != "generic":
            raise lib.AhpError("AHP-E-INPUT", f"未知方向 {args.direction}（可选 {allowed} 或 generic）")
    data = normalize_weights(read_json(args.weights)[0])
    template_path = Path(args.template)
    if not template_path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"模板不存在：{args.template}")
    template_sha = sha256_file(template_path)
    out_path = Path(args.out)
    if out_path.exists() and out_path.is_dir():
        raise lib.AhpError("AHP-E-INPUT", f"--out 是目录，应写具体 .xlsx 文件：{args.out}")
    if out_path.resolve() == template_path.resolve():
        raise lib.AhpError("AHP-E-INPUT", "--out 不能指向 --template 模板本身（会写坏模板）")
    out_json = Path(args.out_json) if args.out_json else out_path.with_suffix(".json")
    scores = read_json(args.scores)[0] if args.scores else None
    matrix_set = read_json(args.matrix_set)[0] if getattr(args, "matrix_set", None) else None
    project = getattr(args, "project", None) or "某某片区改造项目"
    generated = datetime.now(TZ).strftime("%Y-%m-%d %H:%M")
    payload = build_full_payload(args, data, Path(args.weights), template_path, out_path, out_json)
    # 成品：干净工作簿（封面 + 权重表 + 判断矩阵 + 方法说明），静态值、零公式
    write_clean_workbook(out_path, data, matrix_set, project=project, generated=generated,
                         org=getattr(args, "org", None) or COVER_ORG,
                         version=getattr(args, "version", None) or "v1.0")
    # 溯源底稿：模板副本 + 完整矩阵 + 清空示例方案区（供回读对账）
    ann_path = Path(args.out_ann) if getattr(args, "out_ann", None) else None
    if ann_path:
        write_workbook(template_path, ann_path, data, scores,
                       demo=bool(getattr(args, "demo", False)), matrix_set=matrix_set)
    if sha256_file(template_path) != template_sha:
        raise lib.AhpError("AHP-E-INPUT", "模板原件在导出过程中被改动（应只读）", 1)
    if scores and ann_path:
        recomputed = recompute_from_workbook(ann_path)
        payload["scores"] = {"criteriaScores": scores.get("criteriaScores"),
                             "indicatorScores": scores.get("indicatorScores"),
                             "criteriaResults": recomputed["scores"]["criteriaResults"],
                             "indicatorResults": recomputed["scores"]["indicatorResults"]}
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    result = {"xlsx": str(out_path), "json": str(out_json), "schema": SCHEMA,
              "templateSha256": template_sha, "sumGlobal": payload["weights"]["sumGlobal"],
              "templateUnchanged": True}
    if ann_path:
        result["xlsxAnn"] = str(ann_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0
def direction_ids() -> list[str]:
    path = RULES_DIR / "方向适配.yaml"
    if not path.is_file():
        return []
    directions = yaml_lite.load_file(path).get("directions") or []
    if isinstance(directions, dict):
        return [str(k) for k in directions]
    return [str(d.get("id")) for d in directions if isinstance(d, dict) and d.get("id")]
def row_weights(vector: list[float]) -> list[float]:
    n = len(vector); geo = [math.prod([vector[i] / vector[j] for j in range(n)]) ** (1.0 / n) for i in range(n)]
    return [g / sum(geo) for g in geo]
def expected_scores(rows: list[list[float]], weights: list[float]) -> list[float]:
    return [sum(float(v) * w for v, w in zip(row, weights)) for row in rows]
def recompute_from_workbook(path: Path) -> dict:
    wb = load_workbook(path)
    ws = wb["层次分析模板"]
    def _block(top: int, n: int) -> list[float]:
        """按 B..列整块读回，逐行几何平均 —— 与 05 的 full 矩阵口径一致。

        旧实现只读首列 a_i1 再归一：矩阵略有非一致（CR>0）时首列比值
        与全矩阵几何平均并不相等，回读会被误判成"权重不一致"。
        """
        rows = [[float(ws.cell(row=top + r, column=2 + c).value) for c in range(n)]
                for r in range(n)]
        gm = []
        for row in rows:
            prod_ = 1.0
            for v in row:
                prod_ *= v
            gm.append(prod_ ** (1.0 / n))
        total = sum(gm) or 1.0
        return [g / total for g in gm]

    criteria = _block(5, 4)
    locals_map = {cid: _block(top, 4) for cid, top in INDICATOR_BLOCKS.items()}
    def _score_rows(first: int, last: int, width: int) -> list[list[float]]:
        """示例方案区已被清空（R1 P0-1）：空行跳过，不再当成 0 分参与回读。"""
        rows = []
        for r in range(first, last):
            vals = [ws.cell(row=r, column=c).value for c in range(2, 2 + width)]
            if any(v is None for v in vals):
                continue
            rows.append([float(v) for v in vals])
        return rows

    crit_rows = _score_rows(30, 33, 4)
    ind_rows = _score_rows(95, 98, 16)
    global_weights = [criteria[index] * w for index, cid in enumerate(CRITERIA_ORDER)
                      for w in locals_map[cid]]
    return {"criteria": criteria, "locals": locals_map,
            "scores": {"criteriaScores": crit_rows, "indicatorScores": ind_rows,
                       "criteriaResults": expected_scores(crit_rows, criteria),
                       "indicatorResults": expected_scores(ind_rows, global_weights)}}
def verify_weights(args) -> int:
    payload, _ = read_json(args.json)
    if payload.get("schema") != SCHEMA:
        raise lib.AhpError("AHP-E-INPUT", f"不是 {SCHEMA}：{payload.get('schema')}")
    xlsx = Path(args.file)
    if not xlsx.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"xlsx 不存在：{args.file}")
    criteria_json = payload.get("criteria")
    indicators_json = payload.get("indicators")
    if not isinstance(criteria_json, list) or not isinstance(indicators_json, list):
        raise lib.AhpError("AHP-E-INPUT", "权重 JSON 缺 criteria[]/indicators[]，无法回读对账")
    template_path = Path(args.template)
    template_ok = template_path.is_file() and payload.get("display", {}).get("templateSha256") == sha256_file(template_path)
    catalog = lib.indicator_catalog()
    try:
        recomputed = recompute_from_workbook(xlsx)
    except (KeyError, ValueError, TypeError, AttributeError) as exc:
        raise lib.AhpError("AHP-E-INPUT",
                           f"xlsx 结构不符合模板版式，无法回读（{type(exc).__name__}）") from exc
    tol = lib.rules().compare_tol
    # JSON 侧同口径归一化（xlsx 只存比值列向量，回读必然重归一化；12 位小数舍入会让 Σ 天然偏离）
    criteria_raw = {c.get("id"): c.get("weight") for c in criteria_json}
    sum_criteria = sum(float(v) for v in criteria_raw.values() if v is not None) or 1.0
    expected_criteria = {cid: (float(v) / sum_criteria if v is not None else None) for cid, v in criteria_raw.items()}
    local_raw = {i.get("id"): i.get("localWeight") for i in indicators_json}
    local_sums = {c["id"]: (sum(float(local_raw[i["id"]]) for i in catalog["indicators"]
                  if i["criterionId"] == c["id"] and local_raw.get(i["id"]) is not None) or 1.0)
                  for c in catalog["criteria"]}
    problems = []
    for index, c in enumerate(catalog["criteria"]):
        if c["id"] not in expected_criteria or expected_criteria[c["id"]] is None:
            problems.append(f"JSON 缺准则 {c['id']} 或权重为空")
            continue
        if abs(recomputed["criteria"][index] - expected_criteria[c["id"]]) > tol * ITEM_TOL_FACTOR:
            problems.append(f"准则 {c['id']} 权重回读不一致")
    for c in catalog["criteria"]:
        local_ids = [i["id"] for i in catalog["indicators"] if i["criterionId"] == c["id"]]
        for index, iid in enumerate(local_ids):
            value = local_raw.get(iid)
            if value is None:
                problems.append(f"JSON 缺指标 {iid} 或局部权重为空")
                continue
            if abs(recomputed["locals"][c["id"]][index] - float(value) / local_sums[c["id"]]) > tol * ITEM_TOL_FACTOR:
                problems.append(f"指标 {iid} 局部权重回读不一致")
    if abs(sum(recomputed["criteria"]) - 1.0) > tol * 4:
        problems.append("准则权重和 ≠ 1")
    # 回算综合权重（准则×局部）与 JSON globalWeight 对账（E 列为公式，openpyxl 无缓存值）
    recomputed_globals = [recomputed["criteria"][CRITERIA_ORDER.index(cid)] * w for cid in CRITERIA_ORDER for w in recomputed["locals"][cid]]
    expected_globals = [expected_criteria.get(cid, 0.0) * (w / (local_sums.get(cid) or 1.0))
                        for cid in CRITERIA_ORDER for w in [local_raw.get(i["id"], 0.0)
                        for i in catalog["indicators"] if i["criterionId"] == cid]]
    if len(expected_globals) == 16:
        for idx in range(16):
            if abs(recomputed_globals[idx] - expected_globals[idx]) > tol * ITEM_TOL_FACTOR:
                problems.append(f"第 {idx + 1} 项综合权重与 xlsx 回算不一致")
                break
        # JSON globalWeight 也须自洽（防篡改；按份额归一后比，同样留量化余量）
        globals_json = [i.get("globalWeight") for i in indicators_json]
        if all(g is not None for g in globals_json):
            sum_global = sum(float(g) for g in globals_json) or 1.0
            for idx, value in enumerate(globals_json):
                if abs(float(value) / sum_global - expected_globals[idx]) > tol * ITEM_TOL_FACTOR:
                    problems.append(f"第 {idx + 1} 项 JSON globalWeight 不自洽")
                    break
    # scores 若 JSON 带结果，回算需一致
    saved = payload.get("scores") or {}
    for key in ("criteriaResults", "indicatorResults"):
        if saved.get(key):
            if not all(abs(float(a) - float(b)) <= tol
                       for a, b in zip(saved[key], recomputed["scores"][key])):
                problems.append(f"scores.{key} 与 xlsx 回算不一致")
    if args.template and payload.get("display", {}).get("templateSha256") and not template_ok:
        problems.append("模板原件 sha256 与 display.templateSha256 不一致（模板被改？）")
    result = {"ok": not problems, "problems": problems,
              "recomputedCriteria": recomputed["criteria"],
              "templateUnchanged": bool(template_ok) if args.template else None}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not problems else 1
