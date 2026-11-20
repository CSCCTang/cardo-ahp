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
    }
    if not getattr(args, "no_timestamp", False):
        payload["generatedAt"] = now_iso()
    return payload
def write_workbook(template_path: Path, out_path: Path, data: dict, scores: dict | None) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(template_path.read_bytes())
    wb = load_workbook(out_path)
    ws = wb["层次分析模板"]
    criteria_values = [data["criteria"][cid] for cid in data["criteriaOrder"]]
    # 模板第 1 列是矩阵 a_i1，首格 a_11 必须为 1（模板 λmax=Σw·列和 才等于 n）；
    # 故列向量写 w_i / w_1（不是 w_i/min），回读归一仍得原权重。
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
    if scores:
        criteria_scores = scores.get("criteriaScores")
        indicator_scores = scores.get("indicatorScores")
        if criteria_scores is not None and (len(criteria_scores) != 3 or any(len(r) != 4 for r in criteria_scores)):
            raise lib.AhpError("AHP-E-INPUT", "scores.criteriaScores 必须是 3 行 × 4 列")
        if indicator_scores is not None and (len(indicator_scores) != 3 or any(len(r) != 16 for r in indicator_scores)):
            raise lib.AhpError("AHP-E-INPUT", "scores.indicatorScores 必须是 3 行 × 16 列")
        for r, row in enumerate(criteria_scores or [], start=30):
            for c, value in enumerate(row, start=2):
                ws.cell(row=r, column=c).value = value
        for r, row in enumerate(indicator_scores or [], start=95):
            for c, value in enumerate(row, start=2):
                ws.cell(row=r, column=c).value = value
    if "导出值" not in wb.sheetnames:
        ws_out = wb.create_sheet("导出值")
        ws_out.append(["项", "值"])
        for item in data["criteriaOrder"]:
            ws_out.append([f"准则权重.{item}", data["criteria"][item]])
        for item in data["indicatorOrder"]:
            ws_out.append([f"指标局部.{item}", data["indicators"][item][0]])
            ws_out.append([f"指标综合.{item}", data["indicators"][item][1]])
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
    payload = build_full_payload(args, data, Path(args.weights), template_path, out_path, out_json)
    write_workbook(template_path, out_path, data, scores)
    if sha256_file(template_path) != template_sha:
        raise lib.AhpError("AHP-E-INPUT", "模板原件在导出过程中被改动（应只读）", 1)
    if scores:
        recomputed = recompute_from_workbook(out_path)
        payload["scores"] = {"criteriaScores": scores.get("criteriaScores"),
                             "indicatorScores": scores.get("indicatorScores"),
                             "criteriaResults": recomputed["scores"]["criteriaResults"],
                             "indicatorResults": recomputed["scores"]["indicatorResults"]}
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    result = {"xlsx": str(out_path), "json": str(out_json), "schema": SCHEMA,
              "templateSha256": template_sha, "sumGlobal": payload["weights"]["sumGlobal"],
              "templateUnchanged": True}
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
    criteria_vec = [float(ws[c].value) for c in CRITERIA_CELLS]
    criteria = row_weights(criteria_vec)
    locals_map = {}
    for cid, top in INDICATOR_BLOCKS.items():
        locals_map[cid] = row_weights([float(ws[f"B{top + k}"].value) for k in range(4)])
    crit_rows = [[ws.cell(row=r, column=c).value for c in range(2, 6)] for r in range(30, 33)]
    ind_rows = [[ws.cell(row=r, column=c).value for c in range(2, 18)] for r in range(95, 98)]
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
