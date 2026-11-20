#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ahp_calc.py -- cardo-ahp 计算引擎 CLI（标准库，零第三方依赖）。

子命令：
  indicators   打印指标目录（4 准则 × 16 指标；命名映射/方向/来源/模板坐标/锚点）
  validate     校验单个矩阵（三口径任一形态）→ 输出统一互反矩阵；坏数据只报错不修正
  weights      矩阵集 → 准则权重 / 指标局部权重 / 综合权重 + 双算法差异 + 一致性
  consistency  矩阵集 → 批量一致性检验 + 通过率

用法（仓库根）：
  python plugins/cardo-ahp/skills/cardo-ahp/scripts/ahp_calc.py indicators --format md
  python .../ahp_calc.py validate --matrix <矩阵.json> [--scale scale_1_9] [--mode column] [--order 4]
  python .../ahp_calc.py weights --matrix-set <矩阵集.json> [--format json|md] [--no-timestamp]
  python .../ahp_calc.py consistency --matrix-set <矩阵集.json> [--format md]

退出码：0 成功且无 CR 不通过；1 校验不通过（CR 超阈值）；2 输入/参数错；3 需人工确认（保留）。
输出：JSON 到 stdout（schema/版本见 docs/接口契约.md）；`--no-timestamp` 省略 generatedAt（逐字节幂等）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402

TZ = timezone(timedelta(hours=8))
SCHEMA = "cardo/ahp-weights/v1"
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def source_entry(kind: str, source_id: str, path_text: str) -> dict:
    """source[] 条目：带输入文件 sha256；绝对路径脱敏为占位符（契约路径纪律）。"""
    path = Path(path_text)
    digest = ""
    if path.is_file():
        digest = hashlib.sha256(path.read_bytes()).hexdigest().upper()
    shown = path_text
    if path.is_absolute():
        shown = f"<已脱敏绝对路径>/{path.name}"
    entry = {"kind": kind, "id": source_id, "path": shown}
    if digest:
        entry["sha256"] = digest
    return entry


def read_json(path_text: str, label: str):
    path = Path(path_text)
    if not path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"{label}不存在：{path_text}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise lib.AhpError("AHP-E-INPUT", f"{label}不是合法 JSON：{exc}") from exc


def envelope(args, source: list[dict], gaps: list[dict] | None = None) -> dict:
    payload = {"schema": SCHEMA, "schemaVersion": 1}
    if not args.no_timestamp:
        payload["generatedAt"] = datetime.now(TZ).isoformat(timespec="seconds")
    payload["producer"] = PRODUCER
    payload["source"] = source
    payload["gaps"] = gaps or []
    payload["ruleVersion"] = _rule_version()
    payload["engineVersion"] = ENGINE_VERSION
    return payload


def _rule_version() -> int:
    return int((lib.load_scale_rules().get("version") or 0))


def fmt(value: float, digits: int = 6) -> str:
    return f"{value:.{digits}f}"


def rank_map(weights: list[float]) -> list[int]:
    order = sorted(range(len(weights)), key=lambda i: (-weights[i], i))
    ranks = [0] * len(weights)
    for position, index in enumerate(order, start=1):
        ranks[index] = position
    return ranks


# ---------------------------------------------------------------- indicators
def cmd_indicators(args) -> int:
    catalog = lib.indicator_catalog()
    if args.format == "json":
        print(json.dumps(catalog, ensure_ascii=False, indent=2))
        return 0
    print(f"# 指标目录（{catalog['schema']} v{catalog['version']}）")
    print(f"目标层：{catalog['goal'].get('name')}")
    print("\n## 准则层")
    print("| id | 名称 | 模板写法 | 任务书写法 |")
    print("|---|---|---|---|")
    for c in catalog["criteria"]:
        display = c.get("displayName") or {}
        print(f"| {c['id']} | {c['name']} | {display.get('template', '')} | {display.get('taskbook', '')} |")
    print("\n## 指标层（16 项）")
    print("| id | 准则 | 名称 | 方向 | 来源 | 模板格 |")
    print("|---|---|---|---|---|---|")
    for ind in catalog["indicators"]:
        print(f"| {ind['id']} | {ind['criterionId']} | {ind['name']} | {ind['direction']} | "
              f"{ind['source']} | {ind.get('templateCell') or '—'} |")
    return 0


# ---------------------------------------------------------------- validate
def cmd_validate(args) -> int:
    payload = read_json(args.matrix, "矩阵文件")
    if isinstance(payload, dict) and isinstance(payload.get("matrices"), list):
        if len(payload["matrices"]) != 1:
            raise lib.AhpError("AHP-E-INPUT", f"--matrix 只接受单矩阵，实际 {len(payload['matrices'])} 个；多矩阵用 --matrix-set")
        payload = payload["matrices"][0]
    if not isinstance(payload, dict):
        raise lib.AhpError("AHP-E-INPUT", "矩阵文件应为 {scale,mode,order,values}")
    order = args.order or payload.get("order") or 0
    scale_id = args.scale or payload.get("scale")
    mode = args.mode or payload.get("mode") or "column"
    if not scale_id:
        raise lib.AhpError("AHP-E-SCALE", "缺标度：用 --scale 或文件内 scale")
    if not order and mode != "column":
        raise lib.AhpError("AHP-E-ORDER", "缺阶数：用 --order 或文件内 order")
    order = order or len(payload.get("values") or [])
    matrix = lib.build_matrix(payload.get("values"), order, scale_id, mode)
    result = {
        "schema": "cardo/ahp-judgment/v1",
        "schemaVersion": 1,
        "producer": PRODUCER,
        "source": [source_entry("judgment", payload.get("id") or "matrix", args.matrix)],
        "gaps": [],
        "matrix": {"id": payload.get("id") or "matrix", "level": payload.get("level") or "unknown",
                   "scale": scale_id, "mode": mode, "order": order},
        "reciprocal": matrix,
        "validation": {"positive": True, "reciprocal": True, "diagonalOne": True, "orderInRange": True, "errors": []},
        "warnings": sorted(set(lib.WARNINGS)),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def _round(value: float) -> float:
    return round(value, 12)


# ---------------------------------------------------------------- matrix set
def evaluate_set(args) -> tuple[list[dict], list[dict], dict]:
    payload = read_json(args.matrix_set, "矩阵集文件")
    matrices = lib.load_matrix_set(payload)
    warnings = sorted(set(lib.WARNINGS))
    summary = {
        "passCount": sum(1 for m in matrices if m["consistency"]["pass"]),
        "total": len(matrices),
        "passRate": (sum(1 for m in matrices if m["consistency"]["pass"]) / len(matrices)) if matrices else 0.0,
        "threshold": lib.rules().cr_threshold,
        "comparator": "<",
        "tolerance": lib.rules().cr_tolerance,
    }
    return matrices, warnings, summary


def matrix_summary(m: dict) -> dict:
    cons = m["consistency"]
    return {
        "id": m["id"], "level": m["level"], "scale": m["scale"], "mode": m["mode"], "n": m["n"], "order": m["n"],
        "lambdaMax": m["lambdaMax"], "ci": cons["ci"], "ciRaw": cons["ciRaw"], "ri": cons["ri"],
        "cr": cons["cr"], "pass": cons["pass"], "truncatedCi": cons["truncatedCi"],
        "eigenIterations": m["eigenIterations"],
        "weightsGeometric": m["weightsGeometric"],
        "weightsEigen": m["weightsEigen"],
        "maxAbsDelta": _round(m["delta"]["maxAbs"]),
    }


def cmd_consistency(args) -> int:
    matrices, warnings, summary = evaluate_set(args)
    payload = envelope(args, [source_entry("matrix-set", "matrix-set", args.matrix_set)])
    payload["schema"] = "cardo/ahp-consistency/v1"
    payload["matrices"] = [matrix_summary(m) for m in matrices]
    payload["consistency"] = summary
    payload["warnings"] = warnings
    _emit(args, payload, md=render_consistency_md)
    return 0 if summary["passCount"] == summary["total"] else 1


def cmd_weights(args) -> int:
    matrices, warnings, summary = evaluate_set(args)
    by_id = {m["id"]: m for m in matrices}
    catalog = lib.indicator_catalog()
    criterion_ids = [c["id"] for c in catalog["criteria"]]
    criteria_matrix = by_id.get("criteria")
    if criteria_matrix is None:
        raise lib.AhpError("AHP-E-MISSING", "矩阵集缺 criteria 矩阵")
    if criteria_matrix["n"] != len(criterion_ids):
        raise lib.AhpError(
            "AHP-E-ORDER",
            f"criteria 矩阵阶数 {criteria_matrix['n']} ≠ 准则数 {len(criterion_ids)}（{criterion_ids}）",
        )
    criteria_weights = criteria_matrix["weightsGeometric"]
    eigen_criteria = criteria_matrix["weightsEigen"]
    criteria_ranks = rank_map(criteria_weights)

    local_by_criterion: dict[str, list[float]] = {}
    eigen_local: dict[str, list[float]] = {}
    for cid in criterion_ids:
        m = by_id.get(cid)
        if m is None:
            raise lib.AhpError("AHP-E-MISSING", f"矩阵集缺指标层矩阵 {cid}")
        count = len([i for i in catalog["indicators"] if i["criterionId"] == cid])
        if m["n"] != count:
            raise lib.AhpError("AHP-E-ORDER", f"{cid} 矩阵阶数 {m['n']} ≠ 指标数 {count}")
        local_by_criterion[cid] = m["weightsGeometric"]
        eigen_local[cid] = m["weightsEigen"]

    global_weights = lib.weighted_global(criteria_weights, local_by_criterion, criterion_ids)
    if abs(sum(global_weights) - 1.0) > lib.rules().compare_tol:
        raise lib.AhpError("AHP-E-SCALE", f"综合权重和 {sum(global_weights)!r} ≠ 1（1e-12）", 1)
    global_ranks = rank_map(global_weights)

    criteria = []
    for idx, c in enumerate(catalog["criteria"]):
        criteria.append({
            "id": c["id"], "name": c["name"], "displayName": c.get("displayName") or {},
            "weight": _round(criteria_weights[idx]), "geometric": _round(criteria_weights[idx]),
            "eigen": _round(eigen_criteria[idx]), "rank": criteria_ranks[idx],
        })
    indicators = []
    position = 0
    local_rank_within: dict[str, int] = {}
    for cid in criterion_ids:
        local = local_by_criterion[cid]
        ranks = rank_map(local)
        for i, ind in enumerate([x for x in catalog["indicators"] if x["criterionId"] == cid]):
            indicators.append({
                "id": ind["id"], "criterionId": cid, "name": ind["name"],
                "direction": ind["direction"], "source": ind["source"],
                "required": True,
                "localWeight": _round(local[i]), "globalWeight": _round(global_weights[position]),
                "geometric": _round(local[i]), "eigen": _round(eigen_local[cid][i]),
                "rank": global_ranks[position], "localRank": ranks[i],
                "templateCell": ind.get("templateCell"),
            })
            position += 1
        local_rank_within[cid] = 1

    matrices_out = [matrix_summary(m) for m in matrices]
    payload = envelope(args, [source_entry("matrix-set", "matrix-set", args.matrix_set)])
    payload.update({
        "indicatorSetVersion": int(catalog.get("version") or 0),
        "criteria": criteria,
        "indicators": indicators,
        "weights": {
            "unit": "fraction", "tolerance": lib.rules().compare_tol,
            "sumCriteria": _round(sum(criteria_weights)),
            "sumLocalPerCriterion": {cid: _round(sum(local_by_criterion[cid])) for cid in criterion_ids},
            "sumGlobal": _round(sum(global_weights)),
        },
        "methods": {
            "primary": "geometric", "compare": "eigen",
            "delta": [{"id": m["id"], "maxAbsDelta": _round(m["delta"]["maxAbs"])} for m in matrices],
        },
        "consistency": {**summary, "matrices": matrices_out},
        "warnings": warnings,
        "demo": bool(getattr(args, "demo", False)),
    })
    _emit(args, payload, md=render_weights_md)
    return 0 if summary["passCount"] == summary["total"] else 1


# ---------------------------------------------------------------- 输出
def _emit(args, payload: dict, md) -> None:
    if args.format == "md":
        print(md(payload))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))


def render_consistency_md(payload: dict) -> str:
    lines = ["# AHP 一致性检验（批量）", "",
             f"- 通过 {payload['consistency']['passCount']}/{payload['consistency']['total']}，"
             f"通过率 {payload['consistency']['passRate']:.4f}",
             f"- 判定：CR {payload['consistency']['comparator']} {payload['consistency']['threshold']}（容差 {payload['consistency']['tolerance']}）", "",
             "| 矩阵 | n | λmax | CI | RI | CR | 判定 |", "|---|---|---|---|---|---|---|"]
    for m in payload["matrices"]:
        lines.append(f"| {m['id']} | {m['order']} | {fmt(m['lambdaMax'])} | {fmt(m['ci'])} | {fmt(m['ri'], 2)} | "
                     f"{fmt(m['cr'])} | {'通过' if m['pass'] else '不通过'} |")
    return "\n".join(lines)


def render_weights_md(payload: dict) -> str:
    lines = ["# AHP 权重结果", "",
             f"- 准则权重和 {payload['weights']['sumCriteria']}；综合权重和 {payload['weights']['sumGlobal']}",
             f"- 一致性 {payload['consistency']['passCount']}/{payload['consistency']['total']} 通过", "",
             "## 准则层", "| 准则 | 权重 | 排名 |", "|---|---|---|"]
    for c in payload["criteria"]:
        lines.append(f"| {c['name']} | {fmt(c['weight'])} | {c['rank']} |")
    lines += ["", "## 指标层（综合权重降序）", "| 排名 | 指标 | 准则 | 局部权重 | 综合权重 |", "|---|---|---|---|---|"]
    for ind in sorted(payload["indicators"], key=lambda x: x["rank"]):
        lines.append(f"| {ind['rank']} | {ind['name']} | {ind['criterionId']} | {fmt(ind['localWeight'])} | {fmt(ind['globalWeight'])} |")
    return "\n".join(lines)


# ---------------------------------------------------------------- main
def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 计算引擎（几何平均 + 特征值 + 一致性）")
    sub = ap.add_subparsers(dest="command", required=True)

    p_ind = sub.add_parser("indicators", help="打印指标目录")
    p_ind.add_argument("--format", choices=["json", "md"], default="md")
    p_ind.set_defaults(func=cmd_indicators)

    p_val = sub.add_parser("validate", help="校验单个矩阵并输出统一互反矩阵")
    p_val.add_argument("--matrix", required=True)
    p_val.add_argument("--scale", default=None)
    p_val.add_argument("--mode", default=None)
    p_val.add_argument("--order", type=int, default=0)
    p_val.set_defaults(func=cmd_validate)

    for name, func, help_text in (("weights", cmd_weights, "权重（双算法 + 一致性）"),
                                  ("consistency", cmd_consistency, "批量一致性检验")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--matrix-set", required=True)
        p.add_argument("--format", choices=["json", "md"], default="json")
        p.add_argument("--no-timestamp", action="store_true", help="省略 generatedAt（逐字节幂等）")
        p.add_argument("--demo", action="store_true",
                       help="演示样例：权重表 JSON 写 demo=true（不改任何数值）")
        p.set_defaults(func=func)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except lib.AhpError as exc:
        print(json.dumps({"error": {"code": exc.code, "message": exc.message, "detail": exc.detail}},
                         ensure_ascii=False, indent=2), file=sys.stderr)
        return exc.exit_code
    except (RuntimeError, OSError) as exc:
        print(json.dumps({"error": {"code": "AHP-E-INTERNAL", "message": str(exc)}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())