#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ahp_aggregate.py -- cardo-ahp 多专家聚合（逐元素几何平均）+ 专家一致性/离散度。

口径（规则 aggregation 段）：defaultRule=geometric-elementwise（逐上三角元素几何平均、保持互反）；
compareRule=weight-then-average（--compare 另出对比，不替代主口径）；expertEqualWeight=true；
minExperts=5/maxExperts=7（低于 min 仍聚合但写 gap+警告）；dispersion 超阈值只建议复评，不自动剔除离群专家。

用法：python .../ahp_aggregate.py aggregate --dir "<问卷目录>" [--out <json>] [--format json|md] [--compare] [--no-timestamp]
退出码：0 成功且聚合 CR 通过；1 无有效专家或聚合 CR≥0.1；2 输入错；3 专家数不足（需人工确认）。
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402
import _survey_lib as survey  # noqa: E402

TZ = timezone(timedelta(hours=8))
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"
SCHEMA = "cardo/ahp-aggregated/v1"


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def now_iso() -> str: return datetime.now(TZ).isoformat(timespec="seconds")


def load_manifest(directory: Path) -> dict | None:
    path = directory / "问卷清单.json"
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None


def collect_experts(directory: Path) -> tuple[list[dict], list[dict]]:
    """读所有专家问卷；返回 (有效专家, 问题清单)。"""
    manifest = load_manifest(directory)
    entries = []
    if manifest:
        for item in manifest.get("experts", []):
            entries.append({"id": item.get("id"), "file": item.get("file"),
                            "indicatorScale": item.get("indicatorScale", "scale_0_1"),
                            "criteriaScale": item.get("criteriaScale", "scale_1_9")})
    else:
        for path in sorted(directory.glob("AHP专家问卷-专家*.xlsx")):
            entries.append({"id": path.stem.split("专家")[-1], "file": path.name,
                            "indicatorScale": "scale_0_1", "criteriaScale": "scale_1_9"})
    if not entries:
        raise lib.AhpError("AHP-E-INPUT", f"目录下没有可聚合的问卷：{directory}")
    experts: list[dict] = []
    problems: list[dict] = []
    expert_warnings: list[dict] = []
    for entry in entries:
        path = directory / str(entry["file"])
        if not path.is_file():
            problems.append({"code": "AHP-GAP-EXPERT-INPUT", "expertId": entry["id"],
                             "what": f"问卷不存在：{entry['file']}", "requiredAction": "重发问卷或确认清单与实际文件一致", "owner": "用户"})
            continue
        try:
            parsed = survey.read_questionnaire(path, entry["indicatorScale"], entry["criteriaScale"])
        except lib.AhpError as exc:
            problems.append({"code": "AHP-GAP-EXPERT-INPUT", "expertId": entry["id"],
                             "what": f"问卷不可读：{exc.message}", "requiredAction": "重发问卷或请专家重新填写", "owner": "用户"})
            continue
        except Exception as exc:  # 坏 xlsx / zip / 权限等
            problems.append({"code": "AHP-GAP-EXPERT-INPUT", "expertId": entry["id"],
                             "what": f"问卷无法打开（{type(exc).__name__}）", "requiredAction": "确认文件是有效 xlsx 后重发", "owner": "用户"})
            continue
        if parsed["problems"]:
            problems.append({"code": "AHP-GAP-EXPERT-INPUT", "expertId": entry["id"],
                             "what": f"结构问题 {len(parsed['problems'])} 项（首项："
                                     f"{parsed['problems'][0]['code']} {parsed['problems'][0].get('cell', '')}）",
                             "requiredAction": "请专家补齐/修正后重新回收", "owner": "用户"})
            continue
        for warning in parsed.get("warnings", []):
            warning = dict(warning)
            warning["expertId"] = entry["id"]
            expert_warnings.append(warning)
        experts.append({"id": entry["id"], "matrices": parsed["matrices"]})
    return experts, problems, expert_warnings


def geometric_mean(values: list[float]) -> float: return math.exp(sum(math.log(v) for v in values) / len(values))


def aggregate_matrix(expert_matrices: list[list[list[float]]], order: int) -> list[list[float]]:
    matrix = [[1.0] * order for _ in range(order)]
    for i in range(order):
        for j in range(i + 1, order):
            value = geometric_mean([m[i][j] for m in expert_matrices])
            matrix[i][j] = value
            matrix[j][i] = 1.0 / value
    return matrix


def avg_weights(vector_list: list[list[float]]) -> list[float]:
    return [sum(v[i] for v in vector_list) / len(vector_list) for i in range(len(vector_list[0]))]


def dispersion(values: list[float]) -> dict:
    mean = sum(values) / len(values)
    return {
        "range": max(values) - min(values),
        "mean_abs_deviation": sum(abs(v - mean) for v in values) / len(values),
        "mean": mean,
        "min": min(values),
        "max": max(values),
    }


def matrix_report(matrix: list[list[float]], scale_id: str, meta: dict) -> dict:
    evaluation = lib.evaluate(matrix)
    consistency = evaluation["consistency"]
    return {
        "id": meta["id"], "name": meta["name"], "scale": scale_id,
        "n": len(matrix), "matrix": matrix,
        "weights": evaluation["weightsGeometric"],
        "weightsEigen": evaluation["weightsEigen"],
        "lambdaMax": evaluation["lambdaMax"],
        "ci": consistency["ci"], "ri": consistency["ri"], "cr": consistency["cr"],
        "pass": consistency["pass"],
    }

def cmd_aggregate(args) -> int:
    directory = Path(args.dir)
    if not directory.is_dir():
        raise lib.AhpError("AHP-E-INPUT", f"目录不存在：{args.dir}")
    experts, problems, expert_warnings = collect_experts(directory)
    if not experts:
        detail = "；".join(f"{p.get('expertId')}:{p.get('what')}" for p in problems[:5])
        raise lib.AhpError("AHP-E-AGG",
                           f"没有可用专家问卷（全部缺失或结构不合规）{('；' + detail) if detail else ''}", 1)
    cat = lib.indicator_catalog()
    criterion_ids = [c["id"] for c in cat["criteria"]]
    rules = lib.load_scale_rules()
    aggregation = rules.get("aggregation") or {}
    min_experts = int(aggregation.get("minExperts", 5))
    max_experts = int(aggregation.get("maxExperts", 7))
    review = (aggregation.get("dispersion") or {}).get("reviewThreshold", {})
    range_limit = float(review.get("range", 0.2))
    mad_limit = float(review.get("meanAbsDeviation", 0.1))
    scale_of = {c["id"]: "scale_0_1" for c in cat["criteria"]}
    scale_of["criteria"] = "scale_1_9"
    meta = {"criteria": {"id": "criteria", "name": "准则层"}}
    for c in cat["criteria"]:
        meta[c["id"]] = {"id": c["id"], "name": c["name"]}
    matrix_ids = ["criteria"] + criterion_ids
    missing = [mid for mid in matrix_ids if any(mid not in e["matrices"] for e in experts)]
    if missing:
        raise lib.AhpError("AHP-E-AGG", f"专家问卷缺矩阵：{missing}（不补数）", 1)

    matrices = {mid: aggregate_matrix([e["matrices"][mid]["matrix"] for e in experts],
                                      len(experts[0]["matrices"][mid]["matrix"]))
                for mid in matrix_ids}
    reports = {mid: matrix_report(matrices[mid], scale_of[mid], meta[mid]) for mid in matrix_ids}

    # 每位专家权重/CR（单独复算）
    expert_reports = []
    for expert in experts:
        item = {"id": expert["id"], "matrices": {}}
        for mid in matrix_ids:
            evaluation = lib.evaluate(expert["matrices"][mid]["matrix"])
            item["matrices"][mid] = {
                "weights": evaluation["weightsGeometric"],
                "lambdaMax": evaluation["lambdaMax"],
                "cr": evaluation["consistency"]["cr"],
                "pass": evaluation["consistency"]["pass"],
            }
        expert_reports.append(item)

    # 离散度：准则权重 + 16 项局部权重（逐专家）
    dispersion_report = {"criteria": {}, "indicators": {}, "reviewThreshold": {
        "range": range_limit, "meanAbsDeviation": mad_limit}, "reviewSuggestions": []}
    crit_values = {c["id"]: [] for c in cat["criteria"]}
    local_values: dict[str, list[float]] = {i["id"]: [] for i in cat["indicators"]}
    for expert in expert_reports:
        for idx, c in enumerate(cat["criteria"]):
            crit_values[c["id"]].append(expert["matrices"]["criteria"]["weights"][idx])
            for j, ind in enumerate([i for i in cat["indicators"] if i["criterionId"] == c["id"]]):
                local_values[ind["id"]].append(expert["matrices"][c["id"]]["weights"][j])
    for c in cat["criteria"]:
        stats = dispersion(crit_values[c["id"]])
        dispersion_report["criteria"][c["id"]] = stats
        if stats["range"] > range_limit or stats["mean_abs_deviation"] > mad_limit:
            dispersion_report["reviewSuggestions"].append({
                "target": f"criteria.{c['id']}", "range": stats["range"],
                "meanAbsDeviation": stats["mean_abs_deviation"],
                "suggestion": "建议复评/增加专家（不自动剔除）",
            })
    for ind in cat["indicators"]:
        stats = dispersion(local_values[ind["id"]])
        dispersion_report["indicators"][ind["id"]] = stats
        if stats["range"] > range_limit or stats["mean_abs_deviation"] > mad_limit:
            dispersion_report["reviewSuggestions"].append({
                "target": f"indicator.{ind['id']}", "range": stats["range"],
                "meanAbsDeviation": stats["mean_abs_deviation"],
                "suggestion": "建议复评/增加专家（不自动剔除）",
            })

    # 综合权重：聚合准则权重 × 聚合局部权重
    criteria_weights = reports["criteria"]["weights"]
    indicators_out = []
    for c in cat["criteria"]:
        local = reports[c["id"]]["weights"]
        for j, ind in enumerate([i for i in cat["indicators"] if i["criterionId"] == c["id"]]):
            indicators_out.append({
                "id": ind["id"], "criterionId": c["id"], "name": ind["name"],
                "direction": ind["direction"], "source": ind["source"], "required": True,
                "localWeight": local[j],
                "globalWeight": criteria_weights[criterion_ids.index(c["id"])] * local[j],
            })
    criteria_out = [{"id": c["id"], "name": c["name"], "weight": criteria_weights[idx],
                     "lambdaMax": reports[c["id"]]["lambdaMax"], "cr": reports[c["id"]]["cr"],
                     "pass": reports[c["id"]]["pass"]}
                    for idx, c in enumerate(cat["criteria"])]

    # 专家一致性：与均值权重的最大绝对差（权重逐专家平均作对照口径）
    mean_weight_vectors = {}
    for mid in matrix_ids:
        mean_weight_vectors[mid] = avg_weights([e["matrices"][mid]["weights"] for e in expert_reports])
    agreement = []
    for expert in expert_reports:
        diffs = []
        for mid in matrix_ids:
            diffs.extend(abs(a - b) for a, b in zip(expert["matrices"][mid]["weights"],
                                                    mean_weight_vectors[mid]))
        agreement.append({"expertId": expert["id"], "maxAbsDiffToMeanWeight": max(diffs),
                          "meanAbsDiffToMeanWeight": sum(diffs) / len(diffs)})
    agreement.sort(key=lambda x: -x["maxAbsDiffToMeanWeight"])

    gaps = []
    warnings = []
    if len(experts) < min_experts:
        gaps.append({"code": aggregation.get("degradeGapCode", "AHP-GAP-EXPERTS"),
                     "what": f"有效专家 {len(experts)} < 最少 {min_experts}；已降级聚合",
                     "requiredAction": "补充专家问卷或记录降级原因"})
        warnings.append({"code": "AHP-W-DEGRADED-EXPERTS",
                         "what": f"专家数不足（{len(experts)}/{min_experts}），聚合结果仅作参考"})
    if len(experts) > max_experts:
        warnings.append({"code": "AHP-W-EXPERTS-OVER", "what": f"专家数 {len(experts)} > {max_experts}"})
    expert_fail = []
    for item in expert_reports:
        failed = [mid for mid in matrix_ids if not item["matrices"][mid]["pass"]]
        if failed:
            expert_fail.append({"expertId": item["id"], "matrices": failed,
                                "cr": {mid: item["matrices"][mid]["cr"] for mid in failed}})
            warnings.append({"code": "AHP-W-EXPERT-CR",
                             "what": f"专家{item['id']} 有 {len(failed)} 个矩阵 CR 未通过（不阻断聚合）",
                             "expertId": item["id"], "matrices": failed})
    warnings.extend(expert_warnings)
    gaps.extend(problems)

    aggregate_pass = all(reports[mid]["pass"] for mid in matrix_ids)
    payload = {
        "schema": SCHEMA, "schemaVersion": 1, "producer": PRODUCER,
        "source": [{"kind": "survey-dir", "id": "survey-dir",
                     "path": "<已脱敏绝对路径>/" + Path(args.dir).name
                     if Path(args.dir).is_absolute() else args.dir}],
        "gaps": gaps, "warnings": warnings,
        "ruleVersion": int(rules.get("version") or 0), "engineVersion": ENGINE_VERSION,
        "aggregation": {
            "rule": aggregation.get("defaultRule", "geometric-elementwise"),
            "compareRule": aggregation.get("compareRule", "weight-then-average"),
            "expertEqualWeight": bool(aggregation.get("expertEqualWeight", True)),
            "expertCount": len(experts), "minExperts": min_experts, "maxExperts": max_experts,
            "degraded": len(experts) < min_experts,
        },
        "experts": [{"id": e["id"], "matrices": e["matrices"]} for e in expert_reports],
        "expertAgreement": agreement,
        "matrices": reports,
        "criteria": criteria_out,
        "indicators": indicators_out,
        "weights": {
            "unit": "weight", "tolerance": lib.rules().compare_tol,
            "sumCriteria": sum(criteria_weights),
            "sumGlobal": sum(i["globalWeight"] for i in indicators_out),
        },
        "consistency": {
            "passCount": sum(1 for mid in matrix_ids if reports[mid]["pass"]),
            "total": len(matrix_ids),
            "aggregatePass": aggregate_pass,
            "expertFailCount": len(expert_fail),
            "expertFailures": expert_fail,
            "perMatrix": {mid: {"cr": reports[mid]["cr"], "pass": reports[mid]["pass"]} for mid in matrix_ids},
        },
        "dispersion": dispersion_report,
        "calibrationApplied": False,
    }
    if not getattr(args, "no_timestamp", False):
        payload["generatedAt"] = now_iso()

    if getattr(args, "compare", False):
        compare_criteria = avg_weights([e["matrices"]["criteria"]["weights"] for e in expert_reports])
        compare_indicators = []
        for c in cat["criteria"]:
            local_avg = avg_weights([e["matrices"][c["id"]]["weights"] for e in expert_reports])
            for j, ind in enumerate([i for i in cat["indicators"] if i["criterionId"] == c["id"]]):
                compare_indicators.append({"id": ind["id"],
                                           "localWeight": local_avg[j],
                                           "globalWeight": compare_criteria[criterion_ids.index(c["id"])] * local_avg[j]})
        primary_global = {i["id"]: i["globalWeight"] for i in indicators_out}
        compare_global = {i["id"]: i["globalWeight"] for i in compare_indicators}
        payload["weightsCompare"] = {
            "rule": "weight-then-average",
            "note": "对比口径：逐专家权重算术平均后再合成；仅作对比，不替代主口径几何平均。",
            "criteria": [{"id": c["id"], "name": c["name"], "weight": compare_criteria[idx]}
                         for idx, c in enumerate(cat["criteria"])],
            "indicators": compare_indicators,
            "maxAbsDeltaGlobal": max(abs(primary_global[k] - compare_global[k]) for k in primary_global),
            "sumGlobal": sum(compare_global.values()),
        }
    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.format == "md":
        print(render_md(payload))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
    if not aggregate_pass:
        return 1
    if len(experts) < min_experts:
        return 3
    return 0

def render_md(payload: dict) -> str:
    lines = ["# AHP 多专家聚合", "",
             f"- 规则：{payload['aggregation']['rule']}；专家 {payload['aggregation']['expertCount']} 位"
             f"（最少 {payload['aggregation']['minExperts']}）",
             f"- 聚合一致性：{payload['consistency']['passCount']}/{payload['consistency']['total']} 通过"
             f"；总判定 {'通过' if payload['consistency']['aggregatePass'] else '不通过'}", "",
             "## 准则权重", "", "| 准则 | 聚合权重 | 聚合 CR |", "|---|---|---|"]
    for item in payload["criteria"]:
        lines.append(f"| {item['name']} | {item['weight']:.6f} | {item['cr']:.6f} |")
    lines += ["", "## 16 项指标（局部 / 综合）", "", "| 指标 | 局部 | 综合 |", "|---|---|---|"]
    for item in payload["indicators"]:
        lines.append(f"| {item['name']} | {item['localWeight']:.6f} | {item['globalWeight']:.6f} |")
    if payload["dispersion"]["reviewSuggestions"]:
        lines += ["", "## 离散度提示（建议复评，不自动剔除）", "",
                  "| 目标 | range | meanAbsDeviation |", "|---|---|---|"]
        for s in payload["dispersion"]["reviewSuggestions"]:
            lines.append(f"| {s['target']} | {s['range']:.6f} | {s['meanAbsDeviation']:.6f} |")
    if payload["expertAgreement"]:
        lines += ["", "## 专家一致性（与均值权重最大差）", ""]
        for a in payload["expertAgreement"]:
            lines.append(f"- 专家{a['expertId']}：max={a['maxAbsDiffToMeanWeight']:.6f}，"
                         f"mean={a['meanAbsDiffToMeanWeight']:.6f}")
    if payload["gaps"]:
        lines += ["", "## Gaps", ""] + [f"- {g['code']}：{g.get('what', '')}" for g in payload["gaps"]]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 多专家聚合")
    sub = ap.add_subparsers(dest="command", required=True)
    p = sub.add_parser("aggregate", help="聚合专家问卷")
    p.add_argument("--dir", required=True)
    p.add_argument("--out", default=None)
    p.add_argument("--format", choices=["json", "md"], default="json")
    p.add_argument("--compare", action="store_true",
                   help="同时输出 weight-then-average 对比口径（不替代主口径）")
    p.add_argument("--no-timestamp", action="store_true")
    p.set_defaults(func=cmd_aggregate)
    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except lib.AhpError as exc:
        print(json.dumps({"error": {"code": exc.code, "message": exc.message}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return exc.exit_code
    except (RuntimeError, OSError) as exc:
        print(json.dumps({"error": {"code": "AHP-E-INTERNAL", "message": str(exc)}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
