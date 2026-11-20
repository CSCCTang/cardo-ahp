#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""score_project.py -- cardo-ahp 项目打分与综合评分排序（cardo/ahp-scores/v1）。

口径（规则 scoring / ratings / quantification 段）：
  · 10 分制（0-10，1 位小数）；打分依据优先人工（manualScores），其次按量化分档 step-scoring（不内插）；
  · 分档区间 [a,b)：`a-b`、`>=a`、`<a`；端点同时命中两档时取分数更高的一档（higher_band_wins）；
  · 综合得分 composite = Σ(score×globalWeight)；有缺项时除以已覆盖权重并标 partial=true（**不补 0**）；
  · required 缺失 → 该方案不出分，exit 3 列出缺项；optional 缺失 → 可出分，标 partial/coverage；
  · 评级左闭右开：9-10 优秀 / 7-9 良好 / 6-7 一般 / 0-6 不可行；
  · 多方案按 composite 降序（同分保持方案顺序）。

用法（仓库根）：
  python .../score_project.py score --weights "<权重.json>" --scores "<打分.json>" [--out "<评分.json>"]
      [--direction <方向ID>] [--format json|md] [--no-timestamp]
退出码：0 成功；2 输入/权重/打分错误；3 有方案 required 缺失（不出分、需补依据）。
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
import _export_lib as export_lib  # noqa: E402
import _yaml_lite as yaml_lite  # noqa: E402

TZ = timezone(timedelta(hours=8))
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"
SCHEMA = "cardo/ahp-scores/v1"
RULES_DIR = HERE.parent / "规则"
QUANT_FILE = RULES_DIR / "指标量化分档.yaml"
DIRECTION_FILE = RULES_DIR / "方向适配.yaml"


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def now_iso() -> str: return datetime.now(TZ).isoformat(timespec="seconds")


def read_json(path_text: str) -> tuple[dict, Path]:
    path = Path(path_text)
    if not path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"文件不存在：{path_text}")
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise lib.AhpError("AHP-E-INPUT", f"不是合法 JSON：{path_text}（{exc}）") from exc
    if not isinstance(payload, dict):
        raise lib.AhpError("AHP-E-INPUT", "输入 JSON 应为对象")
    return payload, path


def parse_range(text: str) -> tuple[float, float] | None:
    raw = str(text).strip()
    try:
        if raw.startswith(">="):
            return float(raw[2:]), math.inf
        if raw.startswith("<="):
            return -math.inf, float(raw[2:])
        if raw.startswith(">"):
            return float(raw[1:]) + 1e-12, math.inf
        if raw.startswith("<"):
            return -math.inf, float(raw[1:]) - 1e-12
        if "-" in raw:
            left, right = raw.split("-", 1)
            return float(left), float(right)
    except ValueError:
        return None
    return None


def load_bands() -> dict[str, dict]:
    data = yaml_lite.load_file(QUANT_FILE)
    return {item.get("id"): item for item in data.get("items", []) if item.get("id")}


def load_direction(direction_id: str) -> dict | None:
    data = yaml_lite.load_file(DIRECTION_FILE)
    for item in data.get("directions") or []:
        if item.get("id") == direction_id:
            return item
    return None


def direction_statuses(direction: dict | None) -> dict[str, str]:
    if not direction:
        return {}
    statuses = {}
    for entry in direction.get("applicability") or []:
        if entry.get("id"):
            statuses[entry["id"]] = str(entry.get("status") or "required")
    return statuses


def score_from_metric(metric_value: float, item: dict) -> tuple[float | None, dict]:
    matched = []
    for band in item.get("bands") or []:
        if "range" not in band:
            continue
        span = parse_range(band.get("range"))
        if span is None:
            continue
        low, high = span
        if low <= metric_value < high:
            matched.append(band)
    if not matched:
        return None, {"reason": "no_band_matches", "bands": [b.get("range") for b in item.get("bands") or []]}
    chosen = max(matched, key=lambda b: float(b.get("score", 0)))
    return float(chosen["score"]), {"band": chosen.get("range"), "matched": [b.get("range") for b in matched]}


def valid_score(value) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise lib.AhpError("AHP-E-SCORE", f"打分必须是数值（0-10），实际 {value!r}")
    score = float(value)
    if not math.isfinite(score):
        raise lib.AhpError("AHP-E-SCORE", f"打分必须是有限数值，实际 {value!r}")
    if score < 0 or score > 10:
        raise lib.AhpError("AHP-E-SCORE", f"打分越界：{score}（应在 0-10）")
    if abs(score * 10 - round(score * 10)) > 1e-9:
        raise lib.AhpError("AHP-E-SCORE", f"打分精度应为 1 位小数：{score}")
    return score


def rating_of(composite: float) -> dict:
    table = [
        (9.0, 10.0, "优秀", "优先投资、全力推进"),
        (7.0, 9.0, "良好", "可行投资、分期推进"),
        (6.0, 7.0, "一般", "审慎投资、优化方案"),
        (0.0, 6.0, "不可行", "暂缓/放弃投资"),
    ]
    epsilon = 1e-9  # 仅比较用：Σ(score×w) 的 1e-15 级舍入不应改变边界评级
    for low, high, label, action in table:
        if low - epsilon <= composite < high - epsilon or (label == "优秀" and composite >= 9.0 - epsilon):
            return {"label": label, "action": action, "range": [low, high]}
    return {"label": "不可行", "action": "暂缓/放弃投资", "range": [0.0, 6.0]}


def build_scheme(name: str, scheme: dict, weights: dict, bands: dict,
                 statuses: dict[str, str]) -> dict:
    indicators = weights["catalog"]["indicators"]
    manual = scheme.get("manualScores") or {}
    metrics = scheme.get("metrics") or {}
    if not isinstance(manual, dict) or not isinstance(metrics, dict):
        raise lib.AhpError("AHP-E-INPUT", f"方案 {name} 的 manualScores/metrics 必须是对象")
    scores = []
    missing_required = []
    missing_optional = []
    covered_weight = 0.0
    total_weight = 0.0
    for ind in indicators:
        iid = ind["id"]
        status = statuses.get(iid, "required")
        local = weights["indicators"][iid][0]
        global_weight = weights["indicators"][iid][1] if weights["indicators"][iid][1] is not None \
            else weights["criteria"][ind["criterionId"]] * local
        if status == "not_applicable":
            continue
        total_weight += global_weight
        entry = {"id": iid, "name": ind["name"], "criterionId": ind["criterionId"],
                 "globalWeight": global_weight, "status": status}
        value = None
        basis = None
        source = None
        declared = manual.get(iid)
        if declared is not None:
            if isinstance(declared, dict):
                value = valid_score(declared.get("score"))
                basis = declared.get("basis")
            else:
                value = valid_score(declared)
                basis = "人工打分（未注明依据）"
            source = "manual"
        elif iid in metrics:
            item = bands.get(iid) or {}
            raw_metric = metrics.get(iid)
            if isinstance(raw_metric, bool) or not isinstance(raw_metric, (int, float)) \
                    or not math.isfinite(float(raw_metric)):
                raise lib.AhpError("AHP-E-INPUT",
                                   f"指标 {iid} 的 metric 值必须是有限数值（实际 {raw_metric!r}）")
            metric_value = float(raw_metric)
            mapped, detail = score_from_metric(metric_value, item)
            if mapped is None:
                entry.update({"missing": "无法按分档映射", "metric": metric_value, "detail": detail})
                (missing_required if status == "required" else missing_optional).append(entry)
                scores.append({**{k: entry[k] for k in ("id", "name", "criterionId", "globalWeight", "status")},
                               "score": None, "source": None, "basis": None})
                continue
            value = mapped
            basis = f"metric={metric_value} → 档 {detail.get('band')}"
            source = "metric"
        if value is None:
            entry.update({"missing": "缺人工打分与可用 metric"})
            (missing_required if status == "required" else missing_optional).append(entry)
            scores.append({**{k: entry[k] for k in ("id", "name", "criterionId", "globalWeight", "status")},
                           "score": None, "source": None, "basis": None})
            continue
        covered_weight += global_weight
        entry.update({"score": value, "source": source, "basis": basis})
        scores.append(entry)
    result = {"id": scheme.get("id") or "", "name": name, "scores": scores,
              "missingRequired": missing_required, "missingOptional": missing_optional,
              "coverage": {"scored": sum(1 for s in scores if s.get("score") is not None),
                           "required": sum(1 for s in scores if s.get("status") == "required"),
                           "optional": sum(1 for s in scores if s.get("status") == "optional"),
                           "weightCovered": covered_weight, "weightTotal": total_weight}}
    if missing_required:
        result.update({"composite": None, "rating": None, "partial": False,
                       "reason": "required 指标缺依据：不出分（不补 0）"})
        return result
    if total_weight <= 0:
        raise lib.AhpError("AHP-E-SCORE", f"方案 {name} 无可用指标权重")
    composite = sum(float(s["score"]) * s["globalWeight"] for s in scores if s.get("score") is not None)
    partial = bool(missing_optional) or covered_weight < total_weight - 1e-12
    if covered_weight > 0:
        composite = composite / covered_weight
    result.update({"composite": composite, "rating": rating_of(composite), "partial": partial})
    return result


def conclusion_of(schemes: list[dict], weights: dict) -> dict:
    scored = [s for s in schemes if s.get("composite") is not None]
    if not scored:
        return {"summary": "没有方案可出分：先补 required 指标依据。", "topDrivers": [], "deductions": []}
    ranking = sorted(scored, key=lambda s: (-s["composite"], schemes.index(s)))
    top = ranking[0]
    if len(scored) == len(schemes):
        lines = [f"共评估 {len(schemes)} 个方案，全部完成评分。"]
    else:
        lines = [f"共评估 {len(schemes)} 个方案，其中 {len(scored)} 个完成评分。"]
    for order, item in enumerate(ranking, start=1):
        label = item["rating"]["label"] if item.get("rating") else "—"
        suffix = "（未覆盖指标按权重重归一）" if item.get("partial") else ""
        lines.append(f"第 {order} 名：{item['name']}，综合得分 {item['composite']:.2f}，{label}{suffix}。")
    drivers = sorted([s for s in top["scores"] if s.get("score") is not None],
                     key=lambda s: -(float(s["score"]) * s["globalWeight"]))[:3]
    deductions = []
    for item in schemes:
        if item.get("composite") is None:
            deductions.append(f"{item['name']} 未出分：required 缺 {len(item['missingRequired'])} 项。")
        elif item.get("partial"):
            deductions.append(f"{item['name']} 为部分覆盖（optional 缺 {len(item['missingOptional'])} 项），得分按已覆盖权重重归一。")
    return {"summary": "\n".join(lines),
            "topDrivers": [{"id": s["id"], "name": s["name"], "score": s["score"],
                            "weight": s["globalWeight"],
                            "contribution": round(float(s["score"]) * s["globalWeight"], 6)}
                           for s in drivers],
            "deductions": deductions}


def cmd_score(args) -> int:
    weights_payload, weights_path = read_json(args.weights)
    weights = export_lib.normalize_weights(weights_payload)
    scores_payload, scores_path = read_json(args.scores)
    schemes_raw = scores_payload.get("schemes")
    if not isinstance(schemes_raw, list) or not schemes_raw:
        raise lib.AhpError("AHP-E-INPUT", "打分文件需要非空 schemes[]")
    bands = load_bands()
    if args.direction is not None and not str(args.direction).strip():
        raise lib.AhpError("AHP-E-INPUT", "--direction 不能为空字符串")
    direction = load_direction(args.direction) if args.direction else None
    if args.direction and direction is None:
        raise lib.AhpError("AHP-E-INPUT", "未知方向（见 方向适配.yaml）")
    statuses = direction_statuses(direction)
    schemes = []
    for index, item in enumerate(schemes_raw, start=1):
        if not isinstance(item, dict):
            raise lib.AhpError("AHP-E-INPUT", f"schemes[{index}] 应为对象")
        name = item.get("name") or f"方案{index}"
        schemes.append(build_scheme(name, item, weights, bands, statuses))
    scoring_rules = (lib.load_scale_rules().get("scoring") or {})
    payload = {
        "schema": SCHEMA, "schemaVersion": 1, "producer": PRODUCER,
        "demo": bool(getattr(args, "demo", False) or scores_payload.get("demo")),
        "source": [{"kind": "weights", "id": "weights", "path": export_lib.sanitize_path(str(weights_path)),
                    "sha256": export_lib.sha256_file(weights_path)},
                   {"kind": "scores", "id": "scores", "path": export_lib.sanitize_path(str(scores_path)),
                    "sha256": export_lib.sha256_file(scores_path)}],
        "gaps": [], "ruleVersion": int((lib.load_scale_rules().get("version") or 0)),
        "engineVersion": ENGINE_VERSION,
        "scoring": {"scale": scoring_rules.get("scale", 10), "min": scoring_rules.get("min", 0),
                    "max": scoring_rules.get("max", 10), "decimals": scoring_rules.get("decimals", 1),
                    "missingPolicy": scoring_rules.get("missingPolicy", "forbid_zero_fill"),
                    "bandTieBreak": scoring_rules.get("bandTieBreak", "higher_band_wins")},
        "direction": {"type": args.direction or "generic",
                      "basis": "打分参数" if args.direction else "未指定（按 required）"},
        "schemes": schemes,
        "ranking": [{"order": order, "name": item["name"], "composite": item["composite"],
                     "rating": item.get("rating"), "partial": item.get("partial", False)}
                    for order, item in enumerate(sorted([s for s in schemes if s.get("composite") is not None],
                                                        key=lambda s: -s["composite"]), start=1)],
        "conclusion": conclusion_of(schemes, weights),
    }
    if not getattr(args, "no_timestamp", False):
        payload["generatedAt"] = now_iso()
    all_failed = all(s.get("composite") is None for s in schemes)
    any_failed = any(s.get("composite") is None for s in schemes)
    if any(s.get("composite") is None for s in schemes):
        payload["gaps"].append({"code": "AHP-GAP-SCORE", "what": "部分方案 required 缺依据，未出分",
                                "need": "补人工打分或可映射 metric", "owner": "用户"})
    if not args.direction:
        payload["gaps"].append({"code": "AHP-GAP-DIRECTION", "what": "未指定项目方向（按 required 全量打分）",
                                "need": "用 --direction 指定方向后按方向适配排除 not_applicable",
                                "owner": "cardo-status / 用户"})
    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.format == "md":
        print(render_md(payload))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    if all_failed or any_failed:
        return 3
    return 0


def render_md(payload: dict) -> str:
    lines = ["# AHP 项目打分与综合评分", "",
             f"- 方案 {len(payload['schemes'])} 个；打分口径 {payload['scoring']['scale']} 分制"
             f"（{payload['scoring']['min']}-{payload['scoring']['max']}，1 位小数）", "",
             "## 综合得分与评级", "", "| 名次 | 方案 | 综合得分 | 评级 | 说明 |", "|---|---|---|---|---|"]
    for item in payload["ranking"]:
        rating = item.get("rating") or {}
        lines.append(f"| {item['order']} | {item['name']} | {item['composite']:.2f} | "
                     f"{rating.get('label', '—')} | {'部分覆盖' if item.get('partial') else '全覆盖'} |")
    lines += ["", "## 结论", ""] + [line for line in payload["conclusion"]["summary"].splitlines() if line]
    drivers = payload["conclusion"].get("topDrivers") or []
    if drivers:
        lines += ["", "## 主要驱动项（第一名的贡献前 3）", ""]
        for item in drivers:
            lines.append(f"- {item['name']}：得分 {item['score']} × 权重 "
                         f"{float(item.get('weight') or 0):.6f}"
                         f" = 贡献 {item['contribution']:.6f}")
    if payload["conclusion"].get("deductions"):
        lines += ["", "## 限制与缺口", ""] + [f"- {text}" for text in payload["conclusion"]["deductions"]]
    if payload.get("gaps"):
        lines += ["", "## Gaps", ""] + [f"- {gap['code']}：{gap['what']}" for gap in payload["gaps"]]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 项目打分与排序")
    sub = ap.add_subparsers(dest="command", required=True)
    p = sub.add_parser("score", help="按权重给方案打分/评级/排序")
    p.add_argument("--weights", required=True)
    p.add_argument("--scores", required=True)
    p.add_argument("--out", default=None)
    p.add_argument("--direction", default=None)
    p.add_argument("--format", choices=["json", "md"], default="json")
    p.add_argument("--demo", action="store_true",
                   help="声明本次为演示样例（写入过程数据 demo=true；产物另加「演示样例」标识）")
    p.add_argument("--no-timestamp", action="store_true")
    p.set_defaults(func=cmd_score)
    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except lib.AhpError as exc:
        print(json.dumps({"error": {"code": exc.code, "message": exc.message}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return exc.exit_code
    except Exception as exc:  # 兜底：不向用户抛 Traceback
        print(json.dumps({"error": {"code": "AHP-E-INTERNAL", "message": f"{type(exc).__name__}: {exc}"}},
                         ensure_ascii=False, indent=2), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())