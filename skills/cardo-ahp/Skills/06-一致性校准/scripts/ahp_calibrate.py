#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ahp_calibrate.py -- cardo-ahp 一致性校准（诊断 + 候选 + 采用记录）。

纪律（规则 calibration 段）：
  · 最多 3 个候选；每个候选最多改 2 个元素；保持互反；autoApply=false；
  · **不自动采用**：apply 必须带 `--approve-by`（谁/何时），否则 exit 2；
  · 采用前备份原矩阵（`<matrix>.bak`），采用后写 `<matrix>.apply.json` 记录（前后 CR / sha256 / 候选 id / 确认人）。
不做：不改权重表、不生成报告、不替用户选候选。

子命令：
  diagnose --matrix <json> [--format md|json]       定位最不一致元素/行（偏离度、行贡献度）
  suggest  --matrix <json> [--max-candidates 3] [--out <dir>] [--no-timestamp]
  apply    --matrix <json> --candidate <候选.json> --approve-by "<用户/时间>"

退出码：0 成功；1 无需校准（CR 已通过）或无法改进；2 参数/输入错或未确认；3 保留。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402
from _calib_lib import make_candidates, weights_and_deviations  # noqa: E402

TZ = timezone(timedelta(hours=8))
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"
SCHEMA = "cardo/ahp-calibration/v1"
def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def now_iso() -> str:
    return datetime.now(TZ).isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def load_matrix_payload(path_text: str) -> tuple[dict, list[list[float]]]:
    path = Path(path_text)
    if not path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"矩阵文件不存在：{path_text}")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise lib.AhpError("AHP-E-INPUT", f"矩阵文件不是合法 JSON：{exc}") from exc
    if not isinstance(payload, dict):
        raise lib.AhpError("AHP-E-INPUT", "矩阵文件应为 {scale,mode,order,values}")
    scale_id = payload.get("scale")
    mode = payload.get("mode") or "column"
    order = payload.get("order") or (len(payload.get("values") or []) if mode == "column" else 0)
    if not scale_id or not order:
        raise lib.AhpError("AHP-E-SCALE", "矩阵文件缺 scale/order")
    matrix = lib.build_matrix(payload.get("values"), int(order), str(scale_id), str(mode))
    return payload, matrix


def envelope(args, payload: dict, source_path: str) -> dict:
    result = {
        "schema": SCHEMA,
        "schemaVersion": 1,
        "producer": PRODUCER,
        "source": [{"kind": "judgment", "id": payload.get("id") or "matrix",
                    "path": source_path, "sha256": sha256_file(Path(source_path))}],
        "gaps": [],
        "ruleVersion": int((lib.load_scale_rules().get("version") or 0)),
        "engineVersion": ENGINE_VERSION,
    }
    if not getattr(args, "no_timestamp", False):
        result["generatedAt"] = now_iso()
    return result


# ---------------------------------------------------------------- diagnose
def cmd_diagnose(args) -> int:
    payload, matrix = load_matrix_payload(args.matrix)
    diag = weights_and_deviations(matrix)
    result = envelope(args, payload, args.matrix)
    result.update({
        "matrixId": payload.get("id") or "matrix",
        "n": diag["n"],
        "weights": diag["weights"],
        "cr": diag["evaluation"]["consistency"]["cr"],
        "pass": diag["evaluation"]["consistency"]["pass"],
        "worstElements": diag["worstElements"],
        "worstRows": diag["worstRows"],
        "calibrationApplied": False,
    })
    if args.format == "md":
        print(render_diagnose_md(result))
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=_json_default))
    return 0


def render_diagnose_md(result: dict) -> str:
    lines = ["# AHP 一致性诊断", "",
             f"- 矩阵：{result['matrixId']}；阶数 {result['n']}；CR={result['cr']:.6f}；"
             f"判定：{'通过' if result['pass'] else '不通过'}", "",
             "## 最不一致元素（按相对偏离度）", "",
             "| 元素 | 当前值 | 理想值 w_i/w_j | 相对偏离 |", "|---|---|---|---|"]
    for d in result["worstElements"]:
        lines.append(f"| a_{d['row']}{d['col']} | {d['value']:.6g} | {d['ideal']:.6g} | {d['relativeDeviation']:.4f} |")
    lines += ["", "## 最不一致行（行贡献度）", "", "| 行 | 平均偏离 |", "|---|---|"]
    for row in result["worstRows"]:
        lines.append(f"| 第 {row['row']} 行 | {row['meanDeviation']:.4f} |")
    return "\n".join(lines)


# ---------------------------------------------------------------- suggest
def cmd_suggest(args) -> int:
    payload, matrix = load_matrix_payload(args.matrix)
    scale_id = str(payload.get("scale"))
    diag = weights_and_deviations(matrix)
    evaluation = diag["evaluation"]
    if evaluation["consistency"]["pass"]:
        result = envelope(args, payload, args.matrix)
        result.update({
            "matrixId": payload.get("id") or "matrix",
            "cr": evaluation["consistency"]["cr"],
            "message": "CR 已通过（<0.1-1e-12），无需校准；本条不生成候选。",
            "candidates": [],
            "calibrationApplied": False,
        })
        print(json.dumps(result, ensure_ascii=False, indent=2, default=_json_default))
        return 1
    cal = (lib.load_scale_rules().get("calibration") or {})
    requested = args.max_candidates if args.max_candidates is not None else cal.get("maxCandidates", 3)
    if int(requested) < 1:
        raise lib.AhpError("AHP-E-INPUT", f"--max-candidates 必须 ≥1（实际 {requested}）")
    max_candidates = min(int(requested), 3)
    candidates = make_candidates(diag, matrix, scale_id, max_candidates,
                                 int(cal.get("maxChangedCells", 2)), float(cal.get("minCrImprovement", 0.01)))
    if not candidates:
        result = envelope(args, payload, args.matrix)
        result.update({"matrixId": payload.get("id") or "matrix",
                       "cr": evaluation["consistency"]["cr"], "candidates": [],
                       "message": "未找到满足最小改动/最小改进的候选；建议人工重构矩阵。"})
        print(json.dumps(result, ensure_ascii=False, indent=2, default=_json_default))
        return 1
    result = envelope(args, payload, args.matrix)
    result.update({
        "matrixId": payload.get("id") or "matrix",
        "crBefore": evaluation["consistency"]["cr"],
        "weightsBefore": diag["weights"],
        "worstElements": diag["worstElements"][:3],
        "candidates": [{k: v for k, v in c.items() if k != "matrix"} for c in candidates],
        "calibrationApplied": False,
        "note": "候选未采用；用户确认后用 apply --candidate <候选文件> --approve-by <用户/时间>。",
    })
    if args.out:
        out_dir = Path(args.out)
        out_dir.mkdir(parents=True, exist_ok=True)
        for c in candidates:
            write_candidate(out_dir / f"候选{c['id']}.json", c, payload, result)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=_json_default))
    return 0


def write_candidate(path: Path, candidate: dict, payload: dict, result: dict) -> None:
    body = {
        "schema": SCHEMA,
        "schemaVersion": 1,
        "producer": PRODUCER,
        "generatedAt": result.get("generatedAt"),
        "matrixId": payload.get("id") or "matrix",
        "baseSha256": result["source"][0].get("sha256"),
        "id": candidate["id"],
        "changedCells": candidate["changedCells"],
        "crBefore": candidate["crBefore"],
        "crAfter": candidate["crAfter"],
        "maxRelativeChange": candidate["maxRelativeChange"],
        "candidateMatrix": {"scale": payload.get("scale"), "mode": "full",
                            "order": len(candidate["matrix"]), "values": candidate["matrix"]},
        "applied": False,
    }
    path.write_text(json.dumps(body, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- apply
def cmd_apply(args) -> int:
    if not args.approve_by or not str(args.approve_by).strip():
        raise lib.AhpError("AHP-E-CONFIRM", "采用校准候选必须提供 --approve-by <用户/时间>（不自动采用）")
    payload, matrix = load_matrix_payload(args.matrix)
    candidate_path = Path(args.candidate)
    if not candidate_path.is_file():
        raise lib.AhpError("AHP-E-INPUT", f"候选文件不存在：{args.candidate}")
    candidate = json.loads(candidate_path.read_text(encoding="utf-8"))
    candidate_matrix = (candidate.get("candidateMatrix") or {}).get("values")
    if not isinstance(candidate_matrix, list):
        raise lib.AhpError("AHP-E-INPUT", "候选文件缺 candidateMatrix.values")
    order = int((candidate.get("candidateMatrix") or {}).get("order") or len(candidate_matrix))
    if order != len(matrix):
        raise lib.AhpError(
            "AHP-E-ORDER",
            f"候选矩阵阶数 {order} ≠ 原矩阵阶数 {len(matrix)}（禁止跨阶替换）",
        )
    validated = lib.build_matrix(candidate_matrix, order, str(payload.get("scale")), "full")
    actual_changed = []
    for i in range(order):
        for j in range(i + 1, order):
            if abs(matrix[i][j] - validated[i][j]) > lib.rules().compare_tol:
                actual_changed.append({"row": i + 1, "col": j + 1,
                                       "from": matrix[i][j], "to": validated[i][j]})
    declared = candidate.get("changedCells") or []
    declared_pairs = sorted((min(int(c["row"]), int(c["col"])), max(int(c["row"]), int(c["col"]))) for c in declared)
    actual_pairs = sorted((min(c["row"], c["col"]), max(c["row"], c["col"])) for c in actual_changed)
    if declared_pairs and declared_pairs != actual_pairs:
        raise lib.AhpError(
            "AHP-E-INPUT",
            f"候选 changedCells {declared_pairs} 与实际改动 {actual_pairs} 不一致（拒绝采用）",
        )
    before_sha = sha256_file(Path(args.matrix))
    before_eval = lib.evaluate(matrix)
    after_eval = lib.evaluate(validated)
    if not after_eval["consistency"]["pass"]:
        raise lib.AhpError("AHP-E-CALIB", "候选矩阵 CR 仍未通过（<0.1），拒绝采用", 1)
    base_backup = Path(str(args.matrix) + ".bak")
    backup = base_backup
    counter = 2
    while backup.exists():
        backup = Path(str(args.matrix) + f".bak{counter}")
        counter += 1
    backup.write_bytes(Path(args.matrix).read_bytes())
    adopted = {"scale": payload.get("scale"), "mode": "full", "order": int(order),
               "id": payload.get("id") or "matrix", "values": validated}
    Path(args.matrix).write_text(json.dumps(adopted, ensure_ascii=False, indent=2), encoding="utf-8")
    after_sha = sha256_file(Path(args.matrix))
    record = {
        "schema": SCHEMA, "schemaVersion": 1, "producer": PRODUCER,
        "matrixId": payload.get("id") or "matrix",
        "candidateId": candidate.get("id"),
        "changedCellsActual": actual_changed,
        "approveBy": args.approve_by,
        "changedCells": candidate.get("changedCells"),
        "crBefore": before_eval["consistency"]["cr"],
        "crAfter": after_eval["consistency"]["cr"],
        "sha256Before": before_sha, "sha256After": after_sha,
        "backup": backup.name,
        "calibrationApplied": True,
    }
    if not getattr(args, "no_timestamp", False):
        record["generatedAt"] = now_iso()
    record_path = Path(str(args.matrix) + ".apply.json")
    record_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(record, ensure_ascii=False, indent=2, default=_json_default))
    return 0


def _json_default(value):
    if isinstance(value, float):
        return value
    raise TypeError(f"不可序列化：{type(value)}")


# ---------------------------------------------------------------- main
def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 一致性校准（不自动采用）")
    sub = ap.add_subparsers(dest="command", required=True)

    p_diag = sub.add_parser("diagnose", help="定位最不一致元素/行")
    p_diag.add_argument("--matrix", required=True)
    p_diag.add_argument("--format", choices=["json", "md"], default="json")
    p_diag.set_defaults(func=cmd_diagnose, no_timestamp=False)

    p_sug = sub.add_parser("suggest", help="生成 ≤3 个候选最小改动矩阵")
    p_sug.add_argument("--matrix", required=True)
    p_sug.add_argument("--max-candidates", type=int, default=3)
    p_sug.add_argument("--out", default=None)
    p_sug.add_argument("--no-timestamp", action="store_true")
    p_sug.set_defaults(func=cmd_suggest)

    p_app = sub.add_parser("apply", help="采用候选（必须 --approve-by）")
    p_app.add_argument("--matrix", required=True)
    p_app.add_argument("--candidate", required=True)
    p_app.add_argument("--approve-by", default=None)
    p_app.add_argument("--no-timestamp", action="store_true")
    p_app.set_defaults(func=cmd_apply)

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