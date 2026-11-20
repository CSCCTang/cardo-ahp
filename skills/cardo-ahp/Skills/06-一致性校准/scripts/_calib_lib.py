#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_calib_lib.py -- cardo-ahp 校准算法（诊断/候选生成），从 ahp_calibrate.py 拆出以控行数。

纯函数，不写盘；CLI 负责参数、文件与采用记录。标准库零依赖。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402

SNAP_1_9 = [1.0 / 9, 1.0 / 7, 1.0 / 5, 1.0 / 3, 1.0, 3.0, 5.0, 7.0, 9.0]
SNAP_0_1 = [0.1, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.75, 1.0]


def weights_and_deviations(matrix: list[list[float]]) -> dict:
    r = lib.rules()
    n = len(matrix)
    weights = lib.geometric_weights(matrix)
    deviations = []
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            ideal = weights[i] / weights[j]
            observed = matrix[i][j]
            deviations.append({
                "row": i + 1, "col": j + 1, "value": observed, "ideal": ideal,
                "relativeDeviation": abs(observed - ideal) / ideal,
            })
    row_scores = []
    for i in range(n):
        own = [d["relativeDeviation"] for d in deviations if d["row"] == i + 1]
        row_scores.append({"row": i + 1, "meanDeviation": sum(own) / len(own) if own else 0.0})
    row_scores.sort(key=lambda x: -x["meanDeviation"])
    ranked = sorted(deviations, key=lambda d: (-d["relativeDeviation"], d["row"], d["col"]))
    evaluation = lib.evaluate(matrix)
    return {
        "n": n,
        "weights": weights,
        "evaluation": evaluation,
        "worstElements": ranked[:5],
        "worstRows": row_scores[:3],
        "tolerance": r.compare_tol,
    }


def snap(value: float, scale_id: str, matrix_range: tuple[float, float]) -> float:
    candidates = SNAP_1_9 if scale_id == "scale_1_9" else SNAP_0_1
    lo, hi = matrix_range
    best = min(candidates, key=lambda c: abs(c - value))
    return min(max(best, lo), hi)


def build_candidate(base: list[list[float]], changes: list[dict], scale_id: str) -> list[list[float]]:
    matrix = [row[:] for row in base]
    n = len(matrix)
    lo, hi = (1.0 / 9.0, 9.0) if scale_id == "scale_1_9" else (0.1, 10.0)
    for change in changes:
        i, j = int(change["row"]) - 1, int(change["col"]) - 1
        if i == j or not (0 <= i < n and 0 <= j < n):
            raise lib.AhpError("AHP-E-INPUT", f"候选改动坐标非法：({i + 1},{j + 1})")
        value = min(max(float(change["to"]), lo), hi)
        matrix[i][j] = value
        matrix[j][i] = 1.0 / value
    # 校验候选仍为正互反且阶数不变
    return lib.validate_matrix(matrix, n, scale_id, "full")


def make_candidates(diag: dict, matrix: list[list[float]], scale_id: str,
                    max_candidates: int, max_changed: int, min_improvement: float) -> list[dict]:
    r = lib.rules()
    lo, hi = (1.0 / 9.0, 9.0) if scale_id == "scale_1_9" else (0.1, 10.0)
    cr_before = diag["evaluation"]["consistency"]["cr"]
    ranked = diag["worstElements"]
    plans: list[list[dict]] = []
    if ranked:
        top = ranked[0]
        ideal = top["ideal"]
        plans.append([{"row": top["row"], "col": top["col"], "from": top["value"],
                       "to": round(min(max(ideal, lo), hi), 2)}])
        plans.append([{"row": top["row"], "col": top["col"], "from": top["value"],
                       "to": snap(ideal, scale_id, (lo, hi))}])
        if len(ranked) > 1 and max_changed >= 2:
            second = ranked[1]
            plans.append([
                {"row": top["row"], "col": top["col"], "from": top["value"], "to": snap(ideal, scale_id, (lo, hi))},
                {"row": second["row"], "col": second["col"], "from": second["value"],
                 "to": snap(second["ideal"], scale_id, (lo, hi))},
            ])
    candidates: list[dict] = []
    seen: set[str] = set()
    for plan in plans:
        if len(candidates) >= max_candidates:
            break
        changed = [c for c in plan if abs(float(c["from"]) - float(c["to"])) > r.compare_tol]
        if not changed or len(changed) > max_changed:
            continue
        try:
            candidate_matrix = build_candidate(matrix, changed, scale_id)
        except lib.AhpError:
            continue
        evaluation = lib.evaluate(candidate_matrix)
        cr_after = evaluation["consistency"]["cr"]
        if cr_after > cr_before - min_improvement:
            continue
        key = json.dumps(changed, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        max_rel = max(abs(float(c["to"]) - float(c["from"])) / float(c["from"]) for c in changed)
        candidates.append({
            "id": f"C{len(candidates) + 1}",
            "changedCells": changed,
            "crBefore": cr_before,
            "crAfter": cr_after,
            "maxRelativeChange": max_rel,
            "note": "只改最小必要元素；保持互反；crAfter < crBefore。采用与否由用户决定。",
            "matrix": candidate_matrix,
        })
    return candidates
