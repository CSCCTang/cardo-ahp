#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_ahp_lib.py -- cardo-ahp 计算底座（标准库，零第三方依赖）。

职责（唯一真相，CLI 与测试都走这里）：
  · 读规则：`../../规则/标度与口径.yaml`（RI 表/容差/CR 口径/错误码 priority/三口径换算）
  · 矩阵：三口径（1-9 / 0.1-1 / ratio）× 三种输入形态（full / column / pairwise）→ 统一互反矩阵
  · 校验：阶数 2-10、数值合法、正性、对角线=1、互反；**只报错不修正**
  · 权重：几何平均（行几何平均归一化）+ 特征值（幂迭代，1e-12 / 1000 次）
  · 一致性：λmax（Rayleigh 商）、CI（|CI|<1e-12 截断 0）、RI 表 n=1..10、CR=max(0,CI/RI)、
    CR 判定 `cr < 0.1 - 1e-12`；n<=2 或 RI=0 → CR=0
退出码：由 CLI 决定；本模块只抛 AhpError(code, message, exit_code)。
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from _yaml_lite import YamlLiteError, load_file  # noqa: E402

RULES_DIR = HERE.parent / "规则"
SCALE_FILE = RULES_DIR / "标度与口径.yaml"
INDICATOR_FILE = RULES_DIR / "指标体系.yaml"

WARNINGS: list[str] = []


class AhpError(ValueError):
    """带错误码的输入/校验错误。exit_code：1=校验不通过；2=输入或参数错。"""

    def __init__(self, code: str, message: str, exit_code: int = 2, detail: dict | None = None):
        super().__init__(f"[{code}] {message}")
        self.code = code
        self.message = message
        self.exit_code = exit_code
        self.detail = detail or {}


def _need(cond: bool, code: str, message: str, exit_code: int = 2) -> None:
    if not cond:
        raise AhpError(code, message, exit_code)


# ---------------------------------------------------------------- 规则
def load_scale_rules() -> dict:
    try:
        return load_file(SCALE_FILE)
    except (YamlLiteError, FileNotFoundError) as exc:
        raise AhpError("AHP-E-RULES", f"标度口径规则不可读：{exc}", 2) from exc


def load_indicators() -> dict:
    try:
        return load_file(INDICATOR_FILE)
    except (YamlLiteError, FileNotFoundError) as exc:
        raise AhpError("AHP-E-RULES", f"指标体系规则不可读：{exc}", 2) from exc


class Rules:
    """规则快照（一次性读取，字段做防御性取值）。"""

    def __init__(self) -> None:
        cfg = load_scale_rules()
        self.raw = cfg
        self.order_min, self.order_max = (cfg.get("validation") or {}).get("orderRange") or [2, 10]
        self.diagonal_one = bool((cfg.get("validation") or {}).get("diagonalMustBeOne", True))
        self.recip_tol = float((cfg.get("validation") or {}).get("reciprocityTolerance", 1e-12))
        self.compare_tol = float((cfg.get("numeric") or {}).get("compareTolerance", 1e-12))
        self.ci_zero_tol = float((cfg.get("numeric") or {}).get("ciZeroTolerance", 1e-12))
        self.eigen_tol = float((cfg.get("numeric") or {}).get("eigenTolerance", 1e-12))
        self.eigen_max_iter = int((cfg.get("numeric") or {}).get("eigenIterationsMax", 1000))
        cons = cfg.get("consistency") or {}
        self.ri = {int(k): float(v) for k, v in (cons.get("riTable") or {}).items()}
        self.cr_threshold = float(cons.get("crThreshold", 0.1))
        self.cr_tolerance = float(cons.get("crCompareTolerance", 1e-12))
        self.cr_formula = str(cons.get("crPassRule") or "")
        self.scales = {s.get("id"): s for s in (cfg.get("scales") or [])}
        self.profile = cfg.get("defaultProfile") or {}
        self.error_codes = {e.get("code"): e for e in (cfg.get("errorCodes") or [])}
        self.warn_codes = {w.get("code") for w in (cfg.get("warnCodes") or [])}

    def scale(self, scale_id: str) -> dict:
        _need(scale_id in self.scales, "AHP-E-SCALE", f"未知标度：{scale_id}（可选 {sorted(self.scales)}）")
        return self.scales[scale_id]

    def check_cr(self, cr: float) -> bool:
        # crPassRule = pass = cr < crThreshold - crCompareTolerance
        return cr < self.cr_threshold - self.cr_tolerance


_RULES: Rules | None = None


def rules() -> Rules:
    global _RULES
    if _RULES is None:
        _RULES = Rules()
    return _RULES


# ---------------------------------------------------------------- 数值与矩阵
def as_number(value, label: str = "元素"):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise AhpError("AHP-E-SCALE", f"{label}不是数值：{value!r}（字符串/NaN/布尔均不接受）")
    number = float(value)
    if math.isnan(number) or math.isinf(number):
        raise AhpError("AHP-E-SCALE", f"{label}不是有限数值：{value!r}")
    return number


def _input_range(scale_id: str):
    """column 模式的**输入值**允许范围；ratio 无列向量模式（None）。"""
    if scale_id == "scale_1_9":
        return (1.0, 9.0)
    if scale_id == "scale_0_1":
        return (0.1, 1.0)
    return None


def _element_range(scale_id: str):
    """矩阵元素（互反）允许范围；ratio 不设硬域（仅警告），返回 None。"""
    if scale_id == "scale_1_9":
        return (1.0 / 9.0, 9.0)
    if scale_id == "scale_0_1":
        return (0.1, 10.0)          # v_i/v_j，v ∈ [0.1,1]
    return None


def _check_value_in_scale(value: float, scale_id: str, label: str, for_input: bool = False) -> None:
    if value <= 0 or not math.isfinite(value):
        raise AhpError("AHP-E-SCALE", f"{label}={value} 非正或非有限（只报错不修正）")
    lo_hi = _input_range(scale_id) if for_input else _element_range(scale_id)
    tol = rules().compare_tol
    if lo_hi is not None:
        lo, hi = lo_hi
        if value < lo - tol or value > hi + tol:
            raise AhpError(
                "AHP-E-SCALE",
                f"{label}={value} 超出标度 {scale_id} 合法域 [{lo:g}, {hi:g}]（不自动截断）",
            )
    # 警告域：构造后的互反元素超出 1-9 语义 [1/9, 9]（0.1-1 与 ratio 都可能）
    if scale_id in ("scale_0_1", "ratio") and (value < 1.0 / 9.0 - tol or value > 9.0 + tol):
        WARNINGS.append(f"[AHP-W-SCALE-EXTREME] {label}={value} 超出 [1/9,9]（仅警告，不截断）")


def build_matrix(values, order: int, scale_id: str, mode: str) -> list[list[float]]:
    """三口径 × 三形态 → 统一互反矩阵。只报错，不修正。"""
    r = rules()
    scale = r.scale(scale_id)
    _need(isinstance(order, int) and r.order_min <= order <= r.order_max,
          "AHP-E-ORDER", f"矩阵阶数 {order} 不在 [{r.order_min},{r.order_max}]")
    _need(mode in ("full", "column", "pairwise"), "AHP-E-MISSING", f"未知输入形态：{mode}")

    if mode == "column":
        _need(scale_id != "ratio", "AHP-E-SCALE", "ratio 标度不支持 column 模式（列向量展开只属于 0.1-1）")
        _need(isinstance(values, (list, tuple)) and len(values) == order,
              "AHP-E-MISSING", f"column 模式需要 {order} 个值，实际 {len(values) if hasattr(values, '__len__') else '非列表'}")
        vec = [as_number(v, f"列向量第 {i + 1} 项") for i, v in enumerate(values)]
        for i, v in enumerate(vec):
            _check_value_in_scale(v, scale_id, f"列向量第 {i + 1} 项", for_input=True)
        return [[vec[i] / vec[j] for j in range(order)] for i in range(order)]

    if mode == "pairwise":
        _need(isinstance(values, (list, tuple)), "AHP-E-MISSING", "pairwise 模式需要 [[i,j,value], ...] 列表")
        matrix = [[None] * order for _ in range(order)]
        for idx, item in enumerate(values):
            _need(isinstance(item, (list, tuple)) and len(item) == 3,
                  "AHP-E-MISSING", f"pairwise 第 {idx + 1} 项应为 [i, j, value]")
            i, j, value = item
            _need(isinstance(i, int) and isinstance(j, int) and not isinstance(i, bool) and not isinstance(j, bool)
                  and 1 <= i <= order and 1 <= j <= order and i != j,
                  "AHP-E-MISSING", f"pairwise 第 {idx + 1} 项下标越界或自比较：{item!r}")
            number = as_number(value, f"pairwise[{i},{j}]")
            _check_value_in_scale(number, scale_id, f"pairwise[{i},{j}]")
            if matrix[i - 1][j - 1] is not None and abs(matrix[i - 1][j - 1] - number) > r.recip_tol:
                raise AhpError("AHP-E-NONRECIPROCAL", f"pairwise 同一格重复且值冲突：({i},{j})")
            matrix[i - 1][j - 1] = number
            if matrix[j - 1][i - 1] is not None:
                _need(abs(matrix[j - 1][i - 1] * number - 1.0) <= r.recip_tol,
                      "AHP-E-NONRECIPROCAL",
                      f"pairwise a_{i}{j}·a_{j}{i}≠1（{number}×{matrix[j - 1][i - 1]}）")
            # 未提供的对称格**不自动补**（规则：缺 a_ji → AHP-E-MISSING），由下方缺项检查报错
        for i in range(order):
            for j in range(order):
                if i != j and matrix[i][j] is None:
                    raise AhpError("AHP-E-MISSING", f"pairwise 缺 {order} 阶矩阵的格 ({i + 1},{j + 1})")
            matrix[i][i] = 1.0
        return validate_matrix(matrix, order, scale_id, "pairwise")

    # full
    _need(isinstance(values, (list, tuple)) and len(values) == order,
          "AHP-E-MISSING", f"full 模式需要 {order} 行，实际 {len(values) if hasattr(values, '__len__') else '非列表'}")
    matrix = []
    for i, row in enumerate(values):
        _need(isinstance(row, (list, tuple)) and len(row) == order,
              "AHP-E-MISSING", f"full 第 {i + 1} 行列数不是 {order}")
        matrix.append([as_number(v, f"full[{i + 1},{j + 1}]") for j, v in enumerate(row)])
    return validate_matrix(matrix, order, scale_id, "full")


def validate_matrix(matrix: list[list[float]], order: int, scale_id: str, mode: str) -> list[list[float]]:
    r = rules()
    for i in range(order):
        _check_value_in_scale(matrix[i][i], scale_id, f"对角线 a_{i + 1}{i + 1}") if r.diagonal_one else None
        if r.diagonal_one and abs(matrix[i][i] - 1.0) > r.recip_tol:
            raise AhpError("AHP-E-DIAGONAL", f"对角线 a_{i + 1}{i + 1}={matrix[i][i]}≠1")
        for j in range(i + 1, order):
            a, b = matrix[i][j], matrix[j][i]
            _check_value_in_scale(a, scale_id, f"a_{i + 1}{j + 1}")
            _check_value_in_scale(b, scale_id, f"a_{j + 1}{i + 1}")
            if abs(a * b - 1.0) > r.recip_tol:
                raise AhpError(
                    "AHP-E-NONRECIPROCAL",
                    f"a_{i + 1}{j + 1}·a_{j + 1}{i + 1}={a * b}≠1（差值 {abs(a * b - 1.0):.3e}）",
                )
    return matrix


# ---------------------------------------------------------------- 权重
def geometric_weights(matrix: list[list[float]]) -> list[float]:
    n = len(matrix)
    geo = [math.prod(matrix[i]) ** (1.0 / n) for i in range(n)]
    total = sum(geo)
    if total <= 0 or not math.isfinite(total):
        raise AhpError("AHP-E-SCALE", "行几何平均之和非正，无法归一（检查矩阵数值）")
    return [g / total for g in geo]


def eigen_weights(matrix: list[list[float]]) -> tuple[list[float], float, int]:
    """幂迭代求主特征向量与 λmax（Rayleigh 商）。返回 (权重, λmax, 迭代次数)。"""
    r = rules()
    n = len(matrix)
    vector = [1.0 / n] * n
    for iteration in range(1, r.eigen_max_iter + 1):
        product = [sum(matrix[i][j] * vector[j] for j in range(n)) for i in range(n)]
        norm = sum(product)
        if norm <= 0 or not math.isfinite(norm):
            raise RuntimeError("幂迭代出现非正/非有限归一因子")
        candidate = [v / norm for v in product]
        diff = max(abs(candidate[i] - vector[i]) for i in range(n))
        vector = candidate
        if diff <= r.eigen_tol:
            break
    else:
        raise RuntimeError(f"幂迭代 {r.eigen_max_iter} 次未收敛（正互反矩阵不应发生）")
    # Rayleigh 商：w^T A w（w 已归一化，w^T w = 1 仅在 2-范数时成立；这里用 Aw / w 的加权平均）
    product = [sum(matrix[i][j] * vector[j] for j in range(n)) for i in range(n)]
    lam = sum(product[i] / vector[i] for i in range(n)) / n
    return [v / sum(vector) for v in vector], lam, iteration


def consistency(matrix: list[list[float]], lam: float) -> dict:
    r = rules()
    n = len(matrix)
    ci_raw = (lam - n) / (n - 1)
    ci = 0.0 if abs(ci_raw) < r.ci_zero_tol else ci_raw
    ri = r.ri.get(n, 0.0)
    cr = 0.0 if (n <= 2 or ri == 0.0) else max(0.0, ci / ri)
    return {
        "n": n,
        "lambdaMax": lam,
        "ci": ci,
        "ciRaw": ci_raw,
        "ri": ri,
        "cr": cr,
        "pass": r.check_cr(cr),
        "truncatedCi": abs(ci_raw) < r.ci_zero_tol,
        "crThreshold": r.cr_threshold,
        "crTolerance": r.cr_tolerance,
    }


def evaluate(matrix: list[list[float]]) -> dict:
    geo = geometric_weights(matrix)
    eigen, lam, iterations = eigen_weights(matrix)
    cons = consistency(matrix, lam)
    delta = [abs(geo[i] - eigen[i]) for i in range(len(matrix))]
    return {
        "n": len(matrix),
        "weightsGeometric": geo,
        "weightsEigen": eigen,
        "lambdaMax": lam,
        "eigenIterations": iterations,
        "consistency": cons,
        "delta": {"perItem": delta, "maxAbs": max(delta) if delta else 0.0},
    }


# ---------------------------------------------------------------- 指标目录
def indicator_catalog() -> dict:
    data = load_indicators()
    criteria = [
        {
            "id": c.get("id"),
            "name": c.get("name"),
            "displayName": c.get("displayName") or {},
            "order": c.get("order"),
            "description": c.get("description"),
        }
        for c in (data.get("criteria") or [])
    ]
    indicators = [
        {
            "id": i.get("id"),
            "criterionId": i.get("criterionId"),
            "name": i.get("name"),
            "direction": i.get("direction"),
            "directionNote": i.get("directionNote"),
            "definition": i.get("definition"),
            "source": i.get("source"),
            "templateCell": i.get("templateCell"),
            "anchors": i.get("anchors") or [],
        }
        for i in (data.get("indicators") or [])
    ]
    return {
        "schema": "cardo/ahp-indicators/v1",
        "version": data.get("version"),
        "goal": data.get("goal") or {},
        "criteria": criteria,
        "indicators": indicators,
        "namingMap": data.get("namingMap") or [],
    }


# ---------------------------------------------------------------- 批量矩阵集
def load_matrix_set(payload) -> list[dict]:
    _need(isinstance(payload, dict) and isinstance(payload.get("matrices"), list),
          "AHP-E-MISSING", "矩阵集需要 {matrices: [...]} 结构")
    _need(len(payload["matrices"]) > 0, "AHP-E-MISSING", "矩阵集为空（至少 1 个矩阵）")
    result = []
    seen = set()
    for idx, item in enumerate(payload["matrices"]):
        _need(isinstance(item, dict), "AHP-E-MISSING", f"矩阵集第 {idx + 1} 项不是对象")
        mid = item.get("id")
        _need(isinstance(mid, str) and bool(mid), "AHP-E-MISSING", f"矩阵集第 {idx + 1} 项缺 id")
        _need(mid not in seen, "AHP-E-MISSING", f"矩阵 id 重复：{mid}")
        seen.add(mid)
        order = item.get("order") or (len(item.get("values") or []) if item.get("mode") == "column" else 0)
        _need(isinstance(order, int) and order > 0, "AHP-E-ORDER", f"{mid} 缺 order 且无法从 column 推断")
        scale_id = item.get("scale") or (rules().profile.get("criteriaScale") if item.get("level") == "criteria" else rules().profile.get("indicatorsScale"))
        mode = item.get("mode") or "column"
        matrix = build_matrix(item.get("values"), order, scale_id, mode)
        entry = {
            "id": mid,
            "level": item.get("level") or "unknown",
            "scale": scale_id,
            "mode": mode,
            "order": order,
            "matrix": matrix,
        }
        entry.update(evaluate(matrix))
        result.append(entry)
    return result


def weighted_global(criteria_weights: list[float], local_by_criterion: dict[str, list[float]], criterion_order: list[str]) -> list[float]:
    global_weights = []
    for cid, weight in zip(criterion_order, criteria_weights):
        local = local_by_criterion.get(cid)
        _need(local is not None, "AHP-E-MISSING", f"缺准则 {cid} 的指标层权重")
        global_weights.extend([weight * w for w in local])
    return global_weights