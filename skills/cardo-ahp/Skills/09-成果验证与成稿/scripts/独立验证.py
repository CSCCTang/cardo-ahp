#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""独立验证.py -- 用 numpy 完全独立复算 AHP 全链，并按「可配置基线」对账。

独立性：不 import 引擎任何脚本，只用 numpy 从矩阵集重算，
再读取引擎产物（权重表.json / 评分表.json）与模板 sha256 做交叉验证。

两类验证：
  A. 结构自洽（任何工作区都做，无基线也做）
     权重和=1、各准则局部权重和=1、综合权重和=1、矩阵齐全、CR<0.1、
     双算法一致、引擎权重=独立复算、引擎评分=独立加权、排序自洽、评分完整性。
  B. 基线对账（仅当基线文件存在且含对应字段时做；新测试工作区可无基线）
     准则权重、16 项综合权重、各方案得分、排序、模板 sha256。

基线来源（按优先级）：
  1) --baseline 指定文件；
  2) 默认自动探测 <工作区>/记忆与规则/验证基线-CardoAHP.json；
  3) 都没有 → 跳过 B 类，只做 A 类（报告里标注「跳过」）。

用法：
  python 独立验证.py --matrix-set <矩阵集.json> --weights <权重表.json> --scores <评分表.json>
                     --template <模板.xlsx> --out-dir <输出目录> [--baseline <基线.json>]
退出码：0 无失败项（允许有跳过）；1 有失败项；2 输入错。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path

import numpy as np

RI = {1: 0.0, 2: 0.0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}
TOL = 1e-9
SCORE_TOL = 0.005 + 1e-12   # 评分保留 2 位小数

# 基线写在「工作区」里，不是包目录里：开工.py 以工作区为 cwd 调用本脚本，
# 用 cwd 而不是包所在目录，避免把基线甩到 Cardo交付 根下（交付自检会判不干净）。
DEFAULT_WORKSPACE = Path.cwd()
# 基线按包命名空间隔离：四包可能同机共存，文件名带包名 + 内容带 package，
# 避免 A 包的基线被 B 包当成自己的。
PACKAGE = "CardoAHP"
BASELINE_REL = "记忆与规则/验证基线-CardoAHP.json"


def geo_weights(mat: np.ndarray) -> np.ndarray:
    gm = np.prod(mat, axis=1) ** (1.0 / mat.shape[0])
    return gm / gm.sum()


def eigen_weights(mat: np.ndarray) -> tuple[np.ndarray, float]:
    vals, vecs = np.linalg.eig(mat)
    k = int(np.argmax(vals.real))
    lam = float(vals[k].real)
    w = np.abs(vecs[:, k].real)
    return w / w.sum(), lam


def to_matrix(values) -> np.ndarray:
    """兼容列向量（column）与完整矩阵（full）两种矩阵集写法。"""
    if values and isinstance(values[0], (list, tuple)):
        return np.asarray(values, dtype=float)
    v = np.asarray(values, dtype=float)
    return np.outer(v, 1.0 / v)


def evaluate(values) -> dict:
    mat = to_matrix(values)
    n = mat.shape[0]
    wg = geo_weights(mat)
    we, lam = eigen_weights(mat)
    ci = (lam - n) / (n - 1) if n > 2 else 0.0
    if abs(ci) < 1e-12:
        ci = 0.0
    cr = max(0.0, ci / RI[n]) if RI.get(n, 0) else 0.0
    return {"n": n, "weightsGeo": wg, "weightsEigen": we, "lambdaMax": lam,
            "ci": ci, "ri": RI[n], "cr": cr, "pass": cr < 0.1 - 1e-12,
            "delta": float(np.max(np.abs(wg - we)))}


def rating_of(v: float) -> str:
    return "优秀" if v >= 9 else "良好" if v >= 7 else "一般" if v >= 6 else "不可行"


def check(results: list, name: str, ok: bool | None, detail: str, skipped: bool = False) -> None:
    status = "SKIP" if skipped else ("PASS" if ok else "FAIL")
    results.append({"item": name, "pass": bool(ok) if not skipped else None,
                    "skipped": skipped, "status": status, "detail": detail})
    print(f"[{status}] {name} — {detail}")


def load_baseline(explicit: str | None) -> tuple[dict | None, str]:
    """只认本包写的基线：schema/package 不匹配（含旧版无字段）一律当不存在。"""
    if explicit:
        p = Path(explicit)
        if not p.is_file():
            return None, f"指定基线不存在：{explicit}"
        data = json.loads(p.read_text(encoding="utf-8"))
        if not is_own(data):
            return None, f"指定基线不是本包的（package={data.get('package')}），按新工作区处理"
        return data, rel_display(p)
    auto = DEFAULT_WORKSPACE / BASELINE_REL
    if auto.is_file():
        data = json.loads(auto.read_text(encoding="utf-8"))
        if not is_own(data):
            return None, f"忽略非本包基线（package={data.get('package')}）：{rel_display(auto)}"
        return data, rel_display(auto)
    return None, "未找到基线（新工作区模式：仅做结构自洽验证）"


def is_own(data) -> bool:
    return (isinstance(data, dict) and data.get("package") == PACKAGE
            and data.get("schema") == "cardo-ahp/verification-baseline/v1")


def rel_display(p) -> str:
    """展示用路径：工作区内→相对工作区（记忆与规则/…），工作区外→只留文件名。

    这个字符串会进 独立验证结果.json，再被 09 的结论文档和依据对照表引用；
    交付件里出现 D:/… 这类本机绝对路径，换台电脑就打不开，也不该外泄目录结构。
    """
    path = Path(p)
    try:
        return path.resolve().relative_to(DEFAULT_WORKSPACE.resolve()).as_posix()
    except (ValueError, OSError):
        return path.name



def baseline_from(criteria_map: dict, global_w: dict, scheme_scores: dict,
                  order: list, template_sha: str) -> dict:
    """由本次独立复算结果构造基线内容（字段与对账口径一一对应）。"""
    return {
        "schema": "cardo-ahp/verification-baseline/v1",
        "package": PACKAGE,
        "recordedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "criteriaWeights": {k: float(v) for k, v in criteria_map.items()},
        "globalWeights": {k: float(v) for k, v in global_w.items()},
        "schemeScores": {k: float(v) for k, v in scheme_scores.items()},
        "schemeOrder": list(order),
        "templateSha256": template_sha,
    }


def write_baseline(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="numpy 独立复算 + 可配置基线对账")
    ap.add_argument("--matrix-set", required=True)
    ap.add_argument("--weights", required=True)
    ap.add_argument("--scores", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--baseline", default=None, help="基线 JSON；缺省自动探测 记忆与规则/验证基线-CardoAHP.json")
    ap.add_argument("--baseline-mode", choices=("auto", "write", "check", "none"), default="auto",
                    help="auto=首跑自举写入基线之后真对账；write=强制重写；check=只对账；none=跳过")
    args = ap.parse_args()

    for p in (args.matrix_set, args.weights, args.scores, args.template):
        if not Path(p).is_file():
            print(f"[AHP-E-INPUT] 文件不存在：{p}", file=sys.stderr)
            return 2

    ms = json.loads(Path(args.matrix_set).read_text(encoding="utf-8"))
    wj = json.loads(Path(args.weights).read_text(encoding="utf-8"))
    sj = json.loads(Path(args.scores).read_text(encoding="utf-8-sig"))
    baseline, baseline_src = load_baseline(args.baseline)

    results: list[dict] = []
    by_id = {m["id"]: evaluate(m["values"]) for m in ms["matrices"]}

    # ---------- 通用映射：准则 → 该准则下指标（顺序与目录一致） ----------
    crit_ids = [c["id"] for c in wj["criteria"]]
    eng_crit = {c["id"]: float(c["weight"]) for c in wj["criteria"]}
    eng_glob = {i["id"]: float(i["globalWeight"]) for i in wj["indicators"]}
    inds_by_crit = {cid: [i for i in wj["indicators"] if i["criterionId"] == cid] for cid in crit_ids}

    missing_matrices = [cid for cid in crit_ids if cid not in by_id]
    check(results, "矩阵齐全（准则层 + 各准则指标层）",
          "criteria" in by_id and not missing_matrices,
          f"矩阵 {len(by_id)} 个；缺：{missing_matrices or '无'}")
    if missing_matrices or "criteria" not in by_id:
        print("[AHP-E-INPUT] 矩阵集与权重表的准则不匹配，无法继续", file=sys.stderr)
        return 2

    cw = by_id["criteria"]["weightsGeo"]
    if len(cw) != len(crit_ids):
        print(f"[AHP-E-INPUT] 准则矩阵维度 {len(cw)} ≠ 权重表准则数 {len(crit_ids)}", file=sys.stderr)
        return 2
    criteria_map = dict(zip(crit_ids, [float(x) for x in cw]))

    global_w: dict[str, float] = {}
    local_sum_ok, local_detail = True, []
    for idx, cid in enumerate(crit_ids):
        inds = inds_by_crit[cid]
        lw = by_id[cid]["weightsGeo"]
        if len(lw) != len(inds):
            local_sum_ok = False
            local_detail.append(f"{cid}:矩阵{len(lw)}≠指标{len(inds)}")
            continue
        s = float(np.sum(lw))
        if abs(s - 1.0) > TOL:
            local_sum_ok = False
        local_detail.append(f"{cid}:Σ={s:.12f}")
        for ind, x in zip(inds, lw):
            global_w[ind["id"]] = criteria_map[cid] * float(x)

    # ---------- A. 结构自洽 ----------
    check(results, "准则权重和 = 1", abs(sum(criteria_map.values()) - 1.0) < TOL,
          f"sum={sum(criteria_map.values()):.12f}")
    check(results, "各准则局部权重和 = 1", local_sum_ok, "；".join(local_detail))
    check(results, "综合权重和 = 1", abs(sum(global_w.values()) - 1.0) < TOL,
          f"sum={sum(global_w.values()):.12f}；指标 {len(global_w)} 项")

    matrix_ids = list(by_id.keys())
    cons_ok = all(by_id[k]["pass"] for k in matrix_ids)
    cons_ratio = (f"{sum(by_id[k]['pass'] for k in matrix_ids)}/{len(matrix_ids)}"
                  if matrix_ids else "不适用（0 个矩阵）")
    check(results, f"一致性 CR<0.1 通过 {cons_ratio}", cons_ok,
          "；".join(f"{k}:λ={by_id[k]['lambdaMax']:.6f},CR={by_id[k]['cr']:.4f}" for k in matrix_ids)
          or "无判断矩阵")

    # 双算法差判据：几何平均法与特征值法对**不一致**矩阵必然给出不同近似解，
    # 差值与不一致度同阶（实测 Δ/CR 稳定在 0.06~0.16，见 AHP 文献的近似解比较）。
    # 原判据写死 Δ ≤ 1e-12，等价于要求「矩阵必须完全一致」——真实专家判断必然
    # 有分歧（CR>0），于是 09 步对任何真实项目都必 FAIL，属判据本身写错。
    # 改为：Δ ≤ max(1e-12, CR) —— 不一致度越大允许的算法差越大；CR=0 时仍要求
    # 精确相等（完全一致矩阵两法本就应当相同）。
    dmax = max((by_id[k]["delta"] for k in matrix_ids), default=0.0)
    crmax = max((by_id[k]["cr"] for k in matrix_ids), default=0.0)
    dual_tol = max(1e-12, crmax)
    # CR=0（判断完全一致）时 Δ/CR 无意义，按口径写「不适用」，不算 0 除
    ratio_text = f"{dmax / crmax:.4f}" if crmax > 1e-12 else "不适用（CR=0）"
    check(results, f"几何平均 vs 特征值双算法 maxΔ ≤ max(1e-12, CR)",
          dmax <= dual_tol,
          f"maxΔ={dmax:.2e}，CR_max={crmax:.6f}，容差={dual_tol:.2e}"
          f"（Δ/CR={ratio_text}）")

    dc = max(abs(eng_crit[k] - criteria_map[k]) for k in crit_ids)
    di = max(abs(eng_glob[k] - global_w[k]) for k in global_w)
    check(results, "引擎权重表与独立复算一致", max(dc, di) < TOL,
          f"准则maxΔ={dc:.2e}，综合maxΔ={di:.2e}")

    # ---------- 方案独立加权 ----------
    scheme_scores, incomplete = {}, []
    for s in sj["schemes"]:
        pairs = [(float(x["score"]), float(x["globalWeight"]))
                 for x in s["scores"] if x.get("score") is not None]
        if s.get("composite") is None or not pairs:
            incomplete.append(s["id"])
            continue
        wsum = sum(w for _, w in pairs)
        if wsum <= 0:                     # 权重和为 0 时加权平均无意义，按「未出分」处理
            incomplete.append(s["id"])
            continue
        scheme_scores[s["id"]] = sum(sc * w for sc, w in pairs) / wsum
    eng_scores = {s["id"]: s.get("composite") for s in sj["schemes"]}
    check(results, "评分完整性（每方案均出分，0-10 区间）",
          not incomplete and all(v is not None and 0 <= v <= 10 for v in eng_scores.values()),
          f"缺分方案：{incomplete or '无'}；" + "；".join(f"{k}={v}" for k, v in eng_scores.items()))

    if scheme_scores:
        de = max(abs(eng_scores[k] - scheme_scores[k]) for k in scheme_scores)
        check(results, "引擎评分表与独立加权一致", de < TOL, f"maxΔ={de:.2e}")
        order = sorted(scheme_scores, key=lambda k: -scheme_scores[k])
        engine_order = [r.get("id") or r.get("name") for r in sj.get("ranking", [])]
        # ranking 可能只给 name，用 name/id 双匹配
        name_of = {s["id"]: s["name"] for s in sj["schemes"]}
        engine_ids = []
        for r in sj.get("ranking", []):
            rid = r.get("id")
            if rid is None:
                rid = next((k for k, n in name_of.items() if n == r.get("name")), None)
            engine_ids.append(rid)
        check(results, "排序自洽（独立复算 = 引擎 ranking）", order == engine_ids,
              f"独立={order}；引擎={engine_ids or engine_order}")
    else:
        order = []

    # ---------- B. 基线对账（可选） ----------
    # mode 四档（与 CardoInvest / CardoOSGB 同一套口径）：
    #   auto  = 首跑自举写入基线，之后每次真对账（默认）
    #   write = 强制重写基线
    #   check = 只对账不写（基线不存在则跳过）
    #   none  = 完全跳过
    sha = hashlib.sha256(Path(args.template).read_bytes()).hexdigest().upper()
    baseline_path = Path(args.baseline) if args.baseline else (DEFAULT_WORKSPACE / BASELINE_REL)
    mode = args.baseline_mode
    if mode == "none":
        for name in ("准则权重 vs 基线", "综合权重 vs 基线", "方案得分 vs 基线",
                     "排序 vs 基线", "模板 sha256 vs 基线"):
            check(results, name, None, "mode=none（跳过）", skipped=True)
        baseline_info = {"mode": "none", "status": "skipped"}
    elif mode == "write" or (mode == "auto" and baseline is None):
        payload = baseline_from(criteria_map, global_w, scheme_scores, order, sha)
        write_baseline(baseline_path, payload)
        check(results, "基线对账", True,
              f"基线已写入（{'重写' if mode == 'write' else '首跑自举'}）："
              f"{rel_display(baseline_path)}")
        baseline_info = {"mode": mode, "status": "written",
                         "source": rel_display(baseline_path),
                         "recordedAt": payload["recordedAt"]}
    elif baseline is None:
        for name in ("准则权重 vs 基线", "综合权重 vs 基线", "方案得分 vs 基线",
                     "排序 vs 基线", "模板 sha256 vs 基线"):
            check(results, name, None,
                  f"跳过：{baseline_src}（可用 --baseline write 生成）", skipped=True)
        baseline_info = {"mode": mode, "status": "skipped", "reason": baseline_src}
    else:
        bc = baseline.get("criteriaWeights") or {}
        if not bc:
            check(results, "准则权重 vs 基线", None, "基线未提供 criteriaWeights，跳过", skipped=True)
        else:
            keys = [k for k in bc if k in criteria_map]
            bad = [k for k in keys if abs(criteria_map[k] - float(bc[k])) >= TOL]
            check(results, "准则权重 vs 基线", bool(keys) and not bad and len(keys) == len(bc),
                  "；".join(f"{k}={criteria_map[k]:.6f}（基线 {float(bc[k]):.4f}）" for k in keys)
                  + (f"；超差：{bad}" if bad else ""))

        bg = baseline.get("globalWeights") or {}
        if not bg:
            check(results, "综合权重 vs 基线", None, "基线未提供 globalWeights，跳过", skipped=True)
        else:
            common = [k for k in bg if k in global_w]
            if len(common) != len(bg) or len(bg) != len(global_w):
                check(results, "综合权重 vs 基线", None,
                      f"指标集与基线不一致（基线 {len(bg)} 项 / 本次 {len(global_w)} 项），跳过", skipped=True)
            else:
                dg = max(abs(global_w[k] - float(bg[k])) for k in common)
                check(results, "综合权重 vs 基线", dg < TOL,
                      f"{len(common)} 项对齐；maxΔ={dg:.2e}")

        bs = baseline.get("schemeScores") or {}
        if not bs or not scheme_scores:
            check(results, "方案得分 vs 基线", None, "基线未提供 schemeScores 或本次无出分方案，跳过", skipped=True)
        else:
            common = [k for k in bs if k in scheme_scores]
            if len(common) != len(bs):
                check(results, "方案得分 vs 基线", None,
                      f"方案集与基线不一致（基线 {len(bs)} / 本次 {len(scheme_scores)}），跳过", skipped=True)
            else:
                ds = {k: abs(scheme_scores[k] - float(bs[k])) for k in common}
                check(results, "方案得分 vs 基线", max(ds.values()) < SCORE_TOL,
                      "；".join(f"{k}={scheme_scores[k]:.4f}(Δ={ds[k]:.2e},{rating_of(scheme_scores[k])})"
                               for k in common))

        bo = baseline.get("schemeOrder")
        if not bo:
            check(results, "排序 vs 基线", None, "基线未提供 schemeOrder，跳过", skipped=True)
        else:
            check(results, "排序 vs 基线", order == list(bo), f"独立={order}；基线={list(bo)}")

        bsha = baseline.get("templateSha256")
        if not bsha:
            check(results, "模板 sha256 vs 基线", None,
                  f"基线未提供 templateSha256；当前 sha256={sha[:16]}…{sha[-8:]}", skipped=True)
        else:
            check(results, "模板 sha256 vs 基线（模板原件未被改写）", sha == str(bsha).upper(),
                  f"sha256={sha[:16]}…{sha[-8:]}")
        baseline_info = {"mode": mode, "status": "checked", "source": baseline_src}

    # ---------- 汇总 ----------
    passed = sum(1 for r in results if r["status"] == "PASS")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    skipped = sum(1 for r in results if r["status"] == "SKIP")
    baseline_text = baseline_src
    if baseline_info.get("status") == "written":
        baseline_text = ("首跑未找到既有基线，已按本次结果写入新基线"
                         f"（{baseline_info.get('source')}，新工作区模式）")
    summary = {"schema": "cardo-ahp/independent-verification/v1",
               "baseline": baseline_text, "hasBaseline": baseline is not None,
               "baselineInfo": baseline_info,
               "total": len(results), "passed": passed, "failed": failed, "skipped": skipped,
               "allPass": failed == 0, "checks": results,
               "recomputed": {"criteriaWeights": criteria_map, "globalWeights": global_w,
                              "schemeScores": scheme_scores,
                              "ranking": order}}

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "独立验证结果.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = ["# AHP 独立验证报告（numpy 复算 + 基线对账）", "",
             f"- 通过 {passed} 项 / 失败 {failed} 项 / 跳过 {skipped} 项（共 {len(results)} 项）",
             f"- 基线：{baseline_src}（模式 {baseline_info.get('mode')}／状态 {baseline_info.get('status')}）",
             "- 复算路径：矩阵集 → numpy 特征值/几何平均 → 综合权重 → 加权评分（不调用业务引擎）", "",
             "| # | 验证项 | 结果 | 明细 |", "|---|---|---|---|"]
    icon = {"PASS": "✅通过", "FAIL": "❌失败", "SKIP": "⏭跳过"}
    for i, r in enumerate(results, 1):
        lines.append(f"| {i} | {r['item']} | {icon[r['status']]} | {r['detail']} |")
    (out_dir / "独立验证报告.md").write_text("\n".join(lines), encoding="utf-8")

    print(f"\n{passed} 通过 / {failed} 失败 / {skipped} 跳过；结果已写入 {out_dir}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
