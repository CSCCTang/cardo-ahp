#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""聚合转矩阵集.py -- 把 cardo/ahp-aggregated/v1 的群体矩阵转成 ahp_calc 需要的矩阵集。

为什么需要它：
  ahp_aggregate.py 产出的群体矩阵是「全矩阵 full」形态（n×n 互反矩阵）；
  ahp_calc.py weights/consistency 吃的矩阵集 {matrices:[{id,level,order,values}]}，
  values 支持 full（给全部元素）与 column（列向量口径）两种形态。

**口径必须用 full，不能用「首列还原成列向量」**：
  群体矩阵是 n×n 互反矩阵，首列只是「各项相对第 1 项的比值」，**不足以还原整张矩阵**——
  首列还原 a_ij = v_i/v_j 只在「矩阵完全一致（传递性成立）」时才等于原矩阵。
  实测（5 位专家真实判断）：首列还原后矩阵元素最大偏差 0.26、准则层权重最大偏差 0.037
  （3.7 个百分点），且首列元素常 >1，触发 scale_0_1 输入域 [0.1,1] 越界报错、整链中断。
  故群体矩阵一律以 full 形态写出：逐格做正性/互反/对角线校验后原样参与计算，
  05 步权重与 04 步聚合报告口径完全一致（权重不再漂移）。

同时为 06-一致性校准 导出每个群体矩阵的 full 形态单矩阵 JSON。

用法：
  python 聚合转矩阵集.py --agg <群体矩阵.json> --out-dir <输出目录>
退出码：0 成功；2 输入/参数错。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

CRITERIA_ID = "criteria"
INDICATOR_SCALE = "scale_0_1"
CRITERIA_SCALE = "scale_1_9"


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="聚合结果（群体全矩阵）→ 列向量矩阵集 + 单矩阵 full JSON")
    ap.add_argument("--agg", required=True, help="ahp_aggregate 产出的 cardo/ahp-aggregated/v1 JSON")
    ap.add_argument("--out-dir", required=True, help="输出目录（矩阵集.json、全矩阵/*.json）")
    args = ap.parse_args()

    agg_path = Path(args.agg)
    if not agg_path.is_file():
        print(f"[AHP-E-INPUT] 聚合文件不存在：{args.agg}", file=sys.stderr)
        return 2
    try:
        agg = json.loads(agg_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"[AHP-E-INPUT] 聚合文件不是合法 JSON：{exc}", file=sys.stderr)
        return 2

    group = agg.get("matrices")
    if not isinstance(group, dict) or not group:
        print("[AHP-E-INPUT] 聚合文件缺 matrices{} 群体矩阵段", file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir)
    full_dir = out_dir / "全矩阵"
    full_dir.mkdir(parents=True, exist_ok=True)

    # 固定顺序：criteria 在前，四个指标层矩阵按指标体系顺序
    order_preference = ["criteria", "current", "economy", "compliance", "risk"]
    ids = [i for i in order_preference if i in group] + [i for i in group if i not in order_preference]

    matrix_set = {"matrices": []}
    for mid in ids:
        m = group[mid]
        matrix = m.get("matrix")
        n = int(m.get("n") or 0)
        if not isinstance(matrix, list) or n == 0 or len(matrix) != n:
            print(f"[AHP-E-INPUT] 群体矩阵 {mid} 结构异常（缺 n×n matrix）", file=sys.stderr)
            return 2
        level = "criteria" if mid == CRITERIA_ID else "indicators"
        scale = CRITERIA_SCALE if level == "criteria" else INDICATOR_SCALE

        # full 形态：原样写出群体矩阵（逐格由 ahp_calc 再校验正性/互反/对角线）
        full_values = [[float(x) for x in row] for row in matrix]
        matrix_set["matrices"].append({
            "id": mid,
            "level": level,
            "scale": scale,
            "mode": "full",
            "order": n,
            "values": full_values,
        })

        # 校准用 full 单矩阵
        single = {
            "schema": "cardo/ahp-judgment/v1",
            "id": mid,
            "level": level,
            "scale": scale,
            "mode": "full",
            "order": n,
            "values": full_values,
            "note": "群体判断矩阵（各专家逐元素几何平均），供一致性校准诊断使用",
        }
        (full_dir / f"{mid}.json").write_text(
            json.dumps(single, ensure_ascii=False, indent=2), encoding="utf-8")

    set_path = out_dir / "矩阵集.json"
    set_path.write_text(json.dumps(matrix_set, ensure_ascii=False, indent=2), encoding="utf-8")

    # 互反性自检（群体矩阵必须仍是互反矩阵）
    worst = 0.0
    for mid in ids:
        mat = group[mid]["matrix"]
        for i in range(len(mat)):
            for j in range(len(mat)):
                worst = max(worst, abs(mat[i][j] * mat[j][i] - 1.0))
    print(f"已写出：{set_path}")
    print(f"已写出单矩阵 full JSON：{full_dir}（{len(ids)} 个）")
    print(f"群体矩阵互反性最大偏差：{worst:.3e}")
    if worst > 1e-9:
        print("[AHP-E-NONRECIPROCAL] 群体矩阵互反性超差，停止", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
