---
name: cardo-ahp-04-多专家聚合
description: AHP 流程第 4 步。多专家成对比较结果逐元素几何平均（Saaty 推荐口径）成 5 个群体判断矩阵，并同时给出 weight-then-average 对比口径、专家一致性与离散度提示；专家等权、CR 不达标不静默剔除。
---

# 04 · 多专家聚合（AHP 流程第 4 步）

## 何时使用
问卷全部校验合格后。把 5–7 位专家的判断聚合成群体判断矩阵，作为权重计算的唯一输入。

## 输入 / 输出
- 输入：`输入文件/02-专家问卷（已回收）/`（由 03 步确认全部合格）
- 输出（过程归档）：
  - `过程文件/04-多专家聚合/群体矩阵.json`（cardo/ahp-aggregated/v1）
  - `过程文件/04-多专家聚合/聚合报告.md`

## 执行命令（在 CardoAHP 根目录执行）

```powershell
python Skills/04-多专家聚合/scripts/ahp_aggregate.py aggregate `
  --dir "输入文件/02-专家问卷（已回收）" `
  --out "过程文件/04-多专家聚合/群体矩阵.json" `
  --compare --no-timestamp

python Skills/04-多专家聚合/scripts/ahp_aggregate.py aggregate `
  --dir "输入文件/02-专家问卷（已回收）" `
  --format md --no-timestamp > 过程文件/04-多专家聚合/聚合报告.md
```

退出码：0 聚合一致性通过；1 有矩阵 CR≥0.1；3 专家少于 5 人。

## 口径（固定，不静默切换）
- 主口径：逐元素几何平均后再求权重（保留互反性，Saaty 推荐）。
- 对比口径：逐专家求权重后算术平均（`--compare`），只作对比写进 weightsCompare，绝不替代主口径。
- 专家等权；CR 不达标专家只在报告中列出，不自动剔除（本工作区 5 份 CR 全为 0）。

## 闸门
- aggregatePass=true；专家数 ≥5；群体矩阵 5 个且互反。
- 离散度提示（range 过大）只建议复评，不自动改数。

## 调优点
- 聚合规则、离散度阈值、RI 表：改 `规则/标度与口径.yaml` 与 `scripts/ahp_aggregate.py`。
