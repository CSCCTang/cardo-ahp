---
name: cardo-ahp-01-指标体系
description: AHP 流程第 1 步。从本 skill 的 规则/指标体系.yaml 读取「片区城市更新改造方案 AHP 四准则 16 指标」，打印指标目录供全员确认。所有后续问卷、聚合、权重、打分都以本步骤输出的 id/顺序为唯一锚点。
---

# 01 · 指标体系确认（AHP 流程第 1 步）

## 何时使用
每个 AHP 课题开工第一步；指标增删、改名、方向调整后也必须重跑，后续步骤全部重做。

## 输入 / 输出（相对 CardoAHP 根目录）
- 输入（唯一真相来源）：[规则/指标体系.yaml](规则/指标体系.yaml)（4 准则 × 4 指标；附 1/3/5/7/9 锚点、方向）
- 输出（过程归档）：
  - `过程文件/01-指标体系/指标目录.json`
  - `过程文件/01-指标体系/指标目录.md`

## 执行命令（在 CardoAHP 根目录执行）

```powershell
python Skills/01-指标体系/scripts/ahp_calc.py indicators --format json > 过程文件/01-指标体系/指标目录.json
python Skills/01-指标体系/scripts/ahp_calc.py indicators --format md   > 过程文件/01-指标体系/指标目录.md
```

## 闸门（不满足不得进入 02）
- 恰好 4 个准则：现状情况 current / 投资经济 economy / 规划合规 compliance / 风险可控 risk。
- 恰好 16 个末级指标，id 形如 `current.age`，顺序与《4 层次分析法AHP四准则模板.xlsx》§四/§六一致。
- 每个指标具备 definition / direction / anchors（1/3/5/7/9）。

## 调优点（优化成果改这里）
- 指标体系要改：只改 `规则/指标体系.yaml`，然后 01→09 全链重跑（问卷不可复用）。
- 口径（1-9/0.1-1 双标度、CR<0.1、RI 表）在 `规则/标度与口径.yaml`。
