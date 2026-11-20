---
name: cardo-ahp-03-回收校验
description: AHP 流程第 3 步。对已回收专家问卷逐份做结构与口径校验：标度范围、互反、缺失/非法值、矩阵完整性。不合格问卷退回专家修改，不进入聚合。
---

# 03 · 问卷回收校验（AHP 流程第 3 步）

## 何时使用
专家问卷回收后、聚合前。逐份校验，不合格的退回修改，合格的才允许进入第 04 步。

## 输入 / 输出
- 输入：`输入文件/02-专家问卷（已回收）/AHP专家问卷-专家*.xlsx`
- 输出（过程归档）：
  - `过程文件/03-回收校验/回收校验报告.json`
  - `过程文件/03-回收校验/回收校验报告.md`

## 执行命令（在 CardoAHP 根目录执行）

```powershell
python Skills/03-回收校验/scripts/gen_questionnaire.py check `
  --dir "输入文件/02-专家问卷（已回收）" `
  --scale scale_0_1 `
  --criteria-scale scale_1_9 `
  --format json --no-timestamp > 过程文件/03-回收校验/回收校验报告.json
python Skills/03-回收校验/scripts/gen_questionnaire.py check `
  --dir "输入文件/02-专家问卷（已回收）" `
  --scale scale_0_1 `
  --criteria-scale scale_1_9 `
  --format md   --no-timestamp > 过程文件/03-回收校验/回收校验报告.md
```

## 闸门
- 5 份问卷全部 valid：4 个 sheet、30 个上三角输入格齐全；数值落在标度区间；下三角互反。
- 任一份 invalid：退回专家修改后重新校验；**不得**在本步擅自替专家改数。
- 注意：xlsx 公式单元格（下三角）由脚本按上三角重算，不依赖 Excel 是否已打开计算。

## 调优点
- 校验严格度（范围、容差、必填数）：改 `scripts/_survey_lib.py`。
- 校验规则与计算口径分离：标度阈值在 `规则/标度与口径.yaml`。
