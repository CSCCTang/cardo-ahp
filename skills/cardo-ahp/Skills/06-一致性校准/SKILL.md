---
name: cardo-ahp-06-一致性校准
description: AHP 流程第 6 步（条件分支）。仅当 05 步某矩阵 CR≥0.1 时使用：诊断最不一致元素，给出 ≤3 个最小改动候选（改前/改后对比），不自动采用；经人工批准后 apply 并回 05 重算。全部矩阵 CR<0.1 时本步只做诊断留痕，结论为「无需校准」。
---

# 06 · 一致性校准（AHP 流程第 6 步 · 条件分支）

## 何时使用
- 进入条件：第 05 步存在 CR≥0.1 的矩阵。
- 全部矩阵 CR<0.1 时（真实项目的常态）：本步只做诊断留痕，结论「无需校准」，随后直接进入 07。

## 输入 / 输出
- 输入：`过程文件/05-权重与一致性/全矩阵/{criteria,current,economy,compliance,risk}.json`
- 输出（过程归档）：`过程文件/06-一致性校准/诊断-*.json`、`候选-*.json`、`校准记录.md`

## 执行命令（在 CardoAHP 根目录执行；对 5 个群体矩阵逐一诊断）

```powershell
$enc = New-Object System.Text.UTF8Encoding($false)
foreach ($m in @("criteria","current","economy","compliance","risk")) {
  $d = & python Skills/06-一致性校准/scripts/ahp_calibrate.py diagnose `
    --matrix "过程文件/05-权重与一致性/全矩阵/$m.json" --format json 2>&1 | Out-String
  [IO.File]::WriteAllText((Join-Path $PWD "过程文件/06-一致性校准/诊断-$m.json"), $d, $enc)
  # 注意：CR 已通过时 suggest 不写 --out 文件，直接把 stdout 归档为候选结论
  $c = & python Skills/06-一致性校准/scripts/ahp_calibrate.py suggest `
    --matrix "过程文件/05-权重与一致性/全矩阵/$m.json" --max-candidates 3 --no-timestamp 2>&1 | Out-String
  [IO.File]::WriteAllText((Join-Path $PWD "过程文件/06-一致性校准/候选-$m.json"), $c, $enc)
}
```

（PS5 的 `>` 会写成 UTF-16，故统一用 .NET UTF-8 无 BOM 落盘，详见根 README。）

> 退出码语义：`diagnose` CR 通过时退出码 0；`suggest` 在 **CR 已通过（无候选）时退出码=1**，stdout 仍为有效 cardo/ahp-calibration/v1 结论（`candidates: []`），属预期分支，不要当失败拦截。只有 CR≥0.1 时才会产出候选并继续 apply 流程。

## 若真的出现 CR≥0.1（后续项目流程，需人工介入）
1. diagnose 定位最不一致元素/行；suggest 给 ≤3 个最小改动候选（改前/改后、CR 改善幅度）。
2. 候选必须经人工研判选定并留批准人：
   `ahp_calibrate.py apply --matrix <原矩阵.json> --candidate <候选.json> --out <校准后.json> --approve-by "<姓名>"`
3. 用校准后矩阵回到第 05 步重算权重与 CR；批准记录写入校准记录.md。

## 红线
- 禁止为了 CR 过线做最小自动修正后直接采用；禁止删除、静默覆盖原矩阵；校准只作用于群体判断矩阵，不改问卷原件。
