# cardo-ahp

> 安装（二选一）：
> `dsh plugin --profile <profile> add github:CSCCTang/cardo-ahp#v0.1.2`
> `dsh plugin --profile <profile> add <本地路径>\cardo-ahp-0.1.2.tgz`

AHP 投资决策量化：给一批专家问卷，产出指标权重、一致性检验、方案评分与依据对照表。

## 需要提供什么

| 输入 | 形态 | 最小样例 |
|---|---|---|
| 专家问卷 / 打分表 | xlsx，放工作区 `输入文件\` 任意层级 | `skills\cardo-ahp\assets\inputs\专家名单-空白.xlsx` |
| 候选方案清单 | xlsx（可与问卷同表） | 与问卷同表的方案页 |

## 装完怎么调

在 dsh 对话里说「用 cardo-ahp 跑，工作区是 `D:\proj\ahp-01`」。

工作区是独立空文件夹：`输入文件\` 放料、`过程文件\` 留痕、`成果文件\` 出活。

## 会得到什么

`成果文件\`：

- AHP 权重表（xlsx）
- 一致性检验结果
- 方案评分结论（docx）
- 依据对照表

## 许可证与版本

- 版本 `0.1.2`；许可证 **AGPL-3.0-or-later**，见 [LICENSE](LICENSE) 与 [NOTICE](NOTICE)。
- 第三方组件与许可证见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。
- 运行环境：Windows；Python ≥ 3.11；dsh ≥ 0.2.2-rc.5。

## 文档

- [docs/01_使用流程.md](docs/01_使用流程.md)
- [docs/02_输入准备.md](docs/02_输入准备.md)
- [docs/03_常见失败.md](docs/03_常见失败.md)
