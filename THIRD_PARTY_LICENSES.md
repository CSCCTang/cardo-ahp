# 第三方依赖与许可证清单（cardo-ahp）

只列本包 `requirements.txt` 实际声明的依赖；未在此列的组件不由本包引入。

| 依赖 | 用途 | 许可证 |
|---|---|---|
| python-docx | 生成 docx 成果 | MIT |
| openpyxl | xlsx 读写与样式 | MIT |
| numpy | 数值计算（特征向量、一致性） | BSD-3-Clause |

## 说明

- 本包**不依赖 PyMuPDF**；PyMuPDF 只出现在 `cardo-status`。
- 本包不包含任何第三方二进制、wheel 或 Python 运行时；依赖由使用者从 PyPI 安装。
- 四个包统一以 AGPL-3.0-or-later 发布，保持同一套流水线许可证一致；本包所依赖的上表组件均为宽松许可，
  如需闭源分发，可自行将本包改为 MIT 并保留上表声明。
