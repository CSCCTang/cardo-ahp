#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cardo AHP 投资决策量化 · 一条命令开工。

依次做四件事：
  依赖自检（缺就在包根建 .venv，按 阿里云 → 清华 → USTC → 官方源 回退安装）→ 建目录与释放模板
  → 校验四准则模板 / 回收问卷 / 工作区参数.json → 调 01→09 流水线。

流程有两处天然停顿：--to 02 发卷 → 回收问卷；08 首次生成空白评分表 → 填表后 --from 08。
退出码：0 全过；3 缺料/参数不全；4 设计好的拒收（如演示数据未声明 --demo）；1 流水线硬失败（含 08 首次等填表）；2 环境或结构问题。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import traceback

# 运行期不写 __pycache__：Skills/ 是只读交付物，写进 pyc 会污染交付库
# （实测交付库里混进过 19 个 pyc）。这里在导入任何本地模块之前就关掉。
sys.dont_write_bytecode = True

from datetime import datetime
from pathlib import Path

SKILL = Path(__file__).resolve().parent      # 技能包：只读，运行期不写入
RUNTIME_PY = SKILL / "runtime" / "python" / ("python.exe" if os.name == "nt" else "python")
ROOT = SKILL.parent
CACHE_BASE = Path(os.environ.get("CARDO_CACHE_HOME") or os.environ.get("LOCALAPPDATA") or str(Path.home()))
def venv_dir() -> Path:
    """派生仓库：虚拟环境放用户缓存，不写安装目录。"""
    return CACHE_BASE / "cardo" / "venv-ahp"
                          # 工作区：技能包所在的上一级目录
# 工作区是否已成功解析（resolve_workspace 会抛 SystemExit；未解析时不能拿默认 ROOT 去写看板状态）
WORKSPACE_RESOLVED = False
TITLE = "Cardo AHP 投资决策量化"
PY_MIN = (3, 9)

TEMPLATE_DST = "输入文件/01-AHP四准则模板/4 层次分析法AHP四准则模板.xlsx"
PENDING_DIR = "输入文件/02-专家问卷（待发放）"   # 领料区：空白问卷，人可见
RECV_DIR = "输入文件/02-专家问卷（已回收）"     # 回收区：专家填好的问卷
SCORE_DIR = "输入文件/03-方案打分表"            # 评分表区
SCHEME_LIST_REL = SCORE_DIR + "/方案清单.xlsx"  # 候选方案登记表（人填，方案名以此为准）
DIRS = [
    "输入文件",
    "输入文件/01-AHP四准则模板",
    PENDING_DIR,
    RECV_DIR,
    SCORE_DIR,
    "过程文件",
    "成果文件",
    "记忆与规则",
]
ASSETS = [
    {
        "src": "assets/templates/4 层次分析法AHP四准则模板.xlsx",
        "dst": TEMPLATE_DST,
        "sha256": "1C622290CD79ED63C5626E5808813C35E4F026A9127483022F98671AC574D03F",
        "label": "AHP 四准则模板",
    },
]
INPUT_TEMPLATES = [
    {"src": "assets/inputs/专家名单-空白.xlsx", "dst": RECV_DIR + "/专家名单.xlsx"},
]
IMPORTS = ["docx", "openpyxl", "numpy"]

PARAMS = "工作区参数.json"
SCHEMA = "cardo-ahp/workspace/v1"
DEFAULTS = {"experts": 5, "baselineMode": "auto"}
REQUIRED = ["projectName", "experts"]
MIN_QUESTIONNAIRES = 5

# 同一份问卷被复制成多份 → 会被当成多位独立专家，聚合权重其实只含 1 位专家的判断。
# 这是静默失效（份数够、CR 也可能过），比缺料更值得防，所以单独给错误码。
DUP_EXPERT_CODE = "AHP-E-DUP-EXPERTS"
DUP_REPORT_HEADING = "## 查重结果（同一份问卷被复制成多份）"
SURVEY_NAME_PATTERN = "AHP专家问卷-专家*.xlsx"

# 看板顶栏要显示「工作区真实进度」，但看板是纯前端单文件、读不到工作区状态，
# 所以由 开工.py 把状态注入看板里的一个 script 块（看板读不到就优雅降级）。
BOARD_STATUS_BEGIN = "<!-- CARDO-STATUS:BEGIN -->"
BOARD_STATUS_END = "<!-- CARDO-STATUS:END -->"

PIPELINE = "Skills/00-Cardo城市更新AHP层次分析法/scripts/pipeline.py"
DELIVERABLES = [
    "成果文件/成品版/AHP权重表/AHP权重表_成品版.xlsx",
    "成果文件/成品版/AHP方案评分结论/AHP方案评分结论_成品版.xlsx",
    "成果文件/成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.docx",
    "成果文件/成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.pdf",
    "成果文件/内部件/AHP依据对照表/AHP依据对照表.xlsx",
]
GUIDES = {
    "02": "发卷：把 领料区（输入文件/02-专家问卷（待发放））里的空白问卷发给专家（5–7 位），"
          "专家填完原样放回 输入文件/02-专家问卷（已回收）/，文件名不要改。",
    "03": "回收校验：读 过程文件/03-回收校验/回收校验报告.md，逐份修正问卷结构/标度或补收，至少 5 份且无 CR 警告，再 --from 03。",
    "06": "一致性校准：读 过程文件/06-一致性校准/校准记录.md，CR≥0.1 的矩阵需按候选人工批准，处理完 --from 05 重算。",
    "08": "方案打分：08 首次运行会生成空白评分表并停下。填 输入文件/03-方案打分表/AHP方案评分表.xlsx ——"
          "「方案信息」登记编号与名称，「方案评分表」黄色格填 0–10 分并写依据（指标ID/权重/标准不要改），再 --from 08。",
}

# ---------------------------------------------------------------- 收料引导
INTAKE_ITEMS = [
    {
        "key": "questionnaires",
        "name": "回收问卷 AHP专家问卷-专家N.xlsx（≥5 份）",
        "required": True,
        "dst": RECV_DIR,
        "patterns": ["AHP专家问卷-专家*.xlsx"],
        "how": "第 02 步发出的空白问卷，由专家填完原样交回；文件名保持「AHP专家问卷-专家N.xlsx」，"
               "放回收区（输入文件/02-专家问卷（已回收））",
        "why": "专家判断是权重计算的唯一来源；少于 5 份不做聚合",
        "empty_hint": "还没有？先跑「{run}」生成空白问卷（领料区：{pending}），发给专家填完再收回来。",
    },
    {
        "key": "schemes",
        "name": "候选方案清单（方案名以你填的为准）",
        "required": False,
        "dst": SCORE_DIR,
        "patterns": ["方案清单.xlsx"],
        "how": "首次运行自动生成，已预填城市更新常见三方案作建议起点；"
               "把方案名改成你项目的真实方案，可增删行，不比的方案标「否」",
        "why": "第 08 步按它预填评分表的方案列 —— 方案名在这里定，评分表直接对号入座",
    },
    {
        "key": "scoresheet",
        "name": "方案评分表 AHP方案评分表.xlsx",
        "required": False,
        "dst": SCORE_DIR,
        "patterns": ["AHP方案评分表.xlsx"],
        "how": "跑到第 08 步会自动生成空白表并停下；填好后放回同一位置再 --from 08",
        "why": "第 09 步方案评分结论的输入",
    },
    {
        "key": "others",
        "name": "其他材料（有什么放什么）",
        "required": False,
        "catch_all": True,
        "dst": "输入文件/",
        "patterns": ["*"],
        "how": "统计年鉴、上级规划、新闻剪报、访谈录音转写、现场台账、CardoStatus 跑出来的"
               "《现状分析报告》等，凡是你觉得用得上的都能放，文件夹和 zip 也行",
        "why": "01 会自动归类；归不了的会问你怎么归。这些材料会作为事实来源和附件参考",
    },
]


def _files_under(rel_dir: str) -> list[Path]:
    """指定目录下的文件（递归）。相对工作区根。"""
    base = ROOT / rel_dir
    if not base.is_dir():
        return []
    return [p for p in base.rglob("*")
            if p.is_file() and not p.name.startswith("~$") and "空白" not in p.name]


# 系统自动释放的文件不算「用户提供的其他材料」——它们每次运行都会出现，
# 列进「其他材料」会让新手以为自己也放过了什么。
AUTO_RELEASED = {
    "输入文件/01-AHP四准则模板/4 层次分析法AHP四准则模板.xlsx",
    RECV_DIR + "/专家名单.xlsx",
}


def intake_scan():
    """逐项探测落位情况；返回 [(item, [命中文件])]。

    判定按**专属目录**收敛，与 precheck/流水线的实际读取口径一致：
      · 回收问卷只在 回收区 认 —— 领料区的空白卷不算「已回收」；
      · 方案评分表只认 评分表区 里那个固定文件名的表；
      · catch_all（「其他材料」）统计 输入文件/ 下没被前面认领、
        且不是系统自动释放的文件。
    """
    import fnmatch

    rows, claimed = [], set()
    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            continue
        hits = [p for p in _files_under(item["dst"])
                if any(fnmatch.fnmatch(p.name.lower(), pat.lower()) for pat in item["patterns"])]
        claimed.update(hits)
        rows.append((item, sorted(set(hits))))

    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            others = []
            for p in _files_under(item["dst"]):
                if p in claimed:
                    continue
                rel = norm(p.relative_to(ROOT))
                if rel in AUTO_RELEASED:
                    continue
                others.append(p)
            rows.append((item, sorted(others)))

    order = {id(item): index for index, item in enumerate(INTAKE_ITEMS)}
    return sorted(rows, key=lambda pair: order[id(pair[0])])


def write_intake_list(rows, params: dict | None = None) -> Path:
    """生成人读的《放料清单.md》：**一步步照着做**的向导。

    不是「缺什么材料」的清单，而是「现在到第几步、这一步干什么、干完长什么样、
    怎么确认做对了」的路线图 —— 状态由 workflow.py 读文件实际内容判定，不是猜的。
    """
    target = ROOT / "放料清单.md"
    ready = sum(1 for item, hits in rows if hits)
    params = params or {}

    steps = []
    STATUS_MARK = {"done": "[完成]", "doing": "[进行中]", "todo": "[未开始]"}
    STATUS_TEXT = {"done": "已完成", "doing": "进行中", "todo": "未开始"}
    try:
        import sys as _sys
        _p = str(SKILL / "Skills/08-方案打分/scripts")
        if _p not in _sys.path:
            _sys.path.insert(0, _p)
        from workflow import build as _wf_build
        steps = _wf_build(ROOT)
        from workflow import STATUS_MARK as _M, STATUS_TEXT as _T
        STATUS_MARK, STATUS_TEXT = _M, _T
    except Exception:
        pass

    lines = ["# 放料清单 · {}".format(SKILL.name), "",
             "> 本文件由 开工.py 自动生成，**每次运行都会刷新** —— 你做完一步就重跑一次，状态会更新。",
             "> 工作区：`{}`".format(norm(ROOT)), ""]

    # ---------------- 全流程一览 ----------------
    lines += ["## 全流程一览", ""]
    if steps:
        lines += ["| 步骤 | 谁做 | 状态 | 进展 |", "|---|---|---|---|"]
        for s in steps:
            lines.append("| {} {} | {} | {} {} | {} |".format(
                s["no"], s["title"], s["who"], STATUS_MARK[s["state"]],
                STATUS_TEXT[s["state"]], s.get("progress", "")))
        lines += ["", "流程：" + " → ".join(
            "{} {}".format(s["no"], s["title"]) for s in steps), ""]

        # 当前该做的那一步
        cur = next((s for s in steps if s["state"] == "doing"), None)
        if cur is None:
            cur = next((s for s in steps if s["state"] == "todo"), None)
        if cur:
            lines += ["## 现在做这一步：{} {}".format(cur["no"], cur["title"]), "",
                      "**做什么**", "", cur["what"], ""]
            for d in cur["detail"]:
                lines.append("- " + d)
            lines += ["", "**在哪**（双击打开）", "", "```text", cur["where"], "```", "",
                      "**做完长什么样**", "", "- " + cur["done_when"], "",
                      "**怎么确认做对了**", "", "- " + cur["how_check"], "",
                      "> 详细到每一格怎么填：见工作区根目录 **《怎么填.md》**。", ""]

        # 全部步骤的详细卡
        lines += ["## 每一步的详细做法", ""]
        for s in steps:
            lines += ["### {} {} —— {}".format(s["no"], s["title"], STATUS_TEXT[s["state"]]), "",
                      "- **谁做**：{}".format(s["who"]),
                      "- **在哪**：`{}`".format(s["where"]),
                      "- **做什么**：{}".format(s["what"])]
            for d in s["detail"]:
                lines.append("  - " + d)
            lines += ["- **做完长什么样**：{}".format(s["done_when"]),
                      "- **怎么确认做对了**：{}".format(s["how_check"]), ""]

    # ---------------- 本次替你定的值 ----------------
    sug = []
    if params.get("projectName"):
        src = params.get("projectNameSource")
        if src == "sheet":
            mark = "（你在《方案清单.xlsx》B1 格填的）"
        elif src == "user":
            mark = "（你用 --project 指定的）"
        else:
            mark = "（你指定的）"
        sug.append(("项目名", params["projectName"], mark))
    else:
        sug.append(("项目名", "**待填**",
                    "（在《方案清单.xlsx》的 B1 格填；会印在三份成果文件上。"
                    "也可以用 --project \"正式全称\" 指定）"))
    if params.get("experts"):
        src = params.get("expertsSource")
        mark = ("（建议：{}）".format(PROJECT_NAME_HINT["expert"])
                if src == "suggest" else "（你指定的）")
        sug.append(("专家人数", "{} 位".format(params["experts"]), mark))
    if (ROOT / SCHEME_LIST_REL).is_file():
        sug.append(("候选方案", "见《方案清单.xlsx》",
                    "（空白待填，你自己定；第 08 步按它生成评分表。"
                    "想先看常见方案可以加 --preset-schemes 重跑）"))
    if sug:
        lines += ["## 本次替你定的值（都能改）", "",
                  "| 项 | 取值 | 说明 |", "|---|---|---|"]
        for name, val, note in sug:
            lines.append("| {} | {} | {} |".format(name, val, note))
        lines += [""]

    # ---------------- 材料到位情况 ----------------
    lines += ["## 材料到位情况", "",
              "| 项 | 是否必需 | 状态 | 落位（绝对路径） |", "|---|---|---|---|"]
    for item, hits in rows:
        if hits:
            where = "、".join(norm(path.relative_to(ROOT)) if ROOT in path.parents else norm(path)
                              for path in hits[:2])
            status = "已到位（{}）".format(where)
        else:
            status = "**缺**" if item["required"] else "未提供（可选）"
        lines.append("| {} | {} | {} | `{}` |".format(
            item["name"], "必需" if item["required"] else "可选", status,
            norm(ROOT / item["dst"])))
    lines += ["", "已到位 {} / {} 项。".format(ready, len(rows)), ""]

    # ---------------- 常用命令 ----------------
    lines += ["## 常用命令", "",
              "```text",
              "{}          # 生成空白问卷（发卷前）".format(cmd_hint("--to 02")),
              "{}          # 收齐问卷后算权重与一致性".format(cmd_hint("--to 07")),
              "{}        # 填完评分表后出成果".format(cmd_hint("--from 08")),
              "{}         # 一项一项问，确认材料落位".format(cmd_hint("--intake")),
              "```", "",
              "> 已装 Python 3.9+ 的机器，把开头的解释器换成 `python` 也一样：",
              "> `python {}\\开工.py --to 02`".format(SKILL.name),
              "> 详细的填写位置与口径见 **《怎么填.md》**（工作区根目录）。", ""]
    target.write_text("\n".join(lines), encoding="utf-8")
    return target
HOWTO_MD = """# 怎么填 · AHP 专家问卷与方案评分表

> 本文件由 开工.py 自动生成，说明**要人填的表长什么样、填哪一格、填完放哪**。
> 工作区：`{ws}`

## 这条流程怎么走（默认你是空白文件夹）

```
① 定方案      你      方案清单.xlsx      B1 项目名 + 方案名
② 发问卷      你→专家  待发放目录        空白卷转发出去
③ 回收问卷    专家→你  已回收目录        收齐 ≥5 份，跑 --to 07
④ 打分        你      方案评分表.xlsx    黄色格填 0–10 分 + 依据
⑤ 出成果      系统    成果文件/          权重表 + 评分结论 + 结论文档 + 依据对照表
```

**你要动手的只有三处**：方案清单（你填）、专家问卷（专家填）、方案评分表（你填）。
其余全是系统的事。

> **每次跑 `开工.py`，工作区根目录的《放料清单.md》都会刷新** ——
> 里面有五步的实时状态（✅已完成 / ▶进行中 / ○未开始），
> 以及「现在做这一步」—— 做什么、在哪（绝对路径）、做完长什么样、怎么确认做对了。
> 做完一步就重跑一次，状态会跟着走。

---

## 先分清：哪些是你定的，哪些是固定的

| | 谁定 | 能不能改 |
|---|---|---|
| **比哪几个方案** | **你** | 自由。数量不限（一般 2~4 个），名字自定 |
| **每个方案打多少分** | **你** | 0–10 分 + 依据 |
| 打分项（16 项指标） | 系统（按项目方向规则） | 不能随意改 —— 改了权重就作废、专家问卷要重发 |
| 综合权重 | 系统算 | 不能改 |

> 为什么打分项不能留空：专家问卷就是按这 16 项打的，权重从这 16 项算出来，
> 评分表再用这 16 项给每个方案打分。改了打分项，权重就对不上，
> 专家问卷也得重新发、重新打。所以**方案可以随便定，打分项不能动**。

---

## 一、方案清单（你填：项目名 + 比哪几个方案）

| 环节 | 位置 |
|---|---|
| 方案清单 | `{scheme_list}` |

首次运行已自动生成。**要填两样**：

| 位置 | 填什么 |
|---|---|
| **B1 格（项目名称）** | 你的项目全称。这个名字会印在《AHP权重表》《AHP方案评分结论》《AHP权重与评分结论》三份成果文件上。**不填的话成果上会是占位名，第 09 步会拦住不让交付。** |
| 「方案清单」页 | 方案名、方案说明、是否参与评分（预置 6 行，行数可增可删） |

> 想让系统先给几个常见候选方案让你改？跑 `{run}`（默认是全空白）。

| 列 | 填什么 |
|---|---|
| 方案编号 | 默认 S1/S2/S3…，一般不用改 |
| 方案名称 | 黄色格，**填你项目的真实方案名**（如「片区整体拆除重建」「老旧小区综合整治」） |
| 方案说明 | 一句话说清「改什么、怎么改」，建议带上投入量级、见效周期、对居民影响 |
| 是否参与评分 | 填「是」或「否」；标「否」的不进最终排名 |

**「候选方案参考」页**给的是常见类型清单（按改造强度/功能路径/实施节奏/资金模式分类），
是给你灵感、不是让你照抄 —— 真实项目通常只从 1~2 个维度里各取一个交叉组合。

> 方案数一般 2~4 个：少于 2 个没法比；多于 4~5 个评分人负担重、结论不易收敛。
> 方案之间要有**实质差异**（改法不同，而不是面积不同），否则比出来的名次没有决策意义。

---

## 二、专家问卷（由专家填）

### 去哪拿

| 环节 | 目录（绝对路径） |
|---|---|
| 领空白卷 | `{pending}` |
| 填完放回 | `{recv}` |

文件名保持 `AHP专家问卷-专家N.xlsx`，**一位专家一份，不要合并、不要改名**。

### 怎么填（专家那份 xlsx 里）

每份问卷 4 个工作表：

| 工作表 | 要不要填 | 填什么 |
|---|---|---|
| **封面** | 填 2 处 | 「专家角色」填职称/单位；「日期」填 `2026-10-31` 这样 |
| **准则层** | 填 6 格 | 4 个准则两两比较，**只填灰色上三角** |
| **指标层** | 填 24 格 | 4 组指标各一个 4×4 矩阵，**只填灰色上三角** |
| **自动校验** | 不填 | 看结果：CR<0.1 显示「✔ 可接受」才算填对 |

**问卷里已经带好的东西（不用另外找）**：
- 封面有**两套标度的完整口径**（准则层 1–9 标度、指标层比值标度 0.1~10），照着填即可；
- 「指标层」每个矩阵**右侧有一列「指标说明」**，写清这个指标到底指什么 ——
  判断前先看一眼，判断会更靠谱；
- 「自动校验」页实时显示缺项/越界/互反异常与 CR 预估。

**只填灰色格**：对角线固定是 1，下三角是自动取倒数的公式，**这两类都不要碰**。
一共有 30 个灰格要填（准则层 6 + 指标层 4 组各 6）。

### 填什么数字

| 层 | 标度 | 填法 | 例子 |
|---|---|---|---|
| 准则层 | 1–9 标度 | 填 a_ij，1=同等，3=稍重要，5=明显重要，7=强烈重要，9=极端重要；2/4/6/8 为中间值 | 「现状情况 比 投资经济 重要得多」→ 填 7 或 9 |
| 指标层 | 比值标度 0.1~10 | 填 i 是 j 的**几倍**（i 相对 j 的比例） | 「老化程度比安全风险 重要 3 倍」→ 填 3；反过来填 0.33 |

> 记不准就用 1/3/5/7/9 这些中间值；不用纠结小数。
> 一致性（CR）由系统自动算，**不要为了好看去凑数**——真填最重要。

### 四个准则 / 16 个指标

| 准则层 | 该组指标 |
|---|---|
| 现状情况 | 房屋设施老化程度、安全风险等级、现状建筑密度、现状容积率 |
| 投资经济 | 改造成本、投资收益潜力、成本回收期、资产增值空间 |
| 规划合规 | 规划符合性、用地功能调整、基础设施配套、生态与风貌协调 |
| 风险可控 | 政策风险、造价波动、社会稳定、实施工期 |

---

## 三、方案评分表（你填）

### 什么时候出现

**你填好方案清单之后才会生成** —— 评分表的列数 = 你的方案数，
所以要先知道比哪几个方案才能生成对的表。

- 你填好《方案清单.xlsx》→ 重跑一次 `开工.py` → 系统按方案数生成评分表并停下等你填
- 生成位置：`{score}`

**方案名已经自动带过来了** —— 生成时按《方案清单.xlsx》预填了方案列，
所以评分表里的「方案信息」和表头就是你在清单里定的方案，不用再改表头。

### 怎么填

| 工作表 | 要不要填 | 填什么 |
|---|---|---|
| **填写说明** | 不填 | 口径说明，自己看 |
| **方案信息** | 核对 | 已按方案清单预填；只有要改名时才动 |
| **方案评分表** | 填黄色格 | 每个方案在 16 个指标上的 0–10 分 |

**只改黄色格和「依据」列**。「指标ID」「综合权重」「参考分档」三列是系统生成的，改了会对不上。

**「参考分档」列已经写好了这个指标打 1/3/5/7/9 分分别代表什么样子** —— 照着它给分，
不用自己判断标准（比如「房屋设施老化程度」：1 分=普遍老旧失修、5 分=中度老化主体基本可用、
9 分=建成较新设施完好）。这是行业标准口径，不用自己定。

| 列 | 填什么 |
|---|---|
| 方案一/二/三得分（黄色） | 0–10 分，1 位小数，建议按 1/3/5/7/9 打 |
| 依据-方案一/二/三 | 这分为何这么打（现状分析结论 / 财务测算 / 专家研判），一句话 |

「评分标准」列里已经写好了每个指标 1/3/5/7/9 分各代表什么，照着打即可。

### 评分口径

| 综合得分 | 评级 | 含义 |
|---|---|---|
| [9, 10] | 优秀 | 优先投资、全力推进 |
| [7, 9) | 良好 | 可行投资、分期推进 |
| [6, 7) | 一般 | 审慎投资、优化方案 |
| [0, 6) | 不可行 | 暂缓 / 放弃投资 |

**缺项不能打 0 分**：某个指标没依据就别填，系统会点名要求补；用 0 分掩盖缺数据会被拦下。

### 填完怎么续跑

```text
{run_from08}
```

> 已装 Python 3.9+ 的机器，用 `python {skill}\\开工.py --from 08` 也一样。

---

## 三、几条容易踩的坑

1. **别自己造问卷**。用 `--to 02` 生成的空白卷，标度和自动校验页都是配套的，自己另做会缺列。
2. **别改名、别合并**。回收区靠文件名认专家，5 份必须 5 个文件。
3. **问卷是匿名的**。封面项目名固定是「某某片区改造项目」，真实项目名只出现在你的成果文件里——
   这是为了避免专家被项目背景锚定判断。
4. **系统不替你填**。问卷和评分表都必须真人填；缺份就停下登记缺口，不会编数据。
5. **CR 超 0.1 不会硬压**。系统会告诉你是哪位专家、哪个矩阵、CR 多少、建议怎么改，
   你转达给专家本人调整，再重跑。
6. **方案名先在清单里定**。别等评分表出来才想方案叫什么 —— 评分表按《方案清单.xlsx》预填，
   两边不一致会对不上。
"""


def write_howto() -> Path:
    """生成《怎么填.md》：讲清三张人填的表怎么填。"""
    target = ROOT / "怎么填.md"
    target.write_text(HOWTO_MD.format(
        scheme_list=norm(ROOT / SCHEME_LIST_REL),
        ws=norm(ROOT),
        pending=norm(ROOT / PENDING_DIR),
        recv=norm(ROOT / RECV_DIR),
        score=norm(ROOT / SCORE_DIR / "AHP方案评分表.xlsx"),
        skill=SKILL.name,
        run=cmd_hint("--preset-schemes"),
        run_from08=cmd_hint("--from 08"),
    ), encoding="utf-8")
    return target


def run_intake(rows) -> int:
    """一项一问：给绝对路径、给准备方法、确认是否已放好。返回 0 继续，3 待补料。"""
    if not sys.stdin or not sys.stdin.isatty():
        say("  当前不是交互终端，只输出清单；请在终端里重跑 --intake，或直接按清单放料。")
        return 3
    say("")
    say("  ---- 收料引导（一项一项来；回车=已放好，输入 s=暂时没有，q=退出）----")
    for item, hits in rows:
        if hits:
            say("  [已到位] {}：{}".format(item["name"], norm(hits[0].relative_to(ROOT))))
            continue
        say("")
        say("  需要：{}（{}）".format(item["name"], "必需" if item["required"] else "可选"))
        say("  放到：{}".format(norm(ROOT / item["dst"])))
        say("  怎么来：{}".format(item["how"]))
        say("  用途：{}".format(item["why"]))
        if item.get("empty_hint"):
            say("  下一步：{}".format(item["empty_hint"].format(
                run=cmd_hint("--to 02"), pending=norm(ROOT / PENDING_DIR))))
        while True:
            try:
                answer = input("  放好了吗？[回车=好了 / s=没有 / q=退出] ").strip().lower()
            except EOFError:
                # stdin 被重定向（如 < /dev/null）而 isatty() 判断漏了这种情形：
                # 没有输入可读，视为「暂时没有」，按该项的缺料口径处理，不要崩栈。
                answer = "s"
            if answer in ("", "y", "yes"):
                break
            if answer in ("s", "n", "no", "没有"):
                if item["required"]:
                    say("  ! 这一项是硬门槛，缺了不能开工。补好后重跑。")
                    return 3
                say("  · 记为缺口：产物照样给初步判定，缺口进数据缺口清单，不影响开工。")
                break
            if answer in ("q", "quit", "exit"):
                say("  已退出收料引导；清单已保存到 放料清单.md")
                return 3
            say("  只认 回车 / s / q。")
        found = []
        for pattern in item["patterns"]:
            found.extend(path for path in (ROOT / item["dst"]).rglob(pattern)
                         if path.is_file() and not path.name.startswith("~$") and "空白" not in path.name)
        if not found:
            say("  ! 没在 {} 下找到匹配 {} 的文件{}".format(
                norm(ROOT / item["dst"]), "、".join(item["patterns"]),
                "（必需项，先放料）" if item["required"] else ""))
            if item["required"]:
                return 3
    say("")
    say("  收料完成，继续校验与开工。")
    return 0


MIRRORS = [
    ("aliyun", "https://mirrors.aliyun.com/pypi/simple/", "mirrors.aliyun.com"),
    ("tsinghua", "https://pypi.tuna.tsinghua.edu.cn/simple", "pypi.tuna.tsinghua.edu.cn"),
    ("ustc", "https://mirrors.ustc.edu.cn/pypi/simple", "mirrors.ustc.edu.cn"),
    ("pypi", "https://pypi.org/simple", "pypi.org"),
]



# ---------------------------------------------------------------- 小工具

def say(msg: str = "") -> None:
    print(msg, flush=True)


def norm(p: Path) -> str:
    return str(p).replace("\\", "/")


def read_json(path: Path, default=None):
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def write_board_status() -> None:
    """把工作区真实进度写进看板的 WORKSPACE_STATUS 块（顶栏据此显示真实状态）。

    看板是单文件 HTML，双击打开读不到工作区里的文件，所以状态只能由 开工.py 注入。
    只替换标记之间的那一段，其余字节原样保留；写不进去（包只读、看板被换掉等）
    就跳过 —— 看板会退回手动勾选计数，不影响主流程。
    """
    if not WORKSPACE_RESOLVED:
        return                      # 工作区没解析成功，别拿默认 ROOT 写错状态
    board = SKILL / "开工看板.html"
    if not board.is_file():
        return
    try:
        with open(board, "r", encoding="utf-8", newline="") as f:
            text = f.read()
    except OSError:
        return
    if BOARD_STATUS_BEGIN not in text or BOARD_STATUS_END not in text:
        return
    try:
        status = {
            "updatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "questionnaires": len(questionnaires()),
            "questionnairesRequired": MIN_QUESTIONNAIRES,
            "weightsComputed": (ROOT / "过程文件/05-权重与一致性/权重表.json").is_file(),
            "scoreSheetReady": (ROOT / SCORE_DIR / "AHP方案评分表.xlsx").is_file(),
            "deliverables": sum(1 for pat in DELIVERABLES for _ in ROOT.glob(pat)),
        }
    except Exception:
        return
    # </script> 出现在 JSON 里会提前闭合脚本块 —— 转义 "<" 即可
    payload = json.dumps(status, ensure_ascii=False, indent=2).replace("<", "\\u003c")
    eol = "\r\n" if "\r\n" in text else "\n"
    block = (BOARD_STATUS_BEGIN + eol + "<script>window.WORKSPACE_STATUS = "
             + payload + ";</script>" + eol + BOARD_STATUS_END)
    head, _, rest = text.partition(BOARD_STATUS_BEGIN)
    _, _, tail = rest.partition(BOARD_STATUS_END)
    try:
        with open(board, "w", encoding="utf-8", newline="") as f:
            f.write(head + block + tail)
    except OSError:
        pass


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def rel_to_root(path: Path) -> str:
    """相对工作区根的路径（给用户复制命令用）。跨盘时退回绝对路径。"""
    try:
        return str(path.relative_to(ROOT))
    except ValueError:
        return str(path)


def cmd_hint(*args: str) -> str:
    """给用户看的一条可直接粘贴的命令。

    优先用包内便携 python（runtime/python）—— 没装 Python 的机器也能跑；
    包里没有 runtime 时才退回系统 python。
    """
    if RUNTIME_PY.is_file():
        exe = rel_to_root(RUNTIME_PY)
    else:
        exe = "python"
    parts = [exe, rel_to_root(SKILL / "开工.py")]
    parts.extend(args)
    return " ".join(parts)


def missing_mods(py: str, mods: list[str]) -> list[str]:
    code = ("import importlib.util,json,sys\n"
            "mods=json.loads(sys.argv[1])\n"
            "print(json.dumps([m for m in mods if importlib.util.find_spec(m) is None]))")
    try:
        r = subprocess.run([py, "-c", code, json.dumps(mods)], capture_output=True,
                           text=True, timeout=90)
        return json.loads(r.stdout.strip().splitlines()[-1]) if r.returncode == 0 else list(mods)
    except Exception:
        return list(mods)


def venv_python() -> Path:
    name = "python.exe" if os.name == "nt" else "python"
    sub = "Scripts" if os.name == "nt" else "bin"
    return venv_dir() / sub / name


DELIVERY_MARKER = ".cardo_delivery_marker"


def _is_delivery_root(ws: Path) -> bool:
    """判断工作区是否为「交付库根」：交付库只读出包，不能在里面开工。

    唯一判据：**当前目录本身**存在标记文件 .cardo_delivery_marker。
    为什么不看「根下有 ≥2 个 Cardo* 目录」：用户经常把几个技能包并列放在
    同一个项目目录（如 测试-全流程贯通\）里使用，那是正常用法，不能拦。
    """
    try:
        return (ws / DELIVERY_MARKER).exists()
    except OSError:
        return False


def _has_multi_packages(ws: Path) -> bool:
    """并列多个 Cardo* 目录只作提示，不作拒绝条件。"""
    try:
        pkgs = [d for d in ws.iterdir() if d.is_dir() and d.name.startswith("Cardo")]
        return len(pkgs) >= 2
    except OSError:
        return False


def resolve_workspace(value: str) -> Path:
    """工作区 = 技能包所在的上一级目录（可 --workspace 覆盖）；技能包只读。"""
    ws = Path(value).resolve() if value else SKILL.parent
    if _is_delivery_root(ws):
        say(f"[错误] 这是交付库根目录，不能在这里开工：{ws}")
        say("        交付库是只读的发布产物，开工会在库里生成 过程文件/成果文件。")
        say("        正确做法：把技能包复制到项目目录后再运行，例如")
        say("          <项目目录> 下放 CardoAHP\\，再在该目录运行开工.py")
        raise SystemExit(2)
    if ws == SKILL or SKILL in ws.parents:
        say(f"[错误] 工作区不能是技能包目录或其子目录：{ws}")
        raise SystemExit(2)
    if not ws.is_dir():
        say(f"[错误] 工作区不存在：{ws}")
        raise SystemExit(2)
    if _has_multi_packages(ws):
        say("[提示] 当前目录并列了多个 Cardo* 技能包；未发现 .cardo_delivery_marker，按普通工作区继续。")
    if len(ws.parts) <= 1:
        # 只有当工作区本身就是盘符根（如 D://）时才拦；判断看自身层级，
        # 不能用 ws.parent.name —— 那样 D://任意文件夹 的父目录也是 D://，会被误拦。
        say("[错误] 技能包不能直接放在盘符根目录；请放进一个项目文件夹后再跑。")
        raise SystemExit(2)
    return ws


# ---------------------------------------------------------------- 环境

def make_dirs() -> list[str]:
    made = []
    for d in DIRS:
        p = ROOT / d
        if not p.is_dir():
            p.mkdir(parents=True, exist_ok=True)
            made.append(d)
    return made


def release_assets() -> list[str]:
    problems = []
    for a in ASSETS:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if not src.is_file():
            problems.append(f"包内缺资产：{a['src']}（包被裁剪，需重新解压原包）")
            continue
        if dst.is_file():
            if a.get("sha256") and sha256_file(dst).upper() != a["sha256"].upper():
                problems.append(f"模板被改动：{a['dst']}（第 07 步会校验原件；恢复方法：删掉该文件后重跑本命令）")
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        say(f"  释放模板：{a['dst']}")
    for a in INPUT_TEMPLATES:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if src.is_file() and not dst.is_file():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            say(f"  释放输入模板：{a['dst']}")
    return problems


def release_scheme_list(py: str) -> str | None:
    """首跑释放《方案清单.xlsx》（空白待填 + 候选方案参考），返回路径或 None。

    方案名是新项目要定的事 —— 与其让用户面对空白的评分表猜着填，
    不如先给一张带参考清单的登记表，改名/增删行即可。
    """
    dst = ROOT / SCHEME_LIST_REL
    if dst.is_file():
        return None
    script = SKILL / "Skills/08-方案打分/scripts/gen_scheme_list.py"
    if not script.is_file():
        return None
    dst.parent.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([py, str(script), "--out", str(dst)],
                       cwd=str(ROOT), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        say("  ! 方案清单生成失败：{}".format((r.stderr or r.stdout or "").strip()[-200:]))
        return None
    say(f"  释放输入模板：{SCHEME_LIST_REL}（空白待填 + 候选方案参考页）")
    return dst


def make_templates(py: str, experts: int, project_name: str, preset: bool = False) -> bool:
    """空白工作区：一次性铺好所有要人填的空白表。

    默认场景就是空白文件夹 = 全新项目，用户手上什么都没有。
    系统该做的是把表全铺出来、把「要你定的事」列清楚，等用户填完放回继续算，
    而不是先卡在「缺参数」上。返回是否成功。

    preset：True = 预填建议方案名（用户改）；False = 全空白（用户自己从零填）。
    这个选择是策略问题，交给用户定 —— 默认给空白，不替用户塞东西。
    """
    script = SKILL / "Skills/08-方案打分/scripts/make_templates.py"
    if not script.is_file():
        return False
    cmd = [py, str(script), "--workspace", str(ROOT),
           "--experts", str(experts), "--project-name", project_name or ""]
    if preset:
        cmd.append("--preset")
    r = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    out = (r.stdout or "") + (r.stderr or "")
    if r.returncode != 0:
        say("  ! 空白表生成有缺：{}".format(out.strip()[-300:]))
        return False
    for line in (r.stdout or "").rstrip().splitlines():
        if line.strip():
            say(line.rstrip())
    return True


def install(py: str) -> bool:
    req = SKILL / "requirements.txt"
    if not req.is_file():
        return True
    for name, url, host in MIRRORS:
        say(f"  安装依赖（{name}）…")
        r = subprocess.run([py, "-m", "pip", "install", "-r", str(req), "-i", url,
                            "--trusted-host", host, "--disable-pip-version-check", "--no-input"],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=900)
        if r.returncode == 0:
            return True
        say(f"  {name} 失败，换下一个源。")
    return False


def ensure_env(no_install: bool) -> tuple[str | None, bool]:
    # 1) 包内便携环境：拷包即用，不联网、不建 venv
    if RUNTIME_PY.is_file() and not missing_mods(str(RUNTIME_PY), IMPORTS):
        return str(RUNTIME_PY), False
    # 2) 系统 Python 已有依赖
    if not missing_mods(sys.executable, IMPORTS):
        return sys.executable, False
    py = venv_python()
    if py.is_file() and not missing_mods(str(py), IMPORTS):
        return str(py), False
    if no_install:
        say(f"  缺依赖：{'、'.join(missing_mods(sys.executable, IMPORTS))}（已指定 --no-install，不安装）")
        return None, False
    if not py.is_file():
        say("  建虚拟环境 .cardo/venv-ahp …")
        venv_dir().parent.mkdir(parents=True, exist_ok=True)
        r = subprocess.run([sys.executable, "-m", "venv", str(venv_dir())],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=600)
        if r.returncode != 0:
            say("  创建 .venv 失败：")
            say((r.stderr or r.stdout or "")[-800:])
            return None, False
    if not install(str(py)):
        say("  依赖安装失败。离线环境可手动拷 wheel 后执行：")
        say("  python -m pip install --no-index --find-links <目录> -r requirements.txt")
        return None, False
    left = missing_mods(str(py), IMPORTS)
    if left:
        say(f"  安装后仍缺：{'、'.join(left)}")
        return None, False
    return str(py), True


# ---------------------------------------------------------------- 参数与输入

# ---------------------------------------------------------------- 参数建议
# 空文件夹 = 全新项目：项目名、专家人数、方案都是「新定」的。
# 与其干等着要参数，不如先给一套建议值自动填上，让用户看到成品再改。
# 建议值一律标 source="suggest"，与用户显式指定的（source="user"）区分开。

PROJECT_NAME_HINT = {
    "expert": "5 位（聚合下限；能凑 7 位更稳，结论更抗单个专家偏好）",
}


def load_params() -> dict:
    data = read_json(ROOT / PARAMS, {}) or {}
    data.setdefault("schema", SCHEMA)
    for k, v in DEFAULTS.items():
        data.setdefault(k, v)
    return data


def questionnaires() -> list[Path]:
    root = ROOT / RECV_DIR
    if not root.is_dir():
        return []
    return sorted([p for p in root.glob("AHP专家问卷-专家*.xlsx") if not p.name.startswith("~$")])


def duplicate_expert_groups(paths: list[Path]) -> list[tuple[str, list[str]]]:
    """按文件内容 sha256 分组，找出**内容完全相同**的问卷。

    为什么必须查：同一份问卷被复制成多份放进回收区 → 份数检查过、CR 也可能过，
    但聚合出来的"多位专家权重"其实只含 1 位专家的判断。这是静默失效，
    比缺料更隐蔽，所以单独硬拦。

    专家名单.xlsx 不参与（名字不匹配 AHP专家问卷-专家*.xlsx，这里再兜一层）。
    只读：只算哈希，不碰用户问卷文件。
    """
    buckets: dict[str, list[str]] = {}
    for p in paths:
        if "名单" in p.name or p.name.startswith("~$"):
            continue
        try:
            buckets.setdefault(sha256_file(p), []).append(p.name)
        except OSError:
            continue
    return [(h, sorted(names)) for h, names in buckets.items() if len(names) > 1]


def write_duplicate_report(groups: list[tuple[str, list[str]]]) -> Path:
    """把查重结论写进《回收校验报告.md》，让用户和 AI 有一处可查。

    幂等：反复重跑只保留最新一段；查重通过时清掉上一轮留下的那段，
    免得报告里留着过期结论。没有要写的内容时不动文件（不碰流水线生成的报告）。
    """
    target = ROOT / "过程文件/03-回收校验/回收校验报告.md"
    old = target.read_text(encoding="utf-8") if target.is_file() else ""
    # 去掉上一轮写的查重段（本段总是文件最后一段）
    cut = old.find(DUP_REPORT_HEADING)
    if cut >= 0:
        old = old[:cut].rstrip()
    if not groups:
        # 只在确实有过查重段时才回写，避免动流水线生成的报告
        if cut >= 0:
            target.write_text((old + "\n").rstrip("\n") + "\n", encoding="utf-8")
        return target

    target.parent.mkdir(parents=True, exist_ok=True)
    lines = [DUP_REPORT_HEADING, "",
             "> 由 开工.py 在输入校验阶段写入（只读检查，**没有改动任何问卷文件**）。", "",
             "以下文件的 sha256 完全相同 —— 它们是同一份问卷的副本，"
             "不能算作多位独立专家（否则聚合权重只反映 1 位专家的判断）：", "",
             "| 文件 | sha256 |", "|---|---|"]
    for digest, names in groups:
        for name in names:
            lines.append("| {} | `{}` |".format(name, digest))
    lines += ["",
              "**怎么办**（不要替专家改数，也不要替专家重填）：",
              "",
              "1. 确认回收区里这几份确实是同一份问卷的副本；",
              "2. 把多余的副本移出回收区，或换成对应专家真正填写的那一份；",
              "3. 确认回收区里至少 {} 份问卷内容各不相同后重跑：".format(MIN_QUESTIONNAIRES),
              "",
              "```text",
              cmd_hint("--from 03"),
              "```",
              ""]
    body = "\n".join(lines)
    if old:
        body = old + "\n\n" + body
    else:
        body = "# AHP 问卷回收校验\n\n" + body
    target.write_text(body, encoding="utf-8")
    return target


def pending_surveys() -> list[Path]:
    root = ROOT / PENDING_DIR
    if not root.is_dir():
        return []
    return sorted([p for p in root.glob("AHP专家问卷-专家*.xlsx") if not p.name.startswith("~$")])


def precheck(params: dict, need_q: bool, need_project: bool = True) -> tuple[list[str], list[str]]:
    """校验参数与输入。

    need_project：是否必须有项目全称。
      · 发卷阶段（--to 02）不需要 —— 问卷匿名，封面固定占位名；
      · 第 08/09 步需要（成果文件与结论文档要印项目名）；
      · 第 01–07 步不需要（权重表与模板无关项目名）。
      另：问卷还没收齐时本轮必停在 03，此时不必追着要项目名 ——
      先让用户把问卷发出去，项目名等收齐后再给。

    项目名走「用户填」而非「系统猜」：优先取《方案清单.xlsx》B1 格用户自己填的，
    其次是 --project 指定的，都没有就留空 —— 成果文件用占位名，第 09 步会明确提醒。
    专家人数用建议值（默认 5 位），不硬阻断 —— 用户看到成品再改，比先填表再开工顺。
    """
    blockers, notes = [], []
    if need_project and not params.get("projectName"):
        # 不阻断：项目名可以在《方案清单.xlsx》里填，先把能做的做完。
        notes.append(f"项目名：还没填 —— 成果文件上会印占位名「某某片区改造项目」。"
                     f"请在 {SCHEME_LIST_REL} 的 B1 格填，或用 --project \"正式全称\" 指定")
    experts = params.get("experts")
    if not isinstance(experts, int) or not (5 <= experts <= 7):
        blockers.append("参数 experts 必须是 5–7 的整数（用 --experts 指定）")

    if not (ROOT / TEMPLATE_DST).is_file():
        blockers.append(f"四准则模板不在位：{TEMPLATE_DST}（删掉该文件后重跑本命令可恢复）")

    qs = questionnaires()
    pend = pending_surveys()
    if need_q and len(qs) < MIN_QUESTIONNAIRES:
        if pend:
            blockers.append(
                f"回收问卷不足：现有 {len(qs)} 份，至少 {MIN_QUESTIONNAIRES} 份。"
                f"空白问卷已在领料区（{PENDING_DIR}，{len(pend)} 份）——"
                f"发给专家填完原样放回 {RECV_DIR}，收齐后再跑"
            )
        else:
            blockers.append(
                f"回收问卷不足：现有 {len(qs)} 份，至少 {MIN_QUESTIONNAIRES} 份。"
                f"先跑「{cmd_hint('--to 02')}」生成空白问卷（领料区 {PENDING_DIR}），"
                f"发给专家填完原样放回 {RECV_DIR}"
            )
    elif qs:
        notes.append(f"已回收问卷 {len(qs)} 份")

    # 查重：同一份问卷被复制成多份放进回收区 → 份数够、CR 也可能过，
    # 但聚合出来的"多位专家权重"其实只含 1 位专家的判断。属于静默失效，硬拦。
    # 无重复时也调用一次：把上一轮留在报告里的查重段清掉（不误伤流水线自己的报告）。
    if need_q and len(qs) >= 2:
        dup_groups = duplicate_expert_groups(qs)
        if dup_groups:
            lines = []
            for digest, names in dup_groups:
                lines.append("    " + "、".join(names))
                lines.append(f"      sha256 = {digest}")
            blockers.append(
                f"[{DUP_EXPERT_CODE}] 回收问卷内容重复：下面几份文件的 sha256 完全相同，"
                f"实际只算 1 位专家的判断，不能当多位独立专家参与聚合：\n"
                + "\n".join(lines)
                + f"\n  同一份问卷被复制成多份不能算多位专家。请把多余的副本移出回收区"
                  f"（或换成对应专家真正填写的那一份），确认回收区里至少 "
                  f"{MIN_QUESTIONNAIRES} 份内容各不相同的问卷后重跑：\n"
                  f"    {cmd_hint('--from 03')}"
            )
        write_duplicate_report(dup_groups)

    if pend and not qs:
        notes.append(f"领料区有 {len(pend)} 份空白问卷待发放")
    if not need_project and not params.get("projectName"):
        notes.append("项目全称：这一步用不上（问卷匿名）；算成果前再给即可")

    roster = list((ROOT / RECV_DIR).glob("专家名单*.xlsx"))
    notes.append("专家名单：" + ("已填" if roster else "未放（可选，报告可读性更好）"))

    scores = ROOT / SCORE_DIR / "AHP方案评分表.xlsx"
    if scores.is_file():
        notes.append("方案评分表已在位")
    elif need_q:
        notes.append("方案评分表：08 首次运行会生成空白表（那时会停下等你填）")
    return blockers, notes


# ---------------------------------------------------------------- 主流程

def run_pipeline(cmd: list[str], env: dict) -> tuple[int, str]:
    """边跑边打印，同时留全文供判定停顿/失败位置。"""
    lines = []
    p = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
    for line in p.stdout:
        say(line.rstrip("\n"))
        lines.append(line)
    return p.wait(), "".join(lines)


CRASH_LOG = "过程文件/00-总控/异常日志.log"

# 异常类型 → (缺什么, 建议怎么处理)。只覆盖用户真会遇到的几类，
# 其余一律按「程序内部错误」处理，保留日志交给维护人。
_CRASH_HINTS = {
    "ModuleNotFoundError": ("缺少 Python 依赖（包内运行时可能不完整）",
                            "确认 runtime/ 与 Skills/ 是整体复制的；或用系统 Python 按 requirements.txt 装依赖"),
    "ImportError": ("运行环境不完整（缺模块或版本不符）", "重新完整复制 runtime/ 后重跑"),
    "FileNotFoundError": ("需要的文件不存在", "按上面给出的路径补齐后重跑"),
    "PermissionError": ("文件被占用或没有写权限",
                        "关掉正在打开该文件的 Word / Excel / 资源管理器预览窗格后重跑"),
    "UnicodeDecodeError": ("文件不是 UTF-8 编码", "用记事本另存为 UTF-8 后重跑"),
    "KeyError": ("输入表缺少必需的列或字段", "对照 放料清单.md 补齐该列后重跑"),
    "JSONDecodeError": ("JSON 文件内容损坏", "删掉该文件让程序重新生成，或恢复备份"),
}


def _last_step_hint() -> str:
    """从运行状态/运行日志里找最后一个进入的步骤，回答「哪一步失败」。"""
    for name in ("运行状态.json", "运行日志.json"):
        path = ROOT / "过程文件/00-总控" / name
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not isinstance(data, dict):
            continue
        if name == "运行状态.json":
            steps = data.get("steps") or {}
            for key in list(steps)[::-1]:
                st = steps.get(key) or {}
                if isinstance(st, dict) and st.get("status") in ("running", "failed"):
                    return "{} {}".format(key, st.get("name", "")).strip()
        else:
            entries = data.get("entries") or []
            for item in entries[::-1]:
                if isinstance(item, dict) and item.get("step"):
                    return "{} {}".format(item.get("step"), item.get("name", "")).strip()
    return "开工准备阶段"


def crash_report(exc: BaseException) -> int:
    """未预期异常的统一出口：人话四段 + 落日志；--debug 时保留原始堆栈。"""
    # 用 format_exception 而不是 format_exc：后者不在 except 块里调用时只会写出
    # "NoneType: None"，日志等于没记。这里始终带上 exc 自身的堆栈。
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log = ROOT / CRASH_LOG
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        with log.open("a", encoding="utf-8") as f:
            f.write("\n===== {} =====\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            f.write("技能包：{}\n".format(SKILL.name))
            f.write("命令：{}\n".format(" ".join(sys.argv)))
            f.write(tb)
    except OSError:
        log = None
    if "--debug" in sys.argv or os.environ.get("CARDO_DEBUG"):
        raise exc
    name = type(exc).__name__
    what, action = _CRASH_HINTS.get(
        name, ("程序内部错误", "把异常日志发给维护人；急用可先 --from <上一步> 续跑"))
    say("")
    say("[异常] 程序没有正常跑完")
    say("  哪一步失败：{}".format(_last_step_hint()))
    say("  缺什么：{}（{}: {}）".format(what, name, exc))
    say("  建议怎么处理：{}".format(action))
    say("  日志在哪：{}".format(norm(log.relative_to(ROOT)) if log else "（日志写入失败）"))
    say("  要看完整堆栈：命令末尾加 --debug 重跑")
    return 1


def main(argv=None) -> int:
    """跑完（含各种提前返回）后刷新看板顶栏的真实状态。"""
    try:
        return _main(argv)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        say("\n[中断] 已停止；已完成步骤的产物保留，续跑用 --from <步>。")
        return 130
    except BaseException as exc:  # noqa: BLE001
        return crash_report(exc)
    finally:
        try:
            write_board_status()
        except Exception:
            pass


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=f"{TITLE} · 一条命令开工")
    ap.add_argument("--project", default="", help="项目全称（也可在《方案清单.xlsx》B1 格填）")
    ap.add_argument("--demo", action="store_true",
                    help="演示样例：允许 demo 数据出交付物；同时写进 工作区参数.json 的 demo 键，之后重跑自动沿用")
    ap.add_argument("--experts", type=int, default=None, help="专家人数（5–7，默认 5）")
    ap.add_argument("--preset-schemes", action="store_true",
                    help="预填建议候选方案名（用户改名用）；不给则给全空白")
    ap.add_argument("--from", dest="start_from", default="", help="从第几步开始（如 03 / 08）")
    ap.add_argument("--to", dest="stop_at", default="", help="跑到第几步为止（如 02 / 07）")
    ap.add_argument("--intake", action="store_true",
                    help="收料引导：一项一项问，给绝对路径与准备方法，确认后再校验开工")
    ap.add_argument("--check", action="store_true", help="只做结构与输入校验，不跑流水线")
    ap.add_argument("--no-install", action="store_true", help="缺依赖时不安装")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要执行的命令")
    ap.add_argument("--workspace", default="", help="工作区根（缺省为技能包所在目录，即上一级）")
    ap.add_argument("--debug", action="store_true",
                    help="出错时打印完整堆栈（默认只给人话提示）")
    ap.add_argument("passthrough", nargs=argparse.REMAINDER, help="其余参数原样传给流水线")
    args = ap.parse_args(argv)

    if not os.environ.get("CARDO_REEXEC"):
        say(f"== {TITLE} · 开工 ==")
    if sys.version_info < PY_MIN:
        say(f"[环境] 需要 Python 3.{PY_MIN[1]} 及以上，当前 {sys.version.split()[0]}。")
        return 2

    global ROOT, WORKSPACE_RESOLVED
    ROOT = resolve_workspace(args.workspace)
    WORKSPACE_RESOLVED = True
    if not os.environ.get("CARDO_REEXEC"):
        say(f"[0/4] 工作区：{ROOT}（技能包只读：{SKILL.name}）")
    if not os.environ.get("CARDO_REEXEC"):
        board = SKILL / "开工看板.html"
        if board.is_file():
            say("      开工看板：{}（双击打开：步骤、放料清单、可复制提示词）".format(norm(board)))

    py, installed = ensure_env(args.no_install)
    if py is None:
        return 2
    if py != sys.executable:
        say(f"[1/4] 依赖：改用 {py}")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
                   PYTHONDONTWRITEBYTECODE="1", CARDO_REEXEC="1")
        r = subprocess.run([py, str(Path(__file__).resolve())] + list(sys.argv[1:]),
                           cwd=str(ROOT), env=env)
        return r.returncode
    if py == str(RUNTIME_PY):
        say("[1/4] 依赖：就位（包内便携环境，无需联网）")
    else:
        say("[1/4] 依赖：就位" + ("（已装入 .venv）" if installed else ""))

    made = make_dirs()
    problems = release_assets()
    if problems:
        for p in problems:
            say(f"  ✗ {p}")
        return 2
    say(f"[2/4] 目录：{'新建 ' + '、'.join(made) if made else '已就绪'}；四准则模板在位且 sha 一致")

    howto_path = write_howto()
    say(f"  填写说明：{norm(howto_path.relative_to(ROOT))}（专家问卷怎么填、评分表怎么填，逐格说明）")

    # 参数先定下来（建议式），放料清单里才能显示「本次替你定的值」。
    params = load_params()
    if args.demo:
        params["demo"] = True
    # demo 键必须留在参数文件里：否则「--demo 跑过一次」的声明会在下次重写时丢掉，
    # 再跑又变成未声明演示 → 08 重新拒收，提示里给的续跑命令自己失效。
    params["demo"] = bool(params.get("demo"))
    # 项目名：用户显式指定 > 已存参数 > 从文件夹名推断的建议值。
    # 建议值照常写进参数（后面成果文件要用），但标 source=suggest 让人知道可改。
    # 项目名来源，按优先级：
    #   1) 用户在《方案清单.xlsx》B1 格自己填的（拿得到的模板里填，不必问 AI）
    #   2) 命令行 --project（用户在这儿直接说了）
    #   3) 都没有 → 留空，成果文件用占位名，并在清单里提醒用户去填。
    # 刻意**不**从文件夹名推断 —— 文件夹名常是「测试-AHP」这类内部代号，
    # 猜出来的名字会让成果文件印上一个没人认可的项目名。
    sheet_project = ""
    try:
        import openpyxl
        _wb = openpyxl.load_workbook(ROOT / SCHEME_LIST_REL, data_only=True)
        if "方案清单" in _wb.sheetnames:
            sheet_project = str(_wb["方案清单"].cell(row=1, column=2).value or "").strip()
        _wb.close()
    except Exception:
        sheet_project = ""
    if args.project:
        params["projectName"] = args.project
        params["projectNameSource"] = "user"
    elif sheet_project:
        params["projectName"] = sheet_project
        params["projectNameSource"] = "sheet"
    elif not str(params.get("projectName") or "").strip():
        # 三处都没有才留空。B1 为空时**不得**清掉参数里已有的项目名：
        # 否则续跑命令（不带 --project）会把名字抹掉，09 立刻报「占位名不可交付」。
        params["projectName"] = ""
        params["projectNameSource"] = "unset"
    else:
        # B1 为空但参数里已有名字（上次跑过/用户手填过）：原样保留。
        params.setdefault("projectNameSource", "user")
    if args.experts is not None:
        params["experts"] = args.experts
        params["expertsSource"] = "user"
    else:
        params.setdefault("expertsSource", "suggest")
    write_json(ROOT / PARAMS, params)

    # 空白工作区 = 全新项目：一次性把所有要人填的空白表铺出来。
    # （方案清单、方案评分表、专家空白问卷）—— 用户手上什么都没有时，
    # 系统该给表 + 说清要定什么，而不是先卡在缺参数上。
    fresh_workspace = not (ROOT / "工作区参数.json").exists() or not any(
        (ROOT / "过程文件").glob("*/**/*.json")) if (ROOT / "过程文件").is_dir() else True
    if not questionnaires() and not pending_surveys():
        say("  空白工作区：铺好全部待填空白表")
        make_templates(py, params["experts"], params["projectName"],
                       preset=args.preset_schemes)

    rows = intake_scan()
    ready = sum(1 for item, hits in rows if hits)
    list_path = write_intake_list(rows, params)
    say(f"  放料清单：{norm(list_path.relative_to(ROOT))}（已到位 {ready}/{len(rows)} 项）")
    if args.intake:
        code = run_intake(rows)
        rows = intake_scan()
        write_intake_list(rows, params)
        if code != 0:
            say(f"  按清单放料后重跑：{cmd_hint()}，或再跑 {cmd_hint('--intake')} 逐项确认")
            return code

    say("[3/4] 参数与输入")

    # 目标步骤决定要不要项目全称：
    #   发卷（--to 02）不需要（问卷匿名）；第 08/09 步需要（成果要印项目名）；
    #   第 01–07 步不需要。
    need_q = not (args.stop_at and args.stop_at <= "02")
    if args.stop_at:
        need_project = args.stop_at >= "08"
    else:
        need_project = len(questionnaires()) >= MIN_QUESTIONNAIRES

    # 首次开工：还没问卷就直接把空白卷发出来，别让新手自己猜命令。
    # 这一步必须先于「回收问卷不足」的判定 —— 否则会被自己的 blocker 挡住。
    if not args.stop_at and not questionnaires() and not pending_surveys():
        say("  首次开工：自动发卷（生成空白问卷到领料区）")
        auto_issue = True
    else:
        auto_issue = False

    if auto_issue:
        # 发卷阶段不需要项目全称（问卷匿名），也不需要已回收问卷
        blockers, notes = precheck(params, False, False)
        for n in notes:
            say(f"  · {n}")
        if blockers:
            for b in blockers:
                say(f"  ✗ {b}")
            return 3
        say(f"  项目：{params.get('projectName') or '（未指定，问卷匿名用占位名）'}｜专家：{params['experts']} 位")
        if args.check:
            say("  校验通过（--check 不跑流水线）。")
            return 0
        say("[4/4] 流水线（先发卷）")
        cmd = [py, str(SKILL / PIPELINE), "run", "--workspace", ".", "--to", "02"]
        if params.get("demo"):
            cmd += ["--demo"]
        cmd += list(args.passthrough)
        say("  " + " ".join(cmd))
        if args.dry_run:
            return 0
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
                   PYTHONDONTWRITEBYTECODE="1", CARDO_SKILLS_ROOT=str(SKILL / "Skills"))
        code, text = run_pipeline(cmd, env)
        if code != 0:
            say(f"[失败] 发卷阶段退出码 {code}，看上方输出。")
            return code if code in (1, 2) else 1
        write_intake_list(intake_scan(), params)
        pend = pending_surveys()
        say("")
        say("  ---- 现在轮到你（做完重跑开工.py，我接着算）----")
        say(f"  ① 定方案（你）：{norm(ROOT / SCHEME_LIST_REL)}")
        say("     · B1 格填项目名称（会印在三份成果文件上）")
        say("     · 「方案清单」页填方案名与说明（行数不限，一般 2~4 个）")
        say(f"  ② 发问卷（你→专家）：{norm(ROOT / PENDING_DIR)} 的 {len(pend)} 份，一份一人")
        say(f"     专家填完原样放回：{norm(ROOT / RECV_DIR)}")
        say(f"  ③ 回收后跑：{cmd_hint('--to 07')} 算权重与一致性")
        say(f"  ④ 打分：评分表自动生成，黄色格填 0–10 分、依据列写出处")
        say(f"  ⑤ 出成果：{cmd_hint('--from 08')}")
        say(f"  完整步骤与状态：{norm(howto_path.relative_to(ROOT).parent / '放料清单.md')}"
            f"（每次运行刷新）")
        return 0

    blockers, notes = precheck(params, need_q, need_project)
    for n in notes:
        say(f"  · {n}")
    if blockers:
        for b in blockers:
            say(f"  ✗ {b}")
        say(f"  放料目录（绝对路径）：{norm(ROOT / '输入文件')}")
        say(f"  领料区（空白问卷）：{norm(ROOT / PENDING_DIR)}")
        say(f"  回收区（放专家填好的）：{norm(ROOT / RECV_DIR)}")
        say(f"  填写说明：{norm(howto_path.relative_to(ROOT))}")
        say(f"  逐项确认：{cmd_hint('--intake')}")
        say(f"  补齐后重跑：{cmd_hint()}")
        return 3
    say(f"  项目：{params.get('projectName') or '（未指定，成果文件里会显示占位名）'}｜专家：{params['experts']} 位")
    if not need_q:
        say("  本次只到 02（发卷阶段），不要求已回收问卷。")

    # 建议摘要：把「这次替你定的值」摊开，让用户一眼看到哪里是建议、可以改。
    sug = []
    if not params.get("projectName"):
        sug.append(f"项目名：待填 —— 在 {SCHEME_LIST_REL} 的 B1 格填"
                   f"（或用 --project \"正式全称\"）")
    if params.get("expertsSource") == "suggest":
        sug.append(f"专家 {params['experts']} 位（{PROJECT_NAME_HINT['expert']}）")
    if (ROOT / SCHEME_LIST_REL).is_file():
        sug.append(f"候选方案：在 {SCHEME_LIST_REL} 的「方案清单」页自己填"
                   f"（想先看常见方案：--preset-schemes）")
    if sug:
        say("  要你定的（都能改）：")
        for s in sug:
            say(f"    · {s}")

    if args.check:
        say("  校验通过（--check 不跑流水线）。")
        return 0

    say("[4/4] 流水线")
    cmd = [py, str(SKILL / PIPELINE), "run", "--workspace", "."]
    if args.start_from:
        cmd += ["--from", args.start_from]
    if args.stop_at:
        cmd += ["--to", args.stop_at]
    if params.get("demo"):
        cmd += ["--demo"]
    cmd += list(args.passthrough)
    say("  " + " ".join(cmd))
    if args.dry_run:
        return 0

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
               PYTHONDONTWRITEBYTECODE="1", CARDO_SKILLS_ROOT=str(SKILL / "Skills"))
    code, text = run_pipeline(cmd, env)

    # 02 步跑完：把领料区位置直接告诉用户，别让他自己找。
    if code == 0 and args.stop_at == "02":
        write_intake_list(intake_scan(), params)
        pend = pending_surveys()
        say("")
        say("  ---- 现在轮到你（做完重跑开工.py，我接着算）----")
        say(f"  ① 定方案（你）：{norm(ROOT / SCHEME_LIST_REL)}")
        say("     · B1 格填项目名称（会印在三份成果文件上）")
        say("     · 「方案清单」页填方案名与说明（行数不限，一般 2~4 个）")
        say(f"  ② 发问卷（你→专家）：{norm(ROOT / PENDING_DIR)} 的 {len(pend)} 份，一份一人")
        say(f"     专家填完原样放回：{norm(ROOT / RECV_DIR)}")
        say(f"  ③ 回收后跑：{cmd_hint('--to 07')} 算权重与一致性")
        say(f"  ④ 打分：评分表自动生成，黄色格填 0–10 分、依据列写出处")
        say(f"  ⑤ 出成果：{cmd_hint('--from 08')}")
        say(f"  完整步骤与状态：{norm(howto_path.relative_to(ROOT).parent / '放料清单.md')}"
            f"（每次运行刷新）")
        return 0

    if code == 0:
        say("[完成] 全链通过。成果：")
        for pat in DELIVERABLES:
            for p in sorted(ROOT.glob(pat)):
                say(f"  - {norm(p.relative_to(ROOT))}")
        return 0

    # 定位失败的那一步：先读总控的运行日志（字段是 steps / failedStep），
    # 再退回命令流水（commands 里 exit != 0 的那条），最后才用 [STOP] [NN] 标记。
    # 注意：总控写的键是 steps/id/status，不是 entries/step/exit —— 读错键会让
    # 失败永远显示「步骤 ?」，续跑建议永远指 01 步，诊断信息形同虚设。
    failed = ""
    try:
        doc = read_json(ROOT / "过程文件/00-总控/运行日志.json", {}) or {}
        failed = (doc.get("failedStep") or "") or ""
        if not failed:
            steps = doc.get("steps") or []
            bad = next((s for s in reversed(steps)
                        if s.get("status") == "failed" or s.get("id") == doc.get("failedStep")), None)
            failed = (bad or {}).get("id") or ""
        if not failed:
            cmds = doc.get("commands") or []
            badc = next((c for c in reversed(cmds) if c.get("exit") not in (0, None)), None)
            failed = (badc or {}).get("step") or ""
    except Exception:
        failed = ""
    stop = re.search(r"\[STOP\]\s*\[(\d{2})\]", text)
    if not failed and stop:
        failed = stop.group(1)

    if code == 4:
        # 拒收原因与改法已由流水线 [STOP] 段打印，这里只补续跑命令，避免同一段话打两遍
        say(f"[拒收] 按上方提示处理后续跑：{cmd_hint('--demo --from ' + (failed or '08'))}")
        return 4

    if code == 3 or stop:
        say(f"[待确认] 停在需要人工介入的位置（步骤 {failed or '?'}）。")
        g = GUIDES.get(failed)
        if g:
            say(f"  做什么：{g}")
        say(f"  办完后续跑：{cmd_hint('--from ' + (failed or '01'))}")
        return 3

    say(f"[失败] 流水线退出码 {code}（步骤 {failed or '?'}）。")
    g = GUIDES.get(failed)
    if g:
        say(f"  做什么：{g}")
    report = ROOT / "过程文件/00-总控/运行报告.md"
    if report.is_file():
        say("  运行报告末段：")
        for line in report.read_text(encoding="utf-8", errors="replace").splitlines()[-20:]:
            say("  " + line)
    say(f"  修好后续跑：{cmd_hint('--from ' + (failed or '01'))}（不要跳过失败步骤）")
    return code if code in (1, 2) else 1


if __name__ == "__main__":
    sys.exit(main())
