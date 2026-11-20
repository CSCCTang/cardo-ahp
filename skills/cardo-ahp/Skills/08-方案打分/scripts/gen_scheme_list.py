#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_scheme_list.py -- 生成《方案清单.xlsx》（候选改造方案登记表，人填）。

这张表回答「比哪几个方案」，是第 08 步方案评分表的上游：
方案名、编号、说明在这里定，评分表按它预填方案列。

**版式（3 页）**
  · Sheet「方案清单」：空白待填 —— 序号 | 方案编号 | 方案名称 | 方案说明 | 是否参与评分
        预置 6 行空行（够常见项目用；不够就在下面加行），行数不限。
  · Sheet「候选方案参考」：常见方案类型清单（按改造强度/功能路径/实施节奏/资金模式分类），
        供挑选用 —— 不是让你照抄，是给灵感。挑中的抄到「方案清单」里改名即可。
  · Sheet「填写说明」：口径与用法。

设计要点：
  · **不预填固定方案名**。每个项目的真实方案不一样，填死三套名字只会让人
    懒得改、成果里带着不属于本项目的方案名。给参考清单而不是给答案。
  · 方案数不限（2~6 个为宜），第 08 步按实际行数生成评分表。

用法：
  python gen_scheme_list.py --out "输入文件/03-方案打分表/方案清单.xlsx" [--project-name 名称]
退出码：0 成功；2 参数错。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 空白待填行数（够用即可，用户可加行）
BLANK_ROWS = 6

# 预填模式（--preset）用的建议方案名：城市更新常见候选，作为用户改的起点。
# 默认不预填 —— 每个项目的真实方案不同，填死名字只会让人懒得改。
SUGGESTED = [
    ("全面拆除重建", "推倒重来原地重建；见效快、投入大、居民扰动强"),
    ("渐进式微更新", "保留主体结构与肌理，分期分区修缮；投入小、周期长"),
    ("功能置换", "保留建筑体量与立面，更换使用功能；投入居中、需产业导入"),
]

# 候选方案参考：按维度分类，供挑选而非照抄
REFERENCE = [
    ("按改造强度", [
        ("全面拆除重建", "推倒重来原地重建，见效快、投入大、居民扰动强"),
        ("拆建结合", "部分拆除部分保留新建，保留部分历史风貌"),
        ("保留提升", "主体结构保留，做立面、屋面、设施综合提升"),
        ("渐进式微更新", "保留肌理分期分区修缮，投入小、周期长"),
        ("结构加固+功能提升", "主体加固改造为主，功能在原空间内优化"),
    ]),
    ("按功能路径", [
        ("功能置换", "保留建筑体量与立面，更换使用功能，需产业导入"),
        ("产业导入/运营激活", "以业态和运营为核心，先招后建、渐进激活"),
        ("公共空间主导", "以广场绿地等公共空间为抓手，牵引片区活力"),
        ("基础设施补短板", "管网、道路、停车等配套先行，居住品质带动"),
    ]),
    ("按实施节奏", [
        ("分期滚动开发", "划片分期，成熟一片做一片，资金压力分散"),
        ("一次性整体推进", "统一规划统一实施，整体形象快但资金压力大"),
        ("试点先行再推广", "先做 1~2 个样板单元验证再全面铺开"),
    ]),
    ("按资金模式", [
        ("政府投资主导", "财政或专项资金投入，产权清晰、见效稳定"),
        ("社会资本主导", "引入企业投建运营，需明确回报机制"),
        ("多方混合", "政府出基础设施、企业出建设、居民出部分资金"),
    ]),
]

NOTES = [
    "AHP 候选改造方案清单（10 分制方案的登记表）",
    "",
    "一、这张表是干什么的",
    "1. 这里登记「比哪几个方案」：方案编号、名称、一句话说明。",
    "2. 第 08 步生成《AHP方案评分表》时，会按本表预填方案列（得分列与依据列）。",
    "3. 所以先在本表把方案定下来，评分表出来时直接对号入座，不用再改表头。",
    "",
    "二、怎么填（「方案清单」页）",
    "· 方案名称：黄色格，填你项目的真实方案名（如「片区整体重建」「老旧小区综合整治」）。",
    "· 方案说明：一句话说清这个方案「改什么、怎么改」，评分表会带过去。",
    "· 是否参与评分：填「是」或「否」。标「否」的方案不进最终排名（保留在清单里备查）。",
    "· 方案数量不限：预置 6 行不够就在下面加行；删行也可以。第 08 步按实际行数生成评分表。",
    "",
    "三、方案该有几个、怎么选",
    "1. 数量：一般 2~4 个。少于 2 个没法比；多于 4~5 个，评分人负担重、结论不易收敛。",
    "2. 差异：方案之间要有**实质差异**——改法不同（拆还是留、推倒还是分期），",
    "   而不是面积大小不同。否则比出来的名次没有决策意义，等于白比。",
    "3. 「候选方案参考」页给的是常见类型的思路清单，不是让你照抄。",
    "   真实项目的方案组合通常只从其中取 1~2 个维度交叉，需要你按项目实际定。",
    "",
    "四、配套文件",
    "· 专家问卷：输入文件/02-专家问卷（待发放）→ 专家填完放回（已回收）。",
    "· 方案评分表：首次运行即生成在 输入文件/03-方案打分表/AHP方案评分表.xlsx（空白待填）。",
    "· 填写说明见工作区根目录《怎么填.md》。",
]


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 AHP 候选方案清单 xlsx（空白待填 + 候选方案参考）")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    ap.add_argument("--project-name", default="",
                    help="预填的项目名称；留空则该格由用户自己填")
    ap.add_argument("--preset", action="store_true",
                    help="预填建议方案名（用户改）；不给则全空白")
    args = ap.parse_args()

    wb = Workbook()

    # ---------------- 方案清单（项目名 + 方案，都是用户填的） ----------------
    ws = wb.active
    ws.title = "方案清单"

    # 第 1 行：项目名称 —— 成果文件上的名字就取这里
    c = ws.cell(row=1, column=1, value="项目名称")
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEAD_FILL
    c.alignment = WRAP
    c.border = BORDER
    c2 = ws.cell(row=1, column=2, value=args.project_name or None)
    c2.fill = INPUT_FILL
    c2.alignment = LEFT
    c2.border = BORDER
    c2.font = Font(bold=True, size=12)
    for col in range(3, 6):
        cc = ws.cell(row=1, column=col)
        cc.border = BORDER
    c3 = ws.cell(row=1, column=3,
                 value="← 填这里。这个名字会印在《AHP权重表》《AHP方案评分结论》"
                       "《AHP权重与评分结论》三份成果文件上")
    c3.alignment = LEFT
    c3.font = Font(size=9, color="808080")
    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=5)

    # 第 2 行：说明；第 3 行起是表头
    ws.cell(row=2, column=1, value=(
        "本表要填两样：① 上面的项目名称；② 下表的方案名与说明。"
        "专家问卷封面上的项目名是占位的（问卷要匿名），不用你管。"
    )).alignment = LEFT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    HEAD_ROW = 3
    headers = ["序号", "方案编号", "方案名称", "方案说明", "是否参与评分"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=HEAD_ROW, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = WRAP
        c.border = BORDER

    presets = SUGGESTED if args.preset else []
    for idx in range(BLANK_ROWS):
        r = HEAD_ROW + 1 + idx
        c = ws.cell(row=r, column=1, value=idx + 1)
        c.alignment = WRAP
        c.border = BORDER
        c2 = ws.cell(row=r, column=2, value=f"S{idx + 1}")
        c2.alignment = WRAP
        c2.border = BORDER
        # 预填模式：给建议方案名（用户改名）；空白模式：留空
        p_name, p_note = presets[idx] if idx < len(presets) else ("", "")
        for col, val in ((3, p_name), (4, p_note), (5, "是" if p_name else "")):
            cc = ws.cell(row=r, column=col, value=val or None)
            cc.alignment = LEFT if col == 4 else WRAP
            cc.border = BORDER
            cc.fill = INPUT_FILL

    tip_row = HEAD_ROW + BLANK_ROWS + 2
    ws.cell(row=tip_row, column=1, value=(
        "说明：黄色三列（方案名称 / 方案说明 / 是否参与评分）是给你填的；"
        "方案编号默认 S1/S2/S3…，第 08 步按本表预填评分表的方案列。"
        "行数不够就在下面加行，删行也可以 —— 评分表按实际行数生成。"
    )).alignment = LEFT
    ws.merge_cells(start_row=tip_row, start_column=1, end_row=tip_row, end_column=5)
    ws.cell(row=tip_row + 1, column=1, value=(
        "提示：不知道该定哪几个方案？翻到「候选方案参考」页看常见思路 —— "
        "那是给你灵感，不是让你照抄。真实方案得按项目实际定。"
    )).alignment = LEFT
    ws.merge_cells(start_row=tip_row + 1, start_column=1, end_row=tip_row + 1, end_column=5)

    for col, w in enumerate([6, 12, 26, 56, 14], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    # ---------------- 候选方案参考 ----------------
    ws1 = wb.create_sheet("候选方案参考")
    ws1.cell(row=1, column=1, value="候选方案参考（按维度分类，供挑选，不是让你照抄）").font = Font(bold=True, size=13)
    ws1.cell(row=2, column=1, value=(
        "真实项目的方案通常只从 1~2 个维度里各取一个，交叉组合而成。"
        "挑中哪个，就抄到「方案清单」页并改成你项目的叫法。"
    )).alignment = LEFT
    ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    for col, w in enumerate([22, 30, 62], start=1):
        ws1.column_dimensions[get_column_letter(col)].width = w
    r = 4
    for dim, items in REFERENCE:
        c = ws1.cell(row=r, column=1, value=dim)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
        c.alignment = WRAP
        c.border = BORDER
        for col in (2, 3):
            ws1.cell(row=r, column=col).fill = SUB_FILL
            ws1.cell(row=r, column=col).border = BORDER
        r += 1
        for nm, desc in items:
            ws1.cell(row=r, column=1, value="").border = BORDER
            cn = ws1.cell(row=r, column=2, value=nm)
            cn.alignment = WRAP
            cn.border = BORDER
            cd = ws1.cell(row=r, column=3, value=desc)
            cd.alignment = LEFT
            cd.border = BORDER
            r += 1
        r += 1

    # ---------------- 填写说明 ----------------
    ws0 = wb.create_sheet("填写说明")
    guide = [
        "AHP 候选改造方案清单（10 分制方案的登记表）",
        "",
        "【你在整个流程的哪一步】",
        "① 定方案（你在这里）→ ② 发问卷（你→专家）→ ③ 回收问卷（专家→你）",
        "   → ④ 打分（你）→ ⑤ 出成果（系统自动算）",
        "",
        "本表负责第 ① 步。这一步做完，第 08 步才能生成方案评分表。",
        "",
        "【本表要填两样】",
        "1. B1 格的「项目名称」—— 会印在《AHP权重表》《AHP方案评分结论》",
        "   《AHP权重与评分结论》三份成果文件上。不填的话第 09 步会拦住不让交付。",
        "2. 「方案清单」页的方案名称、方案说明、是否参与评分。",
        "",
        "【做完长什么样】",
        "B1 有项目名 + 方案名称列有 2~4 行有字。",
        "填完重跑一次 开工.py，本表状态会变成「已完成」。",
        "",
        "【怎么填】",
        "· 方案编号：默认 S1/S2/S3…，一般不用改。",
        "· 方案名称：黄色格，填你项目的真实方案名（如「片区整体重建」「老旧小区综合整治」）。",
        "· 方案说明：一句话说清「改什么、怎么改」，建议带上投入量级、见效周期、对居民影响。",
        "· 是否参与评分：填「是」或「否」。标「否」的不进最终排名（保留在清单里备查）。",
        "· 方案数量不限：预置 6 行不够就在下面加行；删行也可以。第 08 步按实际行数生成评分表。",
        "",
        "【方案该有几个、怎么选】",
        "1. 数量：一般 2~4 个。少于 2 个没法比；多于 4~5 个，评分人负担重、结论不易收敛。",
        "2. 差异：方案之间要有实质差异——改法不同（拆还是留、推倒还是分期），",
        "   而不是面积大小不同。否则比出来的名次没有决策意义，等于白比。",
        "3. 「候选方案参考」页给的是常见类型的思路清单，不是让你照抄。",
        "   真实项目的方案组合通常只从其中取 1~2 个维度交叉，需要你按项目实际定。",
        "",
        "【配套文件】",
        "· 专家问卷：输入文件/02-专家问卷（待发放）→ 专家填完放回（已回收）。",
        "· 方案评分表：填好方案、算完权重后自动生成在 输入文件/03-方案打分表/AHP方案评分表.xlsx。",
        "· 全流程步骤与状态：工作区根目录 放料清单.md（每次运行都刷新）。",
        "· 逐格填写说明：工作区根目录 怎么填.md。",
    ]
    for i, text in enumerate(guide, start=1):
        ws0.cell(row=i, column=1, value=text)
    ws0.cell(row=1, column=1).font = Font(size=14, bold=True)
    ws0.column_dimensions["A"].width = 105

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"已生成方案清单：{out}（空白待填 {BLANK_ROWS} 行 + 候选方案参考页）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
