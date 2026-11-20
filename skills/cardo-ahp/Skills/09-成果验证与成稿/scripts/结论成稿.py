#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""结论成稿.py -- 把权重/评分/验证结果汇编成交付结论（docx 进成果文件，md 进过程文件）。

边界约定：成果文件只放常规 Office；docx 是交付正稿，md 只是过程留痕（供 Agent 快速读）。

用法：
  python 结论成稿.py --weights <权重表.json> --scores <评分表.json> \
      --verification <独立验证结果.json> --out-docx <…docx> --out-md <…md> \
      [--aggregation <群体矩阵.json>] [--project-name 名称]
退出码：0 成功；2 输入错。

口径纪律（2026-10-30 修）：正文中的**专家人数、λmax、CI、CR 一律取本次运行的实际值**，
不得写死「5 位专家」「λmax=4、CI=CR=0」「本次专家判断一致（CR=0）」——
那些是早期跑通链路时的样例值，写死会在真实项目（CR>0、专家数不同）时说错话。
--aggregation 可选：不传时专家数退化为未知、CR 从权重表一致性段取，并标注来源。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def read_json(p: str) -> dict:
    return json.loads(Path(p).read_text(encoding="utf-8-sig"))


def style_doc(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for name, size in (("Title", 20), ("Heading 1", 14), ("Heading 2", 12)):
        st = doc.styles[name]
        st.font.name = "Microsoft YaHei"
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)


def _row_flags(row, header: bool = False) -> None:
    """禁止表格行跨页断开（否则重跑/重排时页尾留半行残行）；表头行跨页重复。"""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    if header:
        trPr.append(OxmlElement("w:tblHeader"))


def add_table(doc: Document, headers: list[str], rows: list[list], widths: list[float] | None = None) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    _row_flags(table.rows[0], header=True)
    for i, h in enumerate(headers):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(str(h))
        run.bold = True
        run.font.size = Pt(11)
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in rows:
        row = table.add_row()
        _row_flags(row)
        cells = row.cells
        for i, v in enumerate(r):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run("" if v is None else str(v))
            run.font.size = Pt(11)
    # 整表 together：除末行外都「与下段同页」，避免分页后另一页只剩表头 + 一行残行。
    for row in table.rows[:-1]:
        for cell in row.cells:
            for par in cell.paragraphs:
                par.paragraph_format.keep_with_next = True
    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()


SYNTHETIC_NOTE = ("本次为合成问卷演示（5 份模拟问卷、3 个候选方案），"
                  "用于验证方法链路与计算正确性，结论不代表真实专家判断。")

# 结论节的适用范围：不是补免责声明，而是把方法产出与适用边界写进结论本身。
# 必须与结论同屏（紧跟章标题、正文同字号、keep_with_next），不做脚注/小字/文末附注。
SCOPE_NOTE = ("本节适用范围：下面的排序与评级，由本次合成问卷（5 份模拟问卷、3 个候选方案）"
              "产生的权重与指标得分推出，用途是验证「判断矩阵 → 权重 → 一致性检验 → 方案评分」"
              "这条方法链路能否跑通、算法是否正确，属方法输出，不是对片区更新的投资结论；"
              "「优秀／良好／一般／不可行」是 10 分制档位名称，只表示本次假设数据下的相对位置，"
              "不对任何候选方案或该片区作可行性判断。真实项目应以实际回收的专家问卷重算后再定取舍。")

# 演示口径下「投资建议」列降级为「方法输出」：从投资决策改为方案间相对比较。
DEMO_ACTION = {"优秀": "优秀档（[9,10]）",
               "良好": "良好档（[7,9)）",
               "一般": "一般档（[6,7)）",
               "不可行": "不可行档（[0,6)）"}

_CN_NUM = "一二三四五六七八九十"


def _sn(n: int) -> str:
    """章号（1 → 「一、」）：章数随是否含内部自检章变化，不写死。"""
    return f"{_CN_NUM[n - 1]}、"


def _ordinal(idx: int, mode: str) -> str:
    """条目序号（公文层次序数用文字，不用项目符号）。

    mode="num"  → 「1. 」「2. 」…（第三条层次）
    mode="none" → 不加序号（条目自带「结论：」「边界 1：」等文字标签时，
                  再加序号会变成「1. 边界 1：」双编号）
    """
    if mode != "num":
        return ""
    return f"{idx}. "


def ratio(done, total) -> str:
    """通过项/总项；分母为 0 时写「不适用」，不写 0/0。"""
    try:
        total = int(total)
    except (TypeError, ValueError):
        total = 0
    return f"{done}/{total}" if total else "不适用"


_MATRIX_NAMES = {"criteria": "准则层", "current": "现状情况", "economy": "投资经济",
                 "compliance": "规划合规", "risk": "风险可控"}


def polish(text: str, scheme_names: dict | None = None) -> str:
    """成品版措辞：把验证明细里的机读标识换成可读名称、长小数收敛到 4 位。

    独立验证的明细是给复核人看的机读串（S1、current、maxΔ、五位以上小数），
    直接印进对外报告像调试输出。这一层只做措辞与精度，不动任何判定结果。
    """
    out = str(text)
    for sid, name in (scheme_names or {}).items():
        out = re.sub(rf"(?<![A-Za-z0-9]){re.escape(str(sid))}(?![A-Za-z0-9])", str(name), out)
    for mid, name in _MATRIX_NAMES.items():
        out = out.replace(f"{mid}:", f"{name}:")
    out = re.sub(r"=(-?\d+\.\d{6,})", lambda m: "=" + f"{float(m.group(1)):.4f}", out)
    for old, new in (("maxΔ", "最大偏差"), ("CR_max", "最大 CR"), ("首跑自举", "首跑建立"),
                     ("缺：无", "无缺项"), ("缺分方案：无", "各方案均已出分"),
                     ("独立=", "独立复算="), ("引擎=", "引擎=")):
        out = out.replace(old, new)
    return out


def build(args, annotated: bool = False) -> tuple[dict, list[dict]]:
    """返回 (结构化数据, 渲染块列表)；md 与 docx 共用同一份块。

    annotated=True 出溯源版：多一段「溯源说明」，正文保留来源行与字段级依据；
    成品版只留业主需要的方法、数据与结论，演示口径只在方法说明里交代一次。
    """
    w = read_json(args.weights)
    s = read_json(args.scores)
    v = read_json(args.verification)
    agg = read_json(args.aggregation) if getattr(args, "aggregation", None) else None
    crits = w.get("criteria", [])
    inds = w.get("indicators", [])
    cons = w.get("consistency", {})
    cn = {c["id"]: c["name"] for c in crits}
    scheme_names = {x.get("id"): x.get("name") for x in (s.get("schemes") or [])
                    if x.get("id") and x.get("name")}

    # ---- 实际值（不写死）----
    matrices = cons.get("matrices") or []
    expert_count = None
    if agg:
        expert_count = ((agg.get("aggregation") or {}).get("expertCount")
                        or len(agg.get("experts") or []) or None)
    cr_max = max((float(m.get("cr") or 0.0) for m in matrices), default=0.0)
    lam_max = max((float(m.get("lambdaMax") or 0.0) for m in matrices), default=0.0)
    ci_max = max((float(m.get("ci") or 0.0) for m in matrices), default=0.0)
    expert_text = (f"{expert_count} 位专家等权" if expert_count
                   else "多位专家等权")
    if cr_max <= 1e-12:
        cr_text = (f"{len(matrices)} 个矩阵 λmax={lam_max:.4f}、CI=CR=0"
                   f"（专家判断完全一致）")
    else:
        cr_text = (f"{len(matrices)} 个矩阵中最大 λmax={lam_max:.4f}、CI={ci_max:.6f}、"
                   f"CR={cr_max:.6f}（CR<{cons.get('threshold')} 通过 "
                   f"{ratio(cons.get('passCount'), cons.get('total'))}）")

    # 演示口径只认本次运行的显式声明（--demo / 工作区参数 demo:true），不认输入文件内容：
    # 输入表里出现「演示/示例」一类字样不等于本次交付件是演示件，成品页不得据此自述合成。
    demo = bool(getattr(args, "demo", False))
    # 成品版不出「六、独立验证」，后续章号整体前移一位；溯源版保留全部章。
    concl_no = 7 if annotated else 6
    list_no = 9 if annotated else concl_no + 1

    def action_of(item) -> str:
        """评级对应的动作语：演示口径写方案间相对比较，正式口径才写投资建议。"""
        rating = item.get("rating") or {}
        if demo:
            return DEMO_ACTION.get(rating.get("label"), "方法输出")
        return rating.get("action")

    blocks: list[dict] = [
        {"type": "h1", "text": "一、方法与口径"},
    ]
    if demo:
        # 口径只交代一次，放在方法说明首段；成品版其余部分按正式报告写。
        blocks.append({"type": "p", "text": SYNTHETIC_NOTE, "note": "synthetic"})
    blocks += [
        {"type": "p", "text": "层次结构：目标层（投资决策最优方案）— 准则层 4 项 — 指标层 16 项，与标准模板逐项对应。"},
        {"type": "p", "text": "判断与聚合：专家只填判断矩阵上三角，下三角按 a_ji = 1 / a_ij 自动互反；"
                              f"{('回收 ' + str(expert_count) + ' 份、全部有效，') if expert_count else '全部有效回收，'}"
                              f"{expert_text}，逐元素几何平均形成群体判断矩阵。"},
        {"type": "p", "text": "标度：准则层 4 项之间用 1–9 标度（1 同等重要、3 略重要、5 明显重要、"
                              "7 强烈重要、9 极端重要，2/4/6/8 取相邻中间值）；指标层同一准则下 4 项之间用"
                              " 0.1–10 比值标度，输入为 i 相对 j 的重要性倍数（0.1 表示 i 远不如 j，1 表示同等，"
                              "10 表示 i 远强于 j；a_ji = 1 / a_ij，矩阵元素可大于 1），取值粒度 0.1，"
                              "适配 4 项指标内的细微差别；一致性检验口径与阈值不变。"},
        {"type": "p", "text": "求权与检验：几何平均行积法与特征值法双算法计算局部/综合权重；CR=CI/RI（n=4 时 RI=0.90），CR<0.1 通过。"},
        {"type": "p", "text": "方案评分：16 项指标 10 分制人工评分并逐项注明依据，综合得分=Σ(得分×综合权重)；"
                              "评级 [9,10] 优秀 / [7,9) 良好 / [6,7) 一般 / [0,6) 不可行。"},

        {"type": "h1", "text": "二、准则层权重"},
        {"type": "p", "text": "准则层 4 项的权重见表1。", "keep": True},
        {"type": "table",
         "caption": "表1　准则层权重",
         "headers": ["准则", "权重"],
         "rows": [[c["name"], f"{c['weight']:.6f}"] for c in crits],
         "widths": [6, 4]},

        {"type": "h1", "text": "三、指标层权重（局部 / 综合）"},
        {"type": "p", "text": "指标层 16 项的局部权重与综合权重见表2。", "keep": True},
        {"type": "table",
         "caption": "表2　指标层权重（局部 / 综合）",
         "headers": ["准则", "指标", "局部权重", "综合权重"],
         "rows": [[cn.get(i["criterionId"], i["criterionId"]), i["name"],
                   f"{i['localWeight']:.6f}", f"{i['globalWeight']:.6f}"] for i in inds],
         "widths": [3.2, 6.5, 3, 3]},

        {"type": "h1", "text": "四、一致性检验"},
        {"type": "p", "text": f"{cr_text}。逐矩阵检验结果见表3。", "keep": True},
        {"type": "table",
         "caption": "表3　一致性检验结果",
         "headers": ["矩阵", "n", "λmax", "CI", "RI", "CR", "判定"],
         "rows": [[(_MATRIX_NAMES.get(m["id"], m["id"]) + (f"（{m['id']}）" if annotated else "")),
                   m["n"], f"{m['lambdaMax']:.4f}", f"{m['ci']:.6f}",
                   f"{m['ri']:.2f}", f"{m['cr']:.6f}", "通过" if m["pass"] else "不通过"]
                  for m in matrices],
         "widths": [4.8, 1.0, 2.2, 2.0, 1.4, 2.0, 1.6]},
        {"type": "p", "text": "5 个判断矩阵的原始相对重要度与逐格互反值见同包《AHP权重表》的「判断矩阵」页，"
                              "λmax、CI、CR 可在该页按本节口径直接复核；16 项指标的逐项得分与对综合得分的"
                              "贡献见《AHP方案评分结论》的「指标得分明细」页。"},

        {"type": "h1", "text": "五、候选方案评分与排序"},
        {"type": "p", "text": ("本节按 16 项指标加权求和，给出三个候选方案在本次判断矩阵下的相对排序；"
                                "末列由评级档位直接推出，只说明该方案落在哪一档。" if demo else
                                "本节按 16 项指标加权求和，给出候选方案的相对排序与投资建议。"),
         "keep": True},
        {"type": "p", "text": "候选方案综合得分与排序见表4。", "keep": True},
        {"type": "table",
         "caption": "表4　候选方案综合得分与排序",
         "headers": ["名次", "方案", "综合得分", "评级", "方法输出" if demo else "投资建议"],
         "rows": [[r["order"], r["name"], f"{r['composite']:.2f}",
                   (r.get("rating") or {}).get("label"), action_of(r)]
                  for r in s.get("ranking", [])],
         "widths": [1.4, 4, 2.4, 2, 5]},
    ]
    drivers = (s.get("conclusion") or {}).get("topDrivers") or []
    if drivers:
        top = (s.get("ranking") or [{}])[0].get("name")
        top_scheme = next((x for x in (s.get("schemes") or []) if x.get("name") == top),
                          (s.get("schemes") or [{}])[0])
        weight_by_id = {it.get("id"): it.get("globalWeight")
                        for it in (top_scheme.get("scores") or [])}
        items = []
        for d in drivers:
            weight = d.get("weight")
            if weight is None:
                weight = weight_by_id.get(d.get("id"))
            if weight is None and d.get("score"):
                weight = float(d["contribution"]) / float(d["score"])
            items.append(f"{d['name']}：得分 {d['score']} × 权重 {float(weight or 0):.6f}"
                         f" = 贡献 {float(d.get('contribution') or 0):.6f}")
        blocks.append({"type": "p", "text": f"{top}的主要得分贡献："})
        blocks.append({"type": "bullets", "items": items, "ord": "num"})

    bstatus = (v.get("baselineInfo") or {}).get("status")
    baseline_text = {"written": "本次复算结果（首跑自举基线）",
                     "skipped": "未启用基线对账",
                     "checked": "既有基线"}.get(bstatus, "本次复算结果")
    if annotated and v.get("baseline"):
        baseline_text += f"：{v.get('baseline')}"
    # 内部自检整章只进溯源版：基线/引擎/ranking/跳过项属于复核过程，业主稿放这些
    # 既占篇幅又会被读成「交付方自证」。p8 的 CR=0.001 与 7 硬断行随本章搬走一并消失。
    if annotated:
        blocks += [
            {"type": "h1", "text": "六、独立验证"},
            {"type": "p", "text": f"基线：{baseline_text}"},
            {"type": "p", "text": f"结果：通过 {v.get('passed')} 项 · 失败 {v.get('failed')} 项"
                                  f" · 跳过 {v.get('skipped')} 项。验证项明细见表5。", "keep": True},
            {"type": "table",
             "caption": "表5　独立验证结果",
             "headers": ["验证项", "结果", "明细"],
             "rows": [[c["item"], {"PASS": "通过", "FAIL": "失败", "SKIP": "跳过"}.get(
                 c["status"], c["status"]), c["detail"]] for c in v.get("checks", [])],
             "widths": [5.6, 1.4, 9.5]},
        ]
    bounds = [
        f"结论：{(s.get('conclusion') or {}).get('summary', '').replace(chr(10), ' ')}",
        f"边界 1：本次判断矩阵的最大 CR={cr_max:.6f}"
        + ("（判断矩阵完全一致）。专家意见的离散程度不作为权重取舍依据；CR≥0.1 时走一致性校准并由人工批准。"
           if cr_max <= 1e-12 else
           "（矩阵不一致程度在允许范围内，全部通过）。CR≥0.1 时须走一致性校准并由人工批准。"),
        "边界 2：方案评分按 10 分制锚点逐项赋分，赋分口径见第一节与第五节；"
        "重大投资决策还须经现状勘察、财务测算与专家研判复核。",
        "边界 3：16 项指标按通用口径全部适用；如项目方向明确要求裁剪指标，"
        "须重新归一权重后重算，不得直接沿用本报告权重。",
    ]
    if annotated:
        bounds.append("边界 4：权重表导出依赖固定模板版式；更换模板需先核对写入格位。")
    blocks.append({"type": "h1", "text": f"{_sn(concl_no)}结论与边界"})
    if demo:
        # 适用范围与结论同屏：紧跟章标题、正文同字号、绑住下方首条结论，不做脚注/小字/文末附注。
        blocks.append({"type": "p", "text": SCOPE_NOTE, "keep": True})
    blocks.append({"type": "bullets", "items": bounds, "ord": "none"})
    if annotated:
        blocks += [
            {"type": "h1", "text": f"{_sn(concl_no + 1)}溯源说明"},
            {"type": "p", "text": "本稿每个数值均可回指到 依据数据/ 下的源 JSON：权重、λmax、CI、CR 取"
                                  " 权重表.json，综合得分与贡献取 评分表.json，专家人数与群体矩阵取"
                                  " 群体矩阵.json；逐项对应关系见《AHP依据对照表》。"},
            {"type": "p", "text": "过程文件/01..09-*/ 保留全部机读 JSON 与校验报告，供内部复核；成品版不引用内部件。"},
        ]
    blocks += [
        {"type": "h1", "text": f"{_sn(list_no)}产物清单"},
        {"type": "bullets", "items": [
            "《AHP权重表》（AHP权重表_成品版.xlsx）：准则层/指标层权重、一致性检验、判断矩阵与方法说明，静态数值、不含公式",
            "《AHP方案评分结论》（AHP方案评分结论_成品版.xlsx）：排序、评级、指标得分与贡献明细",
            "《AHP权重与评分结论》（AHP权重与评分结论_成品版.pdf）：本报告正式交付版（docx 为可编辑稿）",
        ] + ([
            "成品版/AHP权重表/AHP权重表_成品版.xlsx、成品版/AHP方案评分结论/AHP方案评分结论_成品版.xlsx",
            "成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.docx 与同名 PDF",
            "内部件/AHP权重表/AHP权重表_溯源版.xlsx、内部件/AHP方案评分结论/AHP方案评分结论_溯源版.xlsx"
            "（含「内部核对」页：源 JSON 路径与 sha256）",
            "内部件/AHP权重与评分结论/AHP权重与评分结论_溯源版.docx（本件溯源稿）",
            "内部件/AHP依据对照表/AHP依据对照表.xlsx（成稿数值 → 源 JSON 键的逐项回指表）",
            "内部件/依据数据/（权重表.json、评分表.json、群体矩阵.json、矩阵集.json，随包带走仍可回指）",
            "过程文件/01..09-*/（按技能编号留痕；含全部机读 JSON 与校验报告）",
        ] if annotated else []), "ord": "num"},
    ]
    data = {"criteria": crits, "indicators": inds, "cons": cons, "ranking": s.get("ranking", [])}
    return data, blocks


def render_md(title: str, sub: str, blocks: list[dict]) -> str:
    lines = [f"# {title}", "", f"> {sub}", ""]
    for b in blocks:
        if b["type"] == "h1":
            lines += [f"## {b['text']}", ""]
        elif b["type"] == "p":
            lines += [b["text"], ""]
        elif b["type"] == "bullets":
            mode = b.get("ord", "none")
            lines += [f"{_ordinal(i, mode)}{x}" for i, x in enumerate(b["items"], 1)] + [""]
        elif b["type"] == "table":
            if b.get("caption"):
                lines += [f"**{b['caption']}**", ""]
            lines.append("| " + " | ".join(b["headers"]) + " |")
            lines.append("|" + "---|" * len(b["headers"]))
            for r in b["rows"]:
                lines.append("| " + " | ".join("" if x is None else str(x) for x in r) + " |")
            lines.append("")
    return "\n".join(lines)


def _add_field(par, instr: str, size: float = 11) -> None:
    """在段落里插入 Word 域（PAGE / NUMPAGES），打印时由 Word 自动算页码。

    用 fldChar + instrText 的复合域，而不是 w:fldSimple：前者是 Word 与各类
    转换器都认的标准域结构（instrText 里能查到 PAGE/NUMPAGES），后者在部分
    渲染链里会被当成静态文本。
    域里的 run 必须自带字号：否则 Word 用默认 11pt，同一行页码会比前后
    「第 / 页」大一号（实测 11.05 vs 10.45）。
    """
    begin_run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run.append(begin)
    par._p.append(begin_run)

    instr_run = OxmlElement("w:r")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    instr_run.append(instr_el)
    par._p.append(instr_run)

    sep_run = OxmlElement("w:r")
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    sep_run.append(sep)
    par._p.append(sep_run)

    text_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(round(size * 2))))
    rPr.append(sz)
    text_run.append(rPr)
    text = OxmlElement("w:t")
    text.text = "1"
    text_run.append(text)
    par._p.append(text_run)

    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)
    par._p.append(end_run)


UNIT_NAME = "Cardo 城市更新 AHP 项目组"


def add_footer(doc: Document, tag: str = "") -> None:
    """页脚居中：页码 + 总页数；溯源版在页码前标「内部复核稿」（与正文同为 11pt）。"""
    par = doc.sections[0].footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.text = ""
    if tag:
        run = par.add_run(f"{tag}　")
        run.font.size = Pt(11)
    for piece, field in (("第 ", None), ("", " PAGE "), (" 页 / 共 ", None),
                         ("", " NUMPAGES "), (" 页", None)):
        if field:
            _add_field(par, field)
        else:
            run = par.add_run(piece)
            run.font.size = Pt(11)


def _usable_cm(doc: Document) -> float:
    """正文可用宽度（cm）：页宽减左右页边距，留 0.15 cm 余量。"""
    sec = doc.sections[0]
    return (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0 - 0.15


def heading_new_page(doc: Document, text: str, level: int = 1):
    """另起一页的标题：分页符放在标题段首位，不用空段落承载。

    空段落 + 分页符在正文排满时会自己翻页，多出一张白纸；把分页符塞进标题段，
    既能让 Word 换页，也能让按 XML 依次渲染的查看器在这里断页。
    """
    p = doc.add_heading("", level=level)
    p.add_run().add_break(WD_BREAK.PAGE)
    p.add_run(text)
    return p


def add_cover(doc: Document, project: str, title: str, subtitle: str,
              date_text: str, version: str = "V1.0") -> None:
    """封面：项目名、成果名、编制单位、日期、版本。封面后由「目录」标题分页。"""
    for _ in range(3):
        doc.add_paragraph()
    for text, size, color in ((project, 20, None), (title, 26, None),
                              (subtitle, 12, RGBColor(0x55, 0x55, 0x55))):
        line = doc.add_paragraph()
        line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = line.add_run(text)
        run.bold = size >= 20
        run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = color
    for _ in range(6):
        doc.add_paragraph()
    rows = [["项目名称", project], ["成果名称", title], ["编制单位", UNIT_NAME],
            ["编制日期", date_text], ["版本", version]]
    add_table(doc, ["项", "内容"], rows, [3.5, 11])


def _toc_field(first_par, last_par) -> None:
    """把静态条目包进 TOC 域：Word 里打开仍可按改版刷新页码。

    无 Office 环境算不出分页，故缓存内容不写死页码；实际页码由 pdf_render
    按渲染分页回填成静态文本（见 fill_toc_pages）。
    """
    p = first_par._p
    pos = 1 if p.find(qn("w:pPr")) is not None else 0
    begin = OxmlElement("w:r")
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    begin.append(fld)
    instr = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = ' TOC \\o "1-1" \\h \\z \\u '
    instr.append(it)
    sep = OxmlElement("w:r")
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    sep.append(fld2)
    for i, el in enumerate((begin, instr, sep)):
        p.insert(pos + i, el)
    end_run = last_par.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)


def disable_update_fields(doc: Document) -> None:
    """不要求打开即刷新域：目录页码是静态回填的，Word 刷新会把「目录」自列一条，
    且重排后的页码会与 PDF 不一致。页脚 PAGE/NUMPAGES 仍是真的域，随排版自算。"""
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "false")


def _hide_from_toc(par) -> None:
    """把段落的大纲级别设为正文级：样式仍算一级标题（本工具链靠样式识别），
    Word 的 TOC 域却不会把「目录」自己列成一条。"""
    pPr = par._p.get_or_add_pPr()
    lvl = OxmlElement("w:outlineLvl")
    lvl.set(qn("w:val"), "9")
    pPr.append(lvl)


def add_toc(doc: Document, entries: list, project: str, generated: str) -> None:
    """目录页（章名 + 右侧点线 + 页码）。

    页码不靠 Word 域：成稿后由 pdf_render 按实际渲染分页回填成静态文本，
    不更新域的查看路径（另存 PDF、预览）也能读到；TOC 域保留供 Word 刷新。
    """
    toc_head = heading_new_page(doc, "目录", level=1)
    _hide_from_toc(toc_head)
    note = doc.add_paragraph()
    note.paragraph_format.space_after = Pt(6)
    run = note.add_run("（目录页码按 PDF 成品的实际分页写入；在 Word 中重新排版"
                       "时页码可能微调，以 PDF 版为准。）")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    usable = _usable_cm(doc)
    first = last = None
    for text in entries:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(usable), WD_TAB_ALIGNMENT.RIGHT,
                                                  WD_TAB_LEADER.DOTS)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10.5)
        if first is None:
            first = p
        last = p
    if first is not None:
        _toc_field(first, last)
    # 不再补「项目：… 生成时间：…」灰行：目录区由 pdf_render 按下一个一级标题切回正文，
    # 项目与时间已在上游封面交代；这行游离在目录末会与正文首段挤在一起。


def _is_heading_par(p) -> bool:
    try:
        return bool(re.match(r"^(?:Heading|标题)\s*\d+$", p.style.name or ""))
    except (AttributeError, ValueError):
        return False


_CN_DIGITS = "一二三四五六七八九十"


def sec_key(text: str) -> str:
    """章节号（「一、方法与口径」→「一」；「1.1 x」→「1.1」）。"""
    match = re.match(rf"^\s*(\d+(?:\.\d+)?|[{_CN_DIGITS}]+)(?:[、.\s]|$)", text or "")
    return match.group(1) if match else ""


def toc_entries(blocks: list[dict]) -> list:
    return [b["text"] for b in blocks if b["type"] == "h1"]


def fill_toc_pages(docx_path: str, pages: dict) -> int:
    """把 PDF 的实际页码写进 docx 目录（静态文本 + 点线制表位）。返回回填条数。"""
    doc = Document(docx_path)
    in_toc, filled = False, 0
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if _is_heading_par(p):
            if text == "目录":
                in_toc = True
                continue
            if in_toc:
                break
            continue
        if not in_toc or not text:
            continue
        if text.startswith("项目："):
            break
        if text.startswith(("（", "(")):
            continue
        key = sec_key(text)
        page = pages.get(key)
        if key and page:
            p.add_run(f"\t{page}")
            filled += 1
    if filled:
        doc.save(docx_path)
    return filled


def reset_toc_pages(docx_path: str) -> None:
    """清掉目录里已回填的页码（二次回填前用，避免叠加成「3\t4」）。"""
    doc = Document(docx_path)
    in_toc = False
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if _is_heading_par(p):
            if text == "目录":
                in_toc = True
                continue
            if in_toc:
                break
            continue
        if not in_toc:
            continue
        for run in list(p.runs):
            if run.text.startswith("\t"):
                run._r.getparent().remove(run._r)
    doc.save(docx_path)


def render_docx(path: Path, project: str, title: str, subtitle: str, date_text: str,
                blocks: list[dict], tag: str = "") -> int:
    """封面（独立成页）→ 目录 → 正文（每章另起一页）。返回目录条目数。"""
    doc = Document()
    style_doc(doc)
    # python-docx 默认模板是 Letter（12240×15840 twips）；交付要求 A4，
    # 与 PDF 链（A4 595×842 pt）同源，否则 Word 打印版心与烘进目录的页码不一致。
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    disable_update_fields(doc)
    add_cover(doc, project, title, subtitle, date_text)
    entries = toc_entries(blocks)
    add_toc(doc, entries, project, date_text)
    add_footer(doc, tag)
    for b in blocks:
        if b["type"] == "h1":
            heading_new_page(doc, b["text"], level=1)
        elif b["type"] == "p":
            par = doc.add_paragraph(b["text"])
            if b.get("keep"):
                par.paragraph_format.keep_with_next = True
        elif b["type"] == "bullets":
            # 公文层次序数用文字，不用项目符号（List Bullet 是符号列表，验收判不合格）
            mode = b.get("ord", "none")
            for i, x in enumerate(b["items"], 1):
                doc.add_paragraph(f"{_ordinal(i, mode)}{x}")
        elif b["type"] == "table":
            if b.get("caption"):
                cap_par = doc.add_paragraph()
                cap_run = cap_par.add_run(b["caption"])
                cap_run.bold = True
                cap_run.font.size = Pt(10.5)
                cap_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap_par.paragraph_format.space_before = Pt(6)
                cap_par.paragraph_format.keep_with_next = True
            add_table(doc, b["headers"], b["rows"], b.get("widths"))
    path.parent.mkdir(parents=True, exist_ok=True)
    # 文件属性不留 python-docx 默认值（author=python-docx、created=2013-12-23 一类）
    props = doc.core_properties
    props.author = "Cardo 城市更新 AHP 项目组"
    props.last_modified_by = props.author
    props.title = title
    props.subject = project
    props.category = "AHP 成果文件"
    props.comments = ""
    stamp = datetime.now()
    props.created = stamp
    props.modified = stamp
    doc.save(path)
    return len(entries)


def main() -> int:
    ap = argparse.ArgumentParser(description="AHP 结论成稿（docx 交付 + md 过程留痕）")
    ap.add_argument("--weights", required=True)
    ap.add_argument("--scores", required=True)
    ap.add_argument("--verification", required=True)
    ap.add_argument("--aggregation", default=None,
                    help="可选：群体矩阵 JSON（提供专家人数等实际值，避免成稿写死样例数）")
    ap.add_argument("--out-docx", required=True, help="成品版 docx（对外交付）")
    ap.add_argument("--out-docx-ann", default="",
                    help="可选：溯源版 docx（内部复核，多「溯源说明」一章）")
    ap.add_argument("--out-md", required=True, help="过程留痕 md（成品版内容）")
    ap.add_argument("--out-md-ann", default="", help="可选：溯源版内容的过程留痕 md")
    ap.add_argument("--project-name", default="某某片区改造项目")
    ap.add_argument("--demo", action="store_true",
                    help="演示样例：在方法说明里交代合成问卷口径，不改任何数值")
    args = ap.parse_args()

    for p in (args.weights, args.scores, args.verification):
        if not Path(p).is_file():
            print(f"[AHP-E-INPUT] 文件不存在：{p}", file=sys.stderr)
            return 2
    if args.aggregation and not Path(args.aggregation).is_file():
        print(f"[AHP-E-INPUT] 文件不存在：{args.aggregation}", file=sys.stderr)
        return 2

    title = "AHP 权重与评分结论"
    subtitle = "层次分析法（AHP）权重计算、一致性检验与候选方案比选"
    date_text = datetime.now().strftime("%Y-%m-%d")

    _, blocks = build(args, annotated=False)
    docx_path = Path(args.out_docx)
    entries = render_docx(docx_path, args.project_name.strip(), title, subtitle,
                          date_text, blocks, tag="")
    md_path = Path(args.out_md)
    md_path.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text(render_md(f"{args.project_name} · {title}",
                                 f"编制日期：{date_text}", blocks), encoding="utf-8")
    print(f"已生成成品版结论：{docx_path}（目录 {entries} 条）")

    if args.out_docx_ann:
        _, blocks_ann = build(args, annotated=True)
        ann_path = Path(args.out_docx_ann)
        entries_ann = render_docx(ann_path, args.project_name.strip(), title,
                                  subtitle + "（内部复核稿）", date_text, blocks_ann,
                                  tag="内部复核稿")
        print(f"已生成溯源版结论：{ann_path}（目录 {entries_ann} 条）")
        if args.out_md_ann:
            md_ann = Path(args.out_md_ann)
            md_ann.parent.mkdir(parents=True, exist_ok=True)
            md_ann.write_text(render_md(f"{args.project_name} · {title}（内部复核稿）",
                                        f"编制日期：{date_text}", blocks_ann),
                              encoding="utf-8")
    print(f"过程留痕：{md_path}")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
