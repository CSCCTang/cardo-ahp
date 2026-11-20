#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""依据对照表.py -- AHP 结论成稿的《依据对照表》：每个数都指回权重表/评分表的哪一格。

结论稿里出现的每个数值（准则权重、指标权重、λmax/CI/CR、方案综合得分、拉分项）
都登记成一行：给谁看的值 / 它出自哪个 JSON 的哪个键 / 是否复算通过。
另外做三条独立复算（不引用成稿的计算结果）：
  W-1 准则权重合计 = 1，各准则下指标局部权重合计 = 1
  W-2 评分表用的综合权重与权重表一致（不允许两套权重）
  S-1 方案综合得分 = Σ(指标得分 × 综合权重)

产物：
  成果文件/AHP依据对照表.xlsx
  过程文件/09-成果验证与成稿/依据对照表.json
退出码：0 通过；1 有不通过项；2 输入错。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TOL_W = 0.001
TOL_SCORE = 0.005
BAD_WORDS = ("合成演示", "非真实依据", "示例数据", "演示数据")
SOURCE_PREFIX = ("来源：", "来源:", "（数据来源", "(数据来源", "数据来源：", "数据来源:")
# 本表落在 内部件/AHP依据对照表/ 下，依据数据目录与它同级（内部件/依据数据/），
# 故指针要对本表目录上跳一层：../依据数据/权重表.json。
DATA_DIR_PREFIX = "../依据数据/"


def bundle_pointer(pointer: str) -> str:
    """指针统一写成随包相对路径：权重表.json# → ../依据数据/权重表.json#。

    依据数据/ 随交付包带走（内部件/依据数据/），离开本机也能按同一指针回指。
    """
    text = str(pointer)
    for name in ("权重表.json", "评分表.json", "矩阵集.json", "群体矩阵.json"):
        token = DATA_DIR_PREFIX + name
        if token in text:
            continue
        # 先剥掉旧前缀（依据数据/xxx.json），避免叠成 依据数据/../依据数据/xxx.json
        text = text.replace("依据数据/" + name, name)
        text = text.replace(name, token)
    return text


def read_json(p):
    return json.loads(Path(p).read_text(encoding="utf-8-sig"))


def pick(data, path):
    """按 `criteria[0].weight` 这类路径取值（只支持 名称[序号].名称…）。"""
    cur = data
    for part in path.split("."):
        name, _, tail = part.partition("[")
        if name:
            cur = cur[name]
        while tail:
            idx, _, tail = tail.partition("]")
            if not idx.isdigit():
                raise KeyError(idx)
            cur = cur[int(idx)]
            tail = tail.lstrip(".")
            if tail:
                name2, _, tail = tail.partition("[")
                cur = cur[name2]
    return cur


def same_text(text, value):
    text = str(text).strip()
    if isinstance(value, bool):
        return (text == "通过") == value
    if isinstance(value, (int, float)):
        try:
            return abs(float(text) - float(value)) <= 0.01
        except ValueError:
            return False
    return text == str(value).strip()


def verify_pointer(pointer, value, docs):
    """一行能否回指到源文件的对应键；返回 "" 表示走通，否则返回断在哪一环。

    派生行写法：`派生：文件A#键1 × 文件B#键2`，校验两个因子都取得到、且乘积=登记值。
    """
    if pointer.startswith("派生："):
        parts = [x.strip() for x in pointer[len("派生："):].split(" × ")]
        if len(parts) != 2:
            return "派生行应指向两个上游值（用 × 分隔）"
        factors = []
        for part in parts:
            reason, actual = resolve_pointer(part, docs)
            if reason:
                return f"派生因子断链：{reason}"
            try:
                factors.append(float(actual))
            except (TypeError, ValueError):
                return f"派生因子不是数字：{part}"
        try:
            recorded = float(value)
        except (TypeError, ValueError):
            return f"派生值不是数字：{value}"
        product = factors[0] * factors[1]
        if abs(product - recorded) > max(0.0000005, abs(recorded) * 0.001):
            return f"派生复算 {factors[0]:.6f}×{factors[1]:.6f}={product:.6f} ≠ 登记 {recorded:.6f}"
        return ""
    if "#" not in pointer:
        return "依据未写「文件#键」"
    fname, path = pointer.split("#", 1)
    path = path.split("（")[0].strip()
    if "[]" in path or "criterionId=" in path:
        return ""                      # 独立求和行：值由脚本自算，不指向单格
    data = docs.get(fname.strip().split("/")[-1])
    if data is None:
        return f"文件不在本次输入里：{fname}"

    def one(sub_path, sub_value):
        try:
            actual = pick(data, sub_path)
        except (KeyError, IndexError, TypeError) as exc:
            return f"键取不到：{sub_path}（{exc}）"
        return "" if same_text(sub_value, actual) else f"登记 {sub_value} ≠ 原文 {actual}"

    parts = path.split("/")
    if len(parts) == 2:              # 双键行：局部权重 / 综合权重（同一指标对象下的两个字段）
        halves = [x.strip() for x in str(value).split("/")]
        if len(halves) != 2:
            return f"登记 {value} 与双键对不上"
        if "." not in parts[1] and "[" not in parts[1]:
            parent = parts[0].rpartition(".")[0]
            parts[1] = f"{parent}.{parts[1]}" if parent else parts[1]
        for sub, sub_value in zip(parts, halves):
            reason = one(sub, sub_value)
            if reason:
                return reason
        return ""
    return one(path, value)


def resolve_pointer(pointer, docs):
    """取 `文件#键` 的实际值；返回 (断链原因, 值)。"""
    if "#" not in pointer:
        return "依据未写「文件#键」", None
    fname, path = pointer.split("#", 1)
    path = path.split("（")[0].strip()
    data = docs.get(fname.strip().split("/")[-1])
    if data is None:
        return f"文件不在本次输入里：{fname}", None
    try:
        return "", pick(data, path)
    except (KeyError, IndexError, TypeError) as exc:
        return f"键取不到：{path}（{exc}）", None


def factor_registered(pointer, rows):
    """派生因子是否已有登记行（或同一指标块已被登记），避免只登记乘积、不登记来源。"""
    target = str(pointer).strip()
    if any(target == str(r[3]).split("（")[0].strip() for r in rows):
        return True
    fname, _, path = target.partition("#")
    head = path.split(".")[0]
    return any(fname in str(r[3]) and head in str(r[3]) for r in rows)


def conclusion_table_unsourced(docx_path, rows):
    """结论稿表格里的数值是否都能在对照表找到一行（无源数值 = 断链）。"""
    if not docx_path or not Path(docx_path).is_file():
        return []
    from docx import Document
    known = []
    for row in rows:
        for part in str(row[2]).replace("%", "").split("/"):
            try:
                known.append(float(part.strip()))
            except ValueError:
                continue
    bad = []
    doc = Document(docx_path)
    for ti, table in enumerate(doc.tables):
        for line in table.rows:
            for cell in line.cells:
                text = (cell.text or "").strip().replace(",", "")
                if not re.fullmatch(r"-?\d+(?:\.\d+)?%?", text):
                    continue
                value = float(text.rstrip("%"))
                if "." not in text and abs(value) < 100:
                    continue            # 序号 / 矩阵阶数 / 专家数 这类计数列不参与
                if any(abs(value - n) <= max(0.0005, abs(n) * 0.001) for n in known):
                    continue
                bad.append((ti, text, " | ".join(c.text.strip() for c in line.cells)[:70]))
    return bad


def conclusion_source_lines(docx_path):
    """结论稿正文里残留的「来源：/（数据来源：…）」行数；文件不存在返回 None。"""
    if not docx_path or not Path(docx_path).is_file():
        return None, []
    from docx import Document
    doc = Document(docx_path)
    hits = [(p.text or "").strip()[:60] for p in doc.paragraphs
            if (p.text or "").strip().startswith(SOURCE_PREFIX)]
    return len(doc.paragraphs), hits


# ---- 正文段落数值 × 标签对账 -------------------------------------------------
# 表格里的数有来源只能证明「数存在」，不能证明「标签没被换」。正文里
# 「权重 0.744001」这种写法必须和对应字段对得上：0.744 是贡献值，不是权重。
LABEL_KINDS = {"综合得分": "score", "得分率": "score", "得分": "score",
               "贡献": "contrib", "权重": "weight", "占比": "weight",
               "λmax": "cons", "CI": "cons", "CR": "cons", "RI": "cons"}
LABEL_SEP = set(" 　:：=为约×|（）()【】[],，、-<>≤≥")
NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


def label_before(text, start):
    head = text[max(0, start - 16):start]
    for label in sorted(LABEL_KINDS, key=len, reverse=True):
        idx = head.rfind(label)
        if idx < 0:
            continue
        if not all(ch in LABEL_SEP for ch in head[idx + len(label):]):
            continue
        if label in ("CI", "CR", "RI") and idx > 0 and head[idx - 1].isalpha():
            continue
        return label
    return ""


def number_pools(weights, scores):
    """各标签该对到哪一池数：权重 / 得分 / 贡献 / 一致性。"""
    w, s = weights, scores
    weight, score, contrib, cons = [], [], [], []
    for c in w.get("criteria") or []:
        weight.append(float(c.get("weight") or 0))
    for it in w.get("indicators") or []:
        weight += [float(it.get("localWeight") or 0), float(it.get("globalWeight") or 0)]
    for sch in s.get("schemes") or []:
        if sch.get("composite") is not None:
            score.append(float(sch["composite"]))
        for it in sch.get("scores") or []:
            if it.get("globalWeight") is not None:
                weight.append(float(it["globalWeight"]))
            if it.get("score") is not None:
                score.append(float(it["score"]))
    for r in s.get("ranking") or []:
        if r.get("composite") is not None:
            score.append(float(r["composite"]))
    for d in ((s.get("conclusion") or {}).get("topDrivers") or []):
        if d.get("score") is not None:
            score.append(float(d["score"]))
        if d.get("contribution") is not None:
            contrib.append(float(d["contribution"]))
    c = w.get("consistency") or {}
    for m in c.get("matrices") or []:
        for f in ("lambdaMax", "ci", "ri", "cr"):
            if m.get(f) is not None:
                cons.append(float(m[f]))
    if c.get("threshold") is not None:
        cons.append(float(c["threshold"]))
    return {"weight": weight, "score": score, "contrib": contrib, "cons": cons}


def conclusion_paragraph_unsourced(docx_path, weights, scores):
    """正文段落里「标签 + 数值」是否对得上 JSON 字段；返回 [(标签, 数值, 段落)]。"""
    if not docx_path or not Path(docx_path).is_file():
        return []
    from docx import Document
    pools = number_pools(weights, scores)
    bad = []
    for p in Document(docx_path).paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue
        for m in NUM_RE.finditer(text):
            label = label_before(text, m.start())
            if not label:
                continue
            kind = LABEL_KINDS[label]
            try:
                value = float(m.group())
            except ValueError:
                continue
            if any(abs(value - n) <= max(0.0005, abs(n) * 0.002) for n in pools.get(kind, [])):
                continue
            bad.append((label, m.group(), text[:70]))
    return bad


def apply_print_layout(wb, demo: bool = False):
    """整本打印版式：横向、只按宽度收口、重复表头、页脚「页码 + 演示标注」。不改任何数值。"""
    from openpyxl.worksheet.properties import PageSetupProperties
    footer_text = ("演示样例（demo=true）　" if demo else "") + "第 &P 页 / 共 &N 页"
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        # 不按高度硬压：fitToHeight=2 会把 62 行表格压掉近一半（5 pt 不可读）。
        # 允许页数增加，优先保证字号 ≥10 pt。
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_title_rows = "1:1"
        ws.page_margins.left = ws.page_margins.right = 0.4
        # 上下边距收窄：把第 3/6 页只剩一两行表格行的尾巴收回前一页
        ws.page_margins.top = 0.4
        ws.page_margins.bottom = 0.45
        ws.page_margins.footer = 0.25
        # 页脚字号写死并跟正文一致（正文 11 pt），不随打印缩放变小
        ws.oddFooter.center.text = footer_text
        ws.oddFooter.center.size = 11


def _units(text):
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


def break_pointer(text, limit=52):
    """只在路径确实超过列宽时加断行点，避免 Excel 自动折行把「.」顶到行首。

    能一行放下的路径保持单行（否则每条都多占一行，白白增加页数）。
    """
    text = str(text)
    if "#" not in text or _units(text) <= limit:
        return text
    head, _, tail = text.partition("#")
    left, sep, right = tail.rpartition("/")
    if sep and _units(f"{head}#{left}{sep}") <= limit + 6:
        return f"{head}#{left}{sep}\n{right}"
    return f"{head}\n#{tail}"


def wsheet(ws, headers, rows, widths, bad=None, pointer_col=None):
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font = PatternFill("solid", fgColor="1F3864"), Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for idx, row in enumerate(rows):
        vals = list(row)
        if pointer_col:
            vals[pointer_col - 1] = break_pointer(vals[pointer_col - 1])
        ws.append(vals)
        if bad and bad[idx]:
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor="FCE4E4")
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"


def main(argv=None):
    ap = argparse.ArgumentParser(description="AHP 依据对照表")
    ap.add_argument("--weights", required=True, help="权重表.json（05/07 产出）")
    ap.add_argument("--scores", required=True, help="评分表.json（08 产出）")
    ap.add_argument("--out-xlsx", default="")
    ap.add_argument("--out-json", default="")
    ap.add_argument("--conclusion", default="", help="结论稿 docx（缺省按成果文件目录推）")
    ap.add_argument("--demo", action="store_true",
                    help="演示样例：页脚标注「演示样例」，并在对照表 JSON 写 demo=true")
    ap.add_argument("--verify-json", default="", help="只复查已有依据对照表 JSON 的行能否回指原文")
    args = ap.parse_args(argv)
    for p in (args.weights, args.scores):
        if not Path(p).is_file():
            print(f"[AHP-E-INPUT] 文件不存在：{p}", file=sys.stderr)
            return 2
    if args.verify_json:
        reg = read_json(args.verify_json)
        rows = reg.get("rows") or []
        docs = {Path(args.weights).name: read_json(args.weights),
                Path(args.scores).name: read_json(args.scores)}
        bad = []
        for row in rows:
            reason = verify_pointer(str(row[3] if len(row) > 3 else ""),
                                     row[2] if len(row) > 2 else "", docs)
            if reason:
                bad.append(f"{row[0]} {row[1]}：{reason}")
        print(f"[AHP 依据复查] {len(rows)} 行，回指失败 {len(bad)} 行")
        for line in bad[:10]:
            print("   -", line)
        return 0 if not bad else 1

    w = read_json(args.weights)
    s = read_json(args.scores)
    criteria = w.get("criteria") or []
    indicators = w.get("indicators") or []
    cons = w.get("consistency") or {}

    rows, detail, bad_rows, problems = [], [], [], []

    def add(section, name, value, pointer, ok=True, note=""):
        pointer = bundle_pointer(pointer)
        rows.append([section, name, value, pointer, "通过" if ok else "不通过", note])
        bad_rows.append(not ok)
        detail.append([section, name, value, pointer])
        if not ok:
            problems.append(f"{section} {name}：{note or '与源数据不一致'}")

    # 1) 准则层权重
    for i, c in enumerate(criteria):
        add("准则层权重", c.get("name", ""), f"{c.get('weight', 0):.4f}",
            f"权重表.json#criteria[{i}].weight")
    total = sum(float(c.get("weight") or 0) for c in criteria)
    add("准则层权重", "合计", f"{total:.4f}", "权重表.json#criteria[].weight（独立求和）",
        abs(total - 1) <= TOL_W, f"合计 {total:.4f}，应为 1（容差 {TOL_W}）")

    # 2) 指标层权重
    by_crit = {}
    for j, ind in enumerate(indicators):
        add("指标层权重", ind.get("name", ""),
            f"{float(ind.get('localWeight') or 0):.4f} / {float(ind.get('globalWeight') or 0):.4f}",
            f"权重表.json#indicators[{j}].localWeight/globalWeight")
        by_crit.setdefault(ind.get("criterionId"), []).append(float(ind.get("localWeight") or 0))
    for cid, locals_ in by_crit.items():
        add("指标层权重", f"{cid} 局部权重合计", f"{sum(locals_):.4f}",
            f"权重表.json#indicators[criterionId={cid}].localWeight（独立求和）",
            abs(sum(locals_) - 1) <= max(TOL_W, 0.002 * len(locals_)),
            f"合计 {sum(locals_):.4f}，应为 1")

    # 3) 一致性检验
    for k, m in enumerate(cons.get("matrices") or []):
        for field, label in (("lambdaMax", "λmax"), ("ci", "CI"), ("ri", "RI"), ("cr", "CR")):
            value = m.get(field)
            add("一致性检验", f"{m.get('id', k)} {label}",
                f"{float(value or 0):.6f}", f"权重表.json#consistency.matrices[{k}].{field}")
        add("一致性检验", f"{m.get('id', k)} 判定", "通过" if m.get("pass") else "不通过",
            f"权重表.json#consistency.matrices[{k}].pass", bool(m.get("pass")),
            "CR≥阈值" if not m.get("pass") else "")

    # 4) 方案综合得分（独立复算）
    weights_by_id = {ind.get("id"): float(ind.get("globalWeight") or 0) for ind in indicators}
    for idx, scheme in enumerate(s.get("schemes") or []):
        parts, mismatch = [], []
        for item in scheme.get("scores") or []:
            weight = float(item.get("globalWeight") or 0)
            expect = weights_by_id.get(item.get("id"))
            if expect is not None and abs(expect - weight) > 1e-9:
                mismatch.append(f"{item.get('name')} 权重 评分表 {weight:.9f} ≠ 权重表 {expect:.9f}")
            score = float(item.get("score") or 0)
            parts.append(score * weight)
            basis = str(item.get("basis") or "")
            if any(word in basis for word in BAD_WORDS):
                mismatch.append(f"{item.get('name')} 评分依据含占位措辞「{basis[:20]}」")
            elif not basis.strip():
                mismatch.append(f"{item.get('name')} 未写评分依据")
        recomputed = sum(parts)
        actual = float(scheme.get("composite") or 0)
        ok = abs(recomputed - actual) <= TOL_SCORE and not mismatch
        add("方案综合得分", scheme.get("name", ""), f"{actual:.2f}",
            f"评分表.json#schemes[{idx}].composite（复算 {recomputed:.4f}）", ok,
            "；".join(mismatch) or ("" if ok else f"复算 {recomputed:.4f} ≠ 登记 {actual:.4f}"))

    # 5) 排名与拉分项（得分 / 贡献各登记一行；贡献是派生值，必须写清 = 得分 × 权重）
    ind_index = {it.get("id"): i for i, it in enumerate(w.get("indicators") or [])}
    for idx, rank in enumerate(s.get("ranking") or []):
        add("排序", f"第 {rank.get('order')} 名 {rank.get('name', '')}",
            f"{float(rank.get('composite') or 0):.2f}",
            f"评分表.json#ranking[{idx}].composite")
    for idx, driver in enumerate(((s.get("conclusion") or {}).get("topDrivers") or [])):
        add("拉分项", driver.get("name", ""), str(driver.get("score")),
            f"评分表.json#conclusion.topDrivers[{idx}].score")
        j = ind_index.get(driver.get("id"))
        weight, contribution = driver.get("weight"), driver.get("contribution")
        if j is not None and weight is not None and contribution is not None:
            add("拉分项派生", f"{driver.get('name', '')} 贡献",
                f"{float(contribution):.6f}",
                f"派生：评分表.json#conclusion.topDrivers[{idx}].score"
                f" × 权重表.json#indicators[{j}].globalWeight",
                True,
                f"贡献 = 得分 {driver.get('score')} × 权重 {float(weight):.6f}")

    # 6) 追溯链：每行回指源文件的实际键；结论稿正文 0 处来源行
    docs = {Path(args.weights).name: w, Path(args.scores).name: s}
    for row in rows:
        reason = verify_pointer(str(row[3]), row[2], docs)
        if reason:
            row[4] = "不通过"
            problems.append(f"{row[0]} {row[1]}：回指不到原文（{reason}）")
    # 派生值必须有登记行：拉分项贡献不能只出现在正文里
    top = ((s.get("conclusion") or {}).get("topDrivers") or [])
    derived = [r for r in rows if str(r[3]).startswith("派生：")]
    if top and len(derived) < len(top):
        problems.append(f"拉分项派生值未登记：{len(derived)}/{len(top)} 条贡献有派生行"
                        "（应写清 贡献 = 得分 × 权重）")
    for row in derived:
        for factor in str(row[3])[len("派生："):].split(" × "):
            if not factor_registered(factor.strip(), rows):
                problems.append(f"{row[0]} {row[1]}：派生因子 {factor.strip()} 没有对应登记行")
    docx_path = args.conclusion
    if not docx_path and args.out_xlsx:
        base = Path(args.out_xlsx)
        candidates = [base.parent / "AHP权重与评分结论.docx",
                      base.parent.parent / "成品版/AHP权重与评分结论"
                                           "/AHP权重与评分结论_成品版.docx"]
        docx_path = next((str(p) for p in candidates if p.is_file()), "")
    n_para, src_hits = conclusion_source_lines(docx_path)
    if n_para is None:
        problems.append("结论稿未找到，无法核对「正文 0 处来源行」"
                        f"（{docx_path or '未给 --conclusion'}）")
    elif src_hits:
        problems.append(f"结论稿正文仍有 {len(src_hits)} 处来源行：{'；'.join(src_hits[:3])}")
    for ti, text, context in conclusion_table_unsourced(docx_path, rows):
        problems.append(f"结论稿表{ti + 1} 的 {text} 不在对照表任何一行（无源数值）：{context}")
    for label, value, context in conclusion_paragraph_unsourced(docx_path, w, s):
        problems.append(f"结论稿正文「{label} {value}」与源字段对不上（疑似标签错配）：{context}")

    ok_all = all(not row[4] == "不通过" for row in rows) and not problems
    demo = bool(args.demo or w.get("demo") or s.get("demo"))

    out_xlsx = Path(args.out_xlsx) if args.out_xlsx else None
    if out_xlsx:
        out_xlsx.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "按段对照"
        wsheet(ws, ["成稿章节", "对象", "数值", "依据（源文件#键）", "状态", "说明"],
               rows, [10, 18, 12, 60, 6, 26], bad_rows, pointer_col=4)
        ws2 = wb.create_sheet("逐数明细")
        wsheet(ws2, ["成稿章节", "对象", "数值", "依据"], detail, [12, 22, 16, 54],
               pointer_col=4)
        # 校验结论并入同一页签尾部：避免单独开一页只印 1 行，形成碎片页
        # 校验结论紧跟表尾（不留空行）：否则它会单独溢出成只剩一行的碎片页
        base = len(rows) + 2
        ws.cell(row=base, column=1, value="校验结论").font = Font(bold=True)
        for i, text in enumerate(problems or ["无：所有数值均可追回源文件"]):
            r = base + 1 + i
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            c = ws.cell(row=r, column=1, value=text)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 24
        apply_print_layout(wb, demo=demo)
        wb.save(out_xlsx)

    out_json = Path(args.out_json) if args.out_json else None
    if out_json:
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(json.dumps({
            "schema": "cardo/evidence-register/v1", "domain": "ahp",
            "demo": demo, "ok": ok_all, "rows": rows, "problems": problems}, ensure_ascii=False, indent=1),
            encoding="utf-8")

    print("[AHP 依据对照表] {} 行，不通过 {} 行{}".format(
        len(rows), sum(1 for r in rows if r[4] == "不通过"),
        "" if ok_all else "：详见按段对照页签尾部「校验结论」"))
    for problem in problems[:8]:
        print("   -", problem)
    return 0 if ok_all else 1


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
