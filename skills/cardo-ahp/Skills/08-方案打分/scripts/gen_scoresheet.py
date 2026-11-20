#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_scoresheet.py -- 生成人读《AHP 方案评分表》xlsx（输入文件侧的唯一评分输入）。

版式（xlsx 即输入事实源，JSON 只是过程产物）：
  · Sheet「填写说明」：口径、评级、缺失处理说明；
  · Sheet「方案信息」：序号 | 方案编号 | 方案名称 | 方案说明（编号 S1/S2…，与引擎对应）；
  · Sheet「方案评分表」：第 3 行表头，数据从第 4 行起，按准则分组：
      A序号 B准则层 C指标 D指标ID E指标含义 F综合权重 G参考分档（1/3/5/7/9 各代表什么）
      H.. 方案一/二/三得分（0-10，黄色格为待填）
      其后 依据-方案一/二/三（逐项依据文本）
    末尾「综合得分（加权合计）」行给 SUMPRODUCT 公式（仅展示，最终以引擎 JSON 为准）。

用法：
  python gen_scoresheet.py --weights <权重表.json> --out <评分表.xlsx> [--scores <打分.json>]
    不带 --scores：出空白表（新项目发给项目组/专家填）。
    带 --scores：把已有评分与依据回填（迁移或换版式时用）。
退出码：0 成功；2 输入/参数错。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from _yaml_lite import load_file

RULES_DIR = HERE.parent / "规则"

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

HEADER_ROW = 3
FIRST_DATA_ROW = 4
SCORE_COL = 8          # H 列


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def load_catalog():
    data = load_file(RULES_DIR / "指标体系.yaml")
    criteria = {c["id"]: c for c in data["criteria"]}
    anchors = {}
    for ind in data["indicators"]:
        anchors[ind["id"]] = "；".join(f"{a['score']}分={a['desc']}" for a in ind.get("anchors", []))
    return data, criteria, anchors


def load_anchor_scale():
    """取各指标的评分档（用于「参考分档」列）。

    返回 {指标id: [(分值, 含义), ...]}。这些是行业标准锚点（规则里定义），
    用来告诉打分人「这一格该打几分、各分值代表什么」——
    预填的是**标准**，不是某个方案的得分。
    """
    data = load_file(RULES_DIR / "指标体系.yaml")
    out = {}
    for ind in data.get("indicators", []):
        out[ind["id"]] = [(a["score"], a["desc"]) for a in ind.get("anchors", [])]
    return out


def score_of(manual: dict, iid: str, basis_map: dict | None = None):
    v = manual.get(iid)
    if v is None:
        return None, (basis_map or {}).get(iid, "")
    if isinstance(v, dict):
        return v.get("score"), v.get("basis") or (basis_map or {}).get(iid, "")
    return v, (basis_map or {}).get(iid, "")


def read_scheme_list(path: Path) -> list[dict]:
    """读《方案清单.xlsx》（人填的方案名以此为准）。实现见 scheme_lib。"""
    from scheme_lib import read_scheme_list as _read
    return _read(Path(path))


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 AHP 方案评分表 xlsx（空白模板或回填）")
    ap.add_argument("--weights", default="",
                    help="cardo/ahp-weights/v1 权重表 JSON（配合 --no-weights 可省略）")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    ap.add_argument("--scores", default=None, help="可选：打分 JSON，提供则回填分数与依据")
    ap.add_argument("--scheme-list", default=None,
                    help="可选：《方案清单.xlsx》路径，提供则按其中方案名预填评分表")
    ap.add_argument("--project-name", default="某某片区改造项目")
    ap.add_argument("--no-weights", action="store_true",
                    help="权重尚未算出时出结构表：综合权重列留空待回填，其余版式照常")
    args = ap.parse_args()

    weights = None
    if not args.no_weights:
        weights_path = Path(args.weights)
        if not weights_path.is_file():
            print(f"[AHP-E-INPUT] 权重文件不存在：{args.weights}", file=sys.stderr)
            return 2
        weights = json.loads(weights_path.read_text(encoding="utf-8"))
        if not isinstance(weights.get("indicators"), list):
            print("[AHP-E-INPUT] 权重 JSON 缺 indicators[]", file=sys.stderr)
            return 2

    payload = None
    if args.scores:
        sp = Path(args.scores)
        if not sp.is_file():
            print(f"[AHP-E-INPUT] 打分文件不存在：{args.scores}", file=sys.stderr)
            return 2
        payload = json.loads(sp.read_text(encoding="utf-8-sig"))

    schemes = []
    if payload and isinstance(payload.get("schemes"), list):
        for idx, s in enumerate(payload["schemes"], start=1):
            schemes.append({"id": s.get("id") or f"S{idx}",
                            "name": s.get("name") or f"方案{['一','二','三','四','五'][idx - 1] if idx <= 5 else idx}",
                            "note": s.get("schemeNote") or "",
                            "manual": s.get("manualScores") or {},
                            "basisMap": s.get("basis") or {}})
    if not schemes:
        # 优先读用户填好的《方案清单.xlsx》（人填的方案名以此为准）；
        # 没有清单时才回落到三个默认方案。
        listed = read_scheme_list(Path(args.scheme_list)) if args.scheme_list else []
        if listed:
            schemes = [{"id": s["id"], "name": s["name"], "note": s["note"],
                        "manual": {}, "basisMap": {}} for s in listed]
            print(f"已按方案清单预填 {len(schemes)} 个方案："
                  + "、".join(s["name"] for s in schemes))
        else:
            # 方案清单还没填 → 不静默塞「方案一/二/三」，那不是用户的方案。
            # 给出 3 个空名占位并在填写说明里写清要改名；用户填了清单后重跑即预填真名。
            print("[提示] 方案清单为空或未填，评分表先出空结构；"
                  "请在《方案清单.xlsx》填好方案名后重跑，评分表会自动预填。")
            schemes = [{"id": f"S{i + 1}", "name": "", "note": "", "manual": {}, "basisMap": {}}
                       for i in range(3)]

    data, criteria, anchors = load_catalog()
    anchor_scale = load_anchor_scale()
    # 权重还没算出时（首跑即出表），综合权重列留空待第 07 步后回填。
    gw = ({i["id"]: float(i.get("globalWeight") or 0.0) for i in weights["indicators"]}
          if weights else {})

    wb = Workbook()

    # ---------------- 填写说明 ----------------
    ws0 = wb.active
    ws0.title = "填写说明"
    notes = [
        "AHP 候选改造方案评分表（10 分制）",
        "",
        f"项目名称：{args.project_name}（匿名占位，真实名称仅出现在项目内部材料中）",
        "评价人：______________    评价日期：______________",
        "",
        "一、本表是评分的唯一输入",
        "1. 只改本 xlsx：「方案评分表」黄色单元格填 0-10 分（允许 1 位小数，建议按 1/3/5/7/9 打），",
        "   「依据」列写每项打分出处（现状分析/财务测算/专家研判）；「方案信息」表登记方案编号与名称。",
        "2. 不要手改「指标ID」「综合权重」「参考分档」三列：它们由系统按指标体系与权重表生成，改动会导致对不上。",
        "   「参考分档」列写的是这个指标打 1/3/5/7/9 分分别代表什么样子 —— 照着它给分，不用自己判断标准。",
        "3. 「综合权重」列现在是空的：要等专家问卷回收、权重算出来之后才有值（届时系统自动回填），",
        "   不影响你先填得分。",
        "4. 引擎实际读取的是由本表自动转换出的 过程文件/08-方案打分/打分.json，无需人工维护 JSON。",
        "",
        "二、综合得分与评级（由引擎计算并写入 成果文件/AHP方案评分结论.xlsx）",
        "综合得分 = Σ（各指标得分 × 综合权重）；本表 16 项综合权重之和 = 1。",
        "评级（左闭右开）：[9,10] 优秀＝优先投资、全力推进；[7,9) 良好＝可行投资、分期推进；",
        "[6,7) 一般＝审慎投资、优化方案；[0,6) 不可行＝暂缓/放弃投资。",
        "缺项不得打 0 分：必选指标缺依据则该方案不出分（引擎会拦下并点名），补依据后再算。",
        "",
        "三、溯源：权重见 成果文件/AHP权重表.xlsx；过程留痕（含机读 JSON）见 过程文件/。",
    ]
    for r, text in enumerate(notes, start=1):
        ws0.cell(row=r, column=1, value=text)
    ws0.cell(row=1, column=1).font = Font(size=14, bold=True)
    ws0.column_dimensions["A"].width = 110

    # ---------------- 方案信息 ----------------
    wi = wb.create_sheet("方案信息")
    info_head = ["序号", "方案编号", "方案名称", "方案说明"]
    for col, h in enumerate(info_head, start=1):
        c = wi.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = WRAP
        c.border = BORDER
    for idx, s in enumerate(schemes, start=1):
        # 方案名留空时给出占位提示，避免表头/依据列前缀全是空白格
        name = s["name"] or f"（待填方案名 {s['id']}）"
        vals = [idx, s["id"], s["name"], s["note"]]
        for col, v in enumerate(vals, start=1):
            c = wi.cell(row=idx + 1, column=col, value=v)
            c.alignment = LEFT if col >= 3 else WRAP
            c.border = BORDER
            if col == 3 and not s["name"]:
                c.fill = INPUT_FILL
        # 名称为空时，序号列右侧提示一次（在说明列上方加提示行）
        if not s["name"]:
            c = wi.cell(row=idx + 1, column=4, value="← 方案名请到《方案清单.xlsx》填后重跑")
            c.alignment = LEFT
            c.font = Font(size=9, color="808080")
    for col, w in enumerate([6, 10, 16, 46], start=1):
        wi.column_dimensions[get_column_letter(col)].width = w

    # ---------------- 方案评分表 ----------------
    ws = wb.create_sheet("方案评分表")
    n = len(schemes)
    score_start = SCORE_COL
    basis_start = score_start + n
    last_col = basis_start + n - 1

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(row=1, column=1, value=f"{args.project_name} · AHP 候选改造方案评分表")
    c.font = Font(size=14, bold=True, color="FFFFFF")
    c.fill = HEAD_FILL
    c.alignment = WRAP
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.cell(row=2, column=1,
            value="评价人：____________    评价日期：____________    "
                  "（只填方案得分的黄色格与依据列；「参考分档」列照着它给分；指标ID/权重由系统生成，请勿改动）"
            ).alignment = LEFT

    headers = ["序号", "准则层", "指标", "指标ID", "指标含义", "综合权重",
               "参考分档（照此给分：1/3/5/7/9 各代表什么）"]
    # 方案名为空时用占位名做表头（否则表头是空格，人看不出该填哪列）
    display = [s["name"] or f"方案{s['id']}（待填名）" for s in schemes]
    headers += display + [f"依据-{d}" for d in display]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = WRAP
        cell.border = BORDER

    criterion_order = [c["id"] for c in data["criteria"]]
    ind_by_criterion: dict[str, list] = {}
    for ind in data["indicators"]:
        ind_by_criterion.setdefault(ind["criterionId"], []).append(ind)

    row = FIRST_DATA_ROW
    seq = 0
    for cid in criterion_order:
        inds = ind_by_criterion.get(cid, [])
        start = row
        for ind in inds:
            iid = ind["id"]
            seq += 1
            ws.cell(row=row, column=1, value=seq).alignment = WRAP
            ws.cell(row=row, column=3, value=ind["name"]).alignment = LEFT
            idc = ws.cell(row=row, column=4, value=iid)
            idc.alignment = WRAP
            idc.font = Font(size=9, color="808080")
            ws.cell(row=row, column=5, value=ind.get("definition", "")).alignment = LEFT
            if gw:
                wcell = ws.cell(row=row, column=6, value=gw.get(iid))
                wcell.number_format = "0.00%"
            else:
                # 权重未算出：留空待第 07 步后回填，不写假值
                wcell = ws.cell(row=row, column=6, value=None)
                wcell.number_format = "0.00%"
                wcell.fill = GREY_FILL
            wcell.alignment = WRAP
            # 参考分档：行业标准锚点（规则里定义），照此给分。
            # 每档一行，读的时候一眼能对上「打几分 → 什么样子」。
            scale = anchor_scale.get(iid, [])
            scale_text = "\n".join(f"{sc}分 = {desc}" for sc, desc in scale)
            gcell = ws.cell(row=row, column=7, value=scale_text or anchors.get(iid, ""))
            gcell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            gcell.font = Font(size=9)
            for k, s in enumerate(schemes):
                val, _ = score_of(s["manual"], iid, s.get("basisMap"))
                sc = ws.cell(row=row, column=score_start + k, value=val)
                sc.alignment = WRAP
                sc.fill = INPUT_FILL if val is None else GREY_FILL
                sc.number_format = "0.0"
            for k, s in enumerate(schemes):
                _, b = score_of(s["manual"], iid, s.get("basisMap"))
                bc = ws.cell(row=row, column=basis_start + k, value=b or None)
                bc.alignment = LEFT
                bc.font = Font(size=9)
            for col in range(1, last_col + 1):
                ws.cell(row=row, column=col).border = BORDER
            row += 1
        ws.merge_cells(start_row=start, start_column=2, end_row=row - 1, end_column=2)
        gc = ws.cell(row=start, column=2,
                     value=f"{criteria[cid]['name']}\n（{criteria[cid].get('displayName', {}).get('taskbook', '')}）")
        gc.alignment = WRAP
        gc.fill = GROUP_FILL
        gc.font = Font(bold=True)
        for rr in range(start, row):
            ws.cell(row=rr, column=2).border = BORDER
            ws.cell(row=rr, column=2).fill = GROUP_FILL

    # 综合得分行（公式仅展示；以引擎 评分表.json 为准）
    ws.cell(row=row, column=3, value="综合得分（加权合计）").font = Font(bold=True)
    last_data = row - 1
    for k in range(n):
        col = get_column_letter(score_start + k)
        if gw:
            cell = ws.cell(row=row, column=score_start + k,
                           value=f"=SUMPRODUCT($F${FIRST_DATA_ROW}:$F${last_data},{col}{FIRST_DATA_ROW}:{col}{last_data})")
            cell.number_format = "0.00"
            cell.font = Font(bold=True)
        else:
            # 权重未算出：公式算出来是 0，等于假结果 —— 留空并说明
            cell = ws.cell(row=row, column=score_start + k, value="待算权重")
            cell.font = Font(size=9, color="808080")
        cell.fill = GROUP_FILL
        cell.border = BORDER
    ws.cell(row=row + 1, column=3,
            value="评级：[9,10]优秀 / [7,9)良好 / [6,7)一般 / [0,6)不可行（左闭右开）").font = Font(italic=True)
    if not gw:
        ws.cell(row=row + 2, column=3,
                value="提示：综合权重还没算出（要等专家问卷回收），本行暂时算不出；"
                      "算完权重后重跑即可回填。").font = Font(size=9, color="808080")

    widths = [6, 14, 18, 13, 32, 10, 46] + [10] * n + [34] * n
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # 每行 5 档分档要显示得下（5 行 × 12pt ≈ 60），给足行高
    for r in range(FIRST_DATA_ROW, row):
        ws.row_dimensions[r].height = 62
    ws.freeze_panes = f"{get_column_letter(score_start)}{FIRST_DATA_ROW}"

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"已生成评分表：{out}（方案 {n} 个，指标 {len(data['indicators'])} 项，"
          f"回填：{'是' if payload else '否（空白表）'}）")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
