#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付装配.py -- 09 收尾：把散件装配成「成品版 / 内部件」两套，并出成品 PDF。

做五件事（一件事一个闸门，全部落在 交付装配.json 里）：
  1. 分存：07/08 的 xlsx 从 成果文件/ 根下移进 成品版/<成果名>/，一成果一子目录；
     同时生成同源「溯源版」xlsx（多一页「内部核对」：随包源 JSON 与 sha256）。
  2. 依据随包：把权重表.json / 评分表.json / 群体矩阵.json / 矩阵集.json
     复制到 内部件/依据数据/，交付包离开本机仍能逐数回指。
  3. 成品 PDF：复用 pdf_render（docx → HTML → 无头浏览器打印），按实际分页把
     目录页码回填进 docx，再渲染一次并复核分页稳定，最后把页脚烘进版面。
     溯源版同样回填目录页码（其 PDF 只用于算页码，不留件）。
  4. 词表闸门：成品版 docx 与 PDF 正文不得出现「演示数据/合成演示/示例数据/
     非真实依据」等禁词，也不得出现 过程文件/ 内部件/ 依据数据/ 等内部件引用；
     **唯一豁免**是方法说明里那一句合成问卷口径（按段落位置白名单，逐字比对），
     不按"出现即放过"放宽。
  5. 0 分母：通过率、覆盖率一类比值分母为 0 时写「不适用」，不写 0/0。

退出码：0 全通；1 有闸门未过；2 输入错。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import Alignment

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pdf_render                      # noqa: E402
import 依据对照表 as EVID              # noqa: E402
import 结论成稿 as DOC                 # noqa: E402

# ---- 交付结构（一成果一子目录；成品版对外，内部件供复核）----
PUB, INT = "成品版", "内部件"
BOOK_W, BOOK_S, BOOK_C, BOOK_E = "AHP权重表", "AHP方案评分结论", "AHP权重与评分结论", "AHP依据对照表"
REL_XLS_W_CLEAN = f"成果文件/{PUB}/{BOOK_W}/{BOOK_W}_成品版.xlsx"
REL_XLS_W_ANN = f"成果文件/{INT}/{BOOK_W}/{BOOK_W}_溯源版.xlsx"
REL_XLS_S_CLEAN = f"成果文件/{PUB}/{BOOK_S}/{BOOK_S}_成品版.xlsx"
REL_XLS_S_ANN = f"成果文件/{INT}/{BOOK_S}/{BOOK_S}_溯源版.xlsx"
REL_DOCX_CLEAN = f"成果文件/{PUB}/{BOOK_C}/{BOOK_C}_成品版.docx"
REL_PDF_CLEAN = f"成果文件/{PUB}/{BOOK_C}/{BOOK_C}_成品版.pdf"
REL_DOCX_ANN = f"成果文件/{INT}/{BOOK_C}/{BOOK_C}_溯源版.docx"
REL_EVID = f"成果文件/{INT}/{BOOK_E}/{BOOK_E}.xlsx"
REL_DATA = f"成果文件/{INT}/依据数据"
# 07/08 的中间落位（装配时移走，成果根下不留散件）
REL_LEGACY = {"成果文件/AHP权重表.xlsx": REL_XLS_W_CLEAN,
              "成果文件/AHP方案评分结论.xlsx": REL_XLS_S_CLEAN}

# 正文禁词分七组：任一组命中即不通过。唯一豁免是「段落位置白名单」里的两段
# （合成问卷口径 → 一、方法与口径；结论适用范围 → 结论与边界），逐字比对、各限 1 次。
BANNED_GROUPS = {
    "演示身份": ("演示数据", "合成演示", "演示版", "演示样例"),
    "模拟口径": ("模拟专家判断", "模拟数据", "模拟问卷", "非真实依据"),
    "免责话术": ("仅供参考", "不构成投资建议", "不构成任何投资建议", "示例数据", "示例"),
    "情景假设": ("假设情景数据", "假设数据", "情景数据", "假设条件"),
    "手册引用": ("《手册》", "手册", "内部口径", "内部资料", "台账"),
    "缺项占位": ("未提资", "缺口", "待补", "待补录", "占位", "无缺项"),
    "他件引用": ("详见", "附表", "附件", "附录", "另册", "另附", "随包", "底稿", "导出值"),
}
BANNED = tuple(dict.fromkeys(t for group in BANNED_GROUPS.values() for t in group))

# 阳性对照：每条样本含一个应被拦下的表述。扩词表后必须全部命中，
# 否则闸门就是睁眼瞎（R1 P1-④：原词表只有 4 个，拦不住「模拟专家判断」一类）。
POSITIVE_SAMPLES = (
    "本矩阵由模拟专家判断给出，未做真实访谈",
    "以上结果仅供参考",
    "本报告不构成投资建议",
    "假设情景数据下的排序仅用于跑通链路",
    "详见《手册》第三章",
    "未提资指标按缺省处理，存在缺口",
    "另附 附表 3 与附件 2",
)
# 否定式免责话术：按句式模式认，不穷举整句。R2 P1-1 漏的就是「不构成对片区的投资决策」——
# 旧表只收了「不构成投资建议」这一条整句。白名单两段由调用方先剔除，此处不再放宽。
DISCLAIM_PATTERNS = (
    re.compile(r"不构成[^。；\n]{0,24}"),
    re.compile(r"不代表[^。；\n]{0,24}"),
    re.compile(r"不意味[^。；\n]{0,24}"),
    re.compile(r"仅供[^。；\n]{0,16}"),
    re.compile(r"不作为[^。；\n]{0,16}"),
    re.compile(r"for\s+demonstration\s+only", re.I),
    re.compile(r"not\s+investment\s+advice", re.I),
)


def disclaimer_hits(text: str) -> list:
    """返回该段里命中的免责片段（同一条只报一次）。"""
    out = []
    for pat in DISCLAIM_PATTERNS:
        match = pat.search(text)
        if match and match.group(0) not in out:
            out.append(match.group(0))
    return out


# 成品 xlsx 不得残留模板示例（方案一/二/三 + 示例得分），也不得出现字段名 dump「导出值」。
SCHEME_RESIDUE = ("方案一", "方案二", "方案三", "示例", "导出值")
# 成品 xlsx 封面五项：项目、编制单位、版本、生成时间、页码（页脚域 &P）
COVER_FIELDS = ("项目名称", "编制单位", "版本", "生成时间")
# 成品版不得指向内部件/运行目录（离开本次运行就断链）
INTERNAL_REF = ("过程文件/", "过程文件\\", "内部件", "依据数据", "依据对照表",
                "溯源", "内部复核", "记忆与规则")


def read_json(p):
    return json.loads(Path(p).read_text(encoding="utf-8-sig"))


def sha256_file(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest().upper()


def count(text: str, term: str) -> int:
    return text.count(term)


def docx_sections(docx_path: str) -> list:
    """按一级标题切段：[(章节标题, [该节段落文本])]。封面/目录各自成节。"""
    doc = Document(docx_path)
    out = [("（封面）", [])]
    in_toc = False
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        try:
            level = int(re.match(r"^(?:Heading|标题)\s*(\d+)$", p.style.name or "").group(1))
        except (AttributeError, TypeError, ValueError):
            level = 0
        if level == 1:
            out.append((text, []))
            in_toc = text == "目录"
            continue
        if not text or in_toc:
            continue
        out[-1][1].append(text)
    return out


def docx_body_texts(docx_path: str, skip: tuple = ()) -> list:
    """正文段落 + 表格单元格文本；skip 为豁免段落（逐字相等才跳过）。"""
    doc = Document(docx_path)
    out = []
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if text and text not in skip:
            out.append(text)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                text = (cell.text or "").strip()
                if text and text not in skip:
                    out.append(text)
    return out


def check_whitelist(sections: list, problems: list) -> dict:
    """两段豁免只许各出现 1 次，且必须落在指定章节内（按段落位置白名单）。"""
    spec = ((DOC.SYNTHETIC_NOTE, "一、", "合成问卷口径段"),
            (DOC.SCOPE_NOTE, "结论与边界", "结论适用范围段"))
    result = {}
    for note, prefix, label in spec:
        where = [title for title, texts in sections if note in texts]
        result[label] = {"count": len(where), "sections": where}
        if len(where) != 1:
            problems.append(f"{label}应出现且仅出现 1 次，实际 {len(where)} 次 {where}")
        elif prefix not in where[0]:
            problems.append(f"{label}位置不对：现在「{where[0]}」，应在「{prefix}…」节内")
    return result


def positive_control() -> dict:
    """阳性对照实测：7 组样本必须条条被词表拦下。"""
    missed = [s for s in POSITIVE_SAMPLES if not any(t in s for t in BANNED)]
    return {"total": len(POSITIVE_SAMPLES), "caught": len(POSITIVE_SAMPLES) - len(missed),
            "missed": missed, "ok": not missed}


def scan_clean_workbook(path: Path, where: str, problems: list) -> dict:
    """成品 xlsx 自检：零公式 / 无 fullCalcOnLoad / 封面五项 / 无模板示例残留 / 无禁词。"""
    before = len(problems)
    res = {"file": where, "formulaCells": 0, "fullCalcOnLoad": "", "sheets": [],
           "cover": {}, "residue": [], "banned": [], "ok": False}
    with zipfile.ZipFile(str(path)) as zf:
        wbxml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    match = re.search(r'fullCalcOnLoad="([^"]*)"', wbxml)
    res["fullCalcOnLoad"] = match.group(1) if match else "未设置"
    if res["fullCalcOnLoad"] not in ("0", "false", "False", "未设置"):
        problems.append(f"{where} 带 fullCalcOnLoad={res['fullCalcOnLoad']}，Excel 打开会二次计算")
    wb = load_workbook(str(path), data_only=False)
    res["sheets"] = list(wb.sheetnames)
    texts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if value is None:
                    continue
                if cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                    res["formulaCells"] += 1
                    continue
                texts.append(str(value))
    if res["formulaCells"]:
        problems.append(f"{where} 仍有 {res['formulaCells']} 个公式格（成品应为静态值）")
    if not wb.sheetnames or wb.sheetnames[0] != "封面":
        problems.append(f"{where} 首屏不是「封面」（实际 {wb.sheetnames[:1]}）")
    else:
        cover = wb["封面"]
        found = {str(c.value).strip() for row in cover.iter_rows() for c in row
                 if isinstance(c.value, str) and str(c.value).strip() in COVER_FIELDS}
        footer = ((cover.oddFooter.center.text if cover.oddFooter else "") or "")
        res["cover"] = {"fields": sorted(found), "pageFooter": footer}
        missing = [k for k in COVER_FIELDS if k not in found]
        if missing:
            problems.append(f"{where} 封面缺项：{missing}")
        if "&P" not in footer:
            problems.append(f"{where} 封面页脚没有页码域（第 &P 页 / 共 &N 页）")
    res["residue"] = [t for t in SCHEME_RESIDUE if any(t in x for x in texts)]
    if res["residue"]:
        problems.append(f"{where} 仍有模板示例残留：{res['residue']}")
    res["banned"] = [t for t in list(BANNED) + list(INTERNAL_REF) if any(t in x for x in texts)]
    if res["banned"]:
        problems.append(f"{where} 表内出现禁词/内部引用：{res['banned'][:6]}")
    res["disclaimers"] = sorted({snip for x in texts for snip in disclaimer_hits(x)})
    if res["disclaimers"]:
        problems.append(f"{where} 表内出现否定式免责话术：{res['disclaimers'][:4]}")
    res["ok"] = len(problems) == before
    return res


def scan_clean(texts: list, where: str, problems: list) -> None:
    """成品版正文扫描：禁词 + 内部件引用。豁免段已由调用方剔除，不在这里放宽。"""
    for term in BANNED:
        hits = [t for t in texts if term in t]
        if hits:
            problems.append(f"{where} 正文出现禁词「{term}」：{hits[0][:40]}")
    for term in INTERNAL_REF:
        hits = [t for t in texts if term in t]
        if hits:
            problems.append(f"{where} 引用了内部件/运行目录「{term}」：{hits[0][:40]}")
    seen = set()
    for text in texts:
        for snip in disclaimer_hits(text):
            if snip in seen:
                continue
            seen.add(snip)
            problems.append(f"{where} 出现否定式免责话术「{snip}」：{text[:36]}")


def annotate_xlsx(src: str, dst: str, sources: list, demo: bool) -> None:
    """生成溯源版 xlsx：复制源表，再补一页「内部核对」（源 JSON + sha256 + 用途）。"""
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)
    wb = load_workbook(dst)
    # 溯源版是从成品复制来的，封面「成果名称」里的（成品版）要改成本件真实身份
    if "封面" in wb.sheetnames:
        for row in wb["封面"].iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "成品版" in cell.value:
                    cell.value = cell.value.replace("（成品版）", "（溯源版）")
    ws = wb.create_sheet("内部核对")
    ws.append(["随包源数据", "sha256（前 16 位）", "用途"])
    for rel, path, use in sources:
        ws.append([rel, sha256_file(path)[:16], use])
    ws.append([])
    ws.append(["核对口径",
               "本表每个数值与源 JSON 键的逐项对应关系见 内部件/AHP依据对照表/AHP依据对照表.xlsx"])
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 62
    for row in ws.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    EVID.apply_print_layout(wb, demo=demo)
    wb.save(dst)


def toc_state(docx_path: str) -> tuple[list, int]:
    """返回 (目录条目, 已填页码条数)。"""
    doc = Document(docx_path)
    in_toc, entries, filled = False, [], 0
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if DOC._is_heading_par(p):
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
        entries.append(text)
        if "\t" in text:
            filled += 1
    return entries, filled


def render_and_fill(docx_path: str, pdf_path: str) -> dict:
    """渲染 → 回填目录页码 → 再渲染复核；分页不稳就再回填一次。"""
    info = pdf_render.render_docx(docx_path, pdf_path)
    pages = dict(info.get("headingPages") or {})
    if not pages:
        raise RuntimeError("PDF 里没定位到任何章标题，无法回填目录页码")
    filled = DOC.fill_toc_pages(docx_path, pages)
    info2 = pdf_render.render_docx(docx_path, pdf_path)
    if (info2.get("headingPages") or {}) != pages:
        pages = dict(info2.get("headingPages") or {})
        DOC.reset_toc_pages(docx_path)
        filled = DOC.fill_toc_pages(docx_path, pages)
        info3 = pdf_render.render_docx(docx_path, pdf_path)
        if (info3.get("headingPages") or {}) != pages:
            raise RuntimeError("目录页码在两次渲染间不稳定，无法静态回填")
        info2 = info3
    return {"filled": filled, "pages": info2.get("pages"),
            "headingPages": pages, "tocPages": info2.get("tocPages") or [],
            "emptyPages": info2.get("emptyPages") or []}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AHP 交付装配（分存 / 依据随包 / 成品 PDF / 词表闸门）")
    ap.add_argument("--weights", required=True)
    ap.add_argument("--scores", required=True)
    ap.add_argument("--aggregation", required=True)
    ap.add_argument("--matrix-set", required=True)
    ap.add_argument("--weights-ann", default="过程文件/07-权重表导出/权重表_溯源.xlsx",
                    help="07 的模板底稿（full 矩阵逐格写入），作为溯源版权重表的源")
    ap.add_argument("--out-json", default="过程文件/09-成果验证与成稿/交付装配.json")
    ap.add_argument("--demo", action="store_true")
    args = ap.parse_args(argv)

    ws = Path.cwd()
    for p in (args.weights, args.scores, args.aggregation, args.matrix_set, args.weights_ann):
        if not Path(p).is_file():
            print(f"[AHP-E-INPUT] 文件不存在：{p}", file=sys.stderr)
            return 2
    problems: list[str] = []
    info: dict = {"ok": False, "problems": problems, "files": []}

    # ---- 1. 分存：成品版 xlsx 归位 + 同源溯源版 ----
    for legacy, final in REL_LEGACY.items():
        src, dst = ws / legacy, ws / final
        dst.parent.mkdir(parents=True, exist_ok=True)
        if src.is_file():
            if dst.exists():
                dst.unlink()
            shutil.move(str(src), str(dst))
        elif not dst.is_file():
            problems.append(f"找不到待装配的 xlsx：{legacy}（也不在 {final}）")
    for rel in LEGACY_ROOT_SCAN:
        if (ws / rel).is_file():
            problems.append(f"成果根下仍留散件：{rel}（应只在 成品版/<成果名>/ 下）")

    # ---- 2. 依据随包 ----
    data_dir = ws / REL_DATA
    data_dir.mkdir(parents=True, exist_ok=True)
    sources = [("权重表.json", Path(args.weights), "准则/指标权重与一致性（λmax、CI、CR）"),
               ("矩阵集.json", Path(args.matrix_set), "各专家判断矩阵（复算用）"),
               ("评分表.json", Path(args.scores), "16 项指标得分、综合得分与排序"),
               ("群体矩阵.json", Path(args.aggregation), "专家人数与几何平均后的群体矩阵")]
    for name, path, _use in sources:
        target = data_dir / name
        shutil.copyfile(path, target)
        info["files"].append(f"{REL_DATA}/{name}")
    src_sha = {name: sha256_file(path) for name, path, _ in sources}
    for name, _path, _use in sources:
        if sha256_file(data_dir / name) != src_sha[name]:
            problems.append(f"依据数据复制后校验失败：{name}")

    # 指针相对本表所在目录（内部件/AHP依据对照表/）写，故上跳一层到内部件/依据数据/
    note_refs = [(f"../依据数据/{name}", path, use) for name, path, use in sources]

    # ---- 3. 溯源版 xlsx ----
    # 权重表溯源版取 07 的模板底稿（full 矩阵逐格写入、带公式，供回读对账），
    # 不再从成品静态表派生：成品是零公式干净表，派生出来的溯源版无从核对公式链。
    try:
        annotate_xlsx(str(ws / args.weights_ann), str(ws / REL_XLS_W_ANN),
                      [note_refs[0], note_refs[1]], args.demo)
        annotate_xlsx(str(ws / REL_XLS_S_CLEAN), str(ws / REL_XLS_S_ANN),
                      [note_refs[2], note_refs[0]], args.demo)
        wb_w = load_workbook(str(ws / REL_XLS_W_ANN), read_only=True)
        if "内部核对" not in wb_w.sheetnames:
            problems.append("溯源版权重表缺「内部核对」页")
        wb_w.close()
    except Exception as exc:                       # noqa: BLE001
        problems.append(f"溯源版 xlsx 生成失败：{type(exc).__name__}: {exc}")

    # ---- 4. 成品 PDF + 目录页码 ----
    try:
        pdf_info = render_and_fill(str(ws / REL_DOCX_CLEAN), str(ws / REL_PDF_CLEAN))
        entries, filled = toc_state(str(ws / REL_DOCX_CLEAN))
        info["pdf"] = pdf_info
        if filled < len(entries):
            problems.append(f"成品版目录页码回填不全：{filled}/{len(entries)} 条")
        back_ref = sorted(p for p in set(pdf_info["headingPages"].values())
                          if p in set(pdf_info["tocPages"]))
        if back_ref:
            problems.append(f"目录条目页码指向目录页自身（第 {back_ref} 页）")
        if pdf_info["emptyPages"]:
            problems.append(f"成品 PDF 含空白页：{pdf_info['emptyPages']}")
    except Exception as exc:                       # noqa: BLE001
        problems.append(f"成品 PDF 渲染失败：{type(exc).__name__}: {exc}")

    try:
        with tempfile.TemporaryDirectory(prefix="ahp_ann_") as tmp:
            ann_pdf = str(Path(tmp) / "ann.pdf")
            ann_info = render_and_fill(str(ws / REL_DOCX_ANN), ann_pdf)
        entries_ann, filled_ann = toc_state(str(ws / REL_DOCX_ANN))
        info["annotated"] = {"pages": ann_info["pages"], "filled": filled_ann,
                             "entries": len(entries_ann)}
        if filled_ann < len(entries_ann):
            problems.append(f"溯源版目录页码回填不全：{filled_ann}/{len(entries_ann)} 条")
    except Exception as exc:                       # noqa: BLE001
        problems.append(f"溯源版目录页码渲染失败：{type(exc).__name__}: {exc}")

    # ---- 5. 词表闸门（成品版 docx + PDF），豁免段按段落位置白名单 ----
    exemptions = (DOC.SYNTHETIC_NOTE, DOC.SCOPE_NOTE) if args.demo else ()
    sections = docx_sections(str(ws / REL_DOCX_CLEAN))
    if args.demo:
        info["whitelist"] = check_whitelist(sections, problems)
    scan_clean(docx_body_texts(str(ws / REL_DOCX_CLEAN), skip=exemptions),
               "成品版 docx", problems)
    if (ws / REL_PDF_CLEAN).is_file():
        pdf_text = re.sub(r"\s+", "", pdf_render.pdf_text(str(ws / REL_PDF_CLEAN)))
        for note in exemptions:
            pdf_text = pdf_text.replace(re.sub(r"\s+", "", note), "")
        scan_clean([pdf_text], "成品版 PDF", problems)

    # ---- 5b. 词表阳性对照：7 组样本必须条条被拦 ----
    info["positiveControl"] = positive_control()
    if not info["positiveControl"]["ok"]:
        problems.append(f"词表阳性对照未通过：漏拦 {info['positiveControl']['missed']}")

    # ---- 5c. 成品 xlsx 自检：零公式 / 无 fullCalcOnLoad / 封面五项 / 无模板示例残留 ----
    info["xlsx"] = []
    for rel, where in ((REL_XLS_W_CLEAN, "成品版权重表"), (REL_XLS_S_CLEAN, "成品版评分表")):
        if (ws / rel).is_file():
            info["xlsx"].append(scan_clean_workbook(ws / rel, where, problems))

    # ---- 6. 结构闸门：成品版一成果一子目录、文件齐、无杂项 ----
    pub = ws / "成果文件" / PUB
    expect_pub = {
        f"{BOOK_W}/{BOOK_W}_成品版.xlsx",
        f"{BOOK_S}/{BOOK_S}_成品版.xlsx",
        f"{BOOK_C}/{BOOK_C}_成品版.docx",
        f"{BOOK_C}/{BOOK_C}_成品版.pdf",
    }
    actual_pub = {p.relative_to(pub).as_posix() for p in pub.rglob("*") if p.is_file()} \
        if pub.is_dir() else set()
    missing = sorted(expect_pub - actual_pub)
    extra = sorted(actual_pub - expect_pub)
    if missing:
        problems.append(f"成品版缺件：{missing}")
    if extra:
        problems.append(f"成品版有非交付件：{extra}")
    for rel in (REL_XLS_W_ANN, REL_XLS_S_ANN, REL_DOCX_ANN, REL_XLS_W_CLEAN, REL_XLS_S_CLEAN):
        if not (ws / rel).is_file():
            problems.append(f"缺件：{rel}")

    info["ok"] = not problems
    out = ws / args.out_json
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"[AHP 交付装配] {'通过' if info['ok'] else '不通过'}："
          f"成品版 4 件 + 内部件 4 件，依据数据 4 份"
          + ("" if info["ok"] else f"，问题 {len(problems)} 条"))
    for item in problems[:10]:
        print("   -", item)
    return 0 if info["ok"] else 1


LEGACY_ROOT_SCAN = ("成果文件/AHP权重表.xlsx", "成果文件/AHP方案评分结论.xlsx",
                    "成果文件/AHP权重与评分结论.docx", "成果文件/AHP依据对照表.xlsx")


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
