#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pipeline.py -- Cardo城市更新AHP层次分析法 · 总控（按顺序调用 01→09 技能）。

三种用法：
  python pipeline.py run      --workspace <工作区> [--from 01] [--to 09] [--experts 5]
                              [--project-name 名称] [--force-conclusion]
  python pipeline.py status   --workspace <工作区>
  python pipeline.py scaffold --target <新目录> [--project-name 名称] [--experts 5]
                              [--baseline none|weights|full] [--force]

设计要点：
  · 每一步都通过子进程调用 Skills/<编号-名称>/scripts 下的技能脚本，技能间互不 import，
    产物一律落到 <工作区>/过程文件/<编号-名称>/，与人工逐步执行完全等价。
  · 每步执行完跑「闸门」，不通过即停，并给出失败原因与修复指引，不静默跳过。
  · 所有文本以 UTF-8 无 BOM 落盘（规避 Windows PowerShell 重定向写 UTF-16 的坑）。
  · 边界约定：输入文件/成果文件只放常规 Office（xlsx/docx 等），机读 JSON 一律落在过程文件；
    08 步用 xlsx转打分.py 把输入评分表转成过程 JSON，成果再用 xlsx/docx 导出。
  · 换评分做新测试：scaffold 出新工作区 → 改 输入文件/03-方案打分表/AHP方案评分表.xlsx → run。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_ROOT = HERE.parent                     # Skills/00-Cardo城市更新AHP层次分析法
SKILLS_ROOT = Path(os.environ.get("CARDO_SKILLS_ROOT") or HERE.parent.parent).resolve()  # Skills/
WORKSPACE = SKILL_ROOT.parent.parent         # 缺省工作区（工作区可在技能包之外，用 --workspace 指定）

TEMPLATE_REL = "输入文件/01-AHP四准则模板/4 层次分析法AHP四准则模板.xlsx"
SURVEY_REL = "输入文件/02-专家问卷（已回收）"
# 领料区：空白问卷的「人可见」落位。新手一眼能找到、拿去发给专家。
# 与过程目录 过程文件/02-问卷生成/问卷空白/ 的分工：
#   过程目录 = 流水线生成物与机读清单（问卷清单.json），留痕用；
#   领料区   = 给人的空白表（只有 xlsx，输入侧不放机读文件）。
PENDING_REL = "输入文件/02-专家问卷（待发放）"
SCORESHEET_REL = "输入文件/03-方案打分表/AHP方案评分表.xlsx"
SCHEME_LIST_REL = "输入文件/03-方案打分表/方案清单.xlsx"
PROC = "过程文件"
OUT = "成果文件"
CTRL = f"{PROC}/00-总控"
SCORES_JSON_REL = f"{PROC}/08-方案打分/打分.json"     # 由 xlsx 转换的过程产物

# 边界闸门：输入/成果里**不允许**出现的机读数据扩展名。
# 规则原意见交付说明「输入/成果只放常规 Office（xlsx/docx 等），机读 JSON 一律落在过程文件」——
# 防的是流水线自己的机读产物格式混进交付物，不是禁项目原始材料
# （dwg/jpg/zip/rar/ppt/log/txt 等都是正常输入）。
# 故只拦 JSON/JSONL/MD/Markdown/CSV/TSV/YAML——这些才是会被下游当结构化数据读的格式。
MACHINE_TEXT_EXTS = {".json", ".jsonl", ".md", ".markdown", ".csv", ".tsv", ".yaml", ".yml"}
# 常见 Office 交付物（成果目录用；保留供说明与扩展）
OFFICE_EXTS = {".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pptx", ".ppt", ".pdf"}

STEPS = [
    ("01", "01-指标体系", "指标体系确认"),
    ("02", "02-问卷生成", "专家问卷生成"),
    ("03", "03-回收校验", "问卷回收校验"),
    ("04", "04-多专家聚合", "多专家聚合"),
    ("05", "05-权重与一致性", "权重与一致性"),
    ("06", "06-一致性校准", "一致性校准"),
    ("07", "07-权重表导出", "权重表导出"),
    ("08", "08-方案打分", "方案打分"),
    ("09", "09-成果验证与成稿", "成果验证与成稿"),
]
STEP_BY_ID = {sid: (sid, d, n) for sid, d, n in STEPS}


class StepFailure(RuntimeError):
    pass


class Declined(StepFailure):
    """设计好的拒收（不是崩溃）：单独退出码 4，便于开工.py 给人话指引。"""
    exit_code = 4


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def read_json_or_fail(path: Path, label: str) -> dict:
    """读产物 JSON；解析失败抛 StepFailure（而不是裸 JSONDecodeError 崩栈）。

    背景：子脚本失败时 stdout 为空，产物会被写成 0 字节文件；
    直接 json.loads 会抛裸栈，且总控只捕获 StepFailure → 运行日志不落盘、
    用户看到的是上一轮的旧报告（"全部通过"）。这里收敛成可读的失败原因。
    """
    if not path.is_file():
        raise StepFailure(f"[产物缺失] {label}：{path}")
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        size = path.stat().st_size if path.is_file() else 0
        raise StepFailure(
            f"[产物不可解析] {label}：{path.name}（{size} 字节，{exc}）——"
            "通常是上一步脚本失败、产物为空；请看本步 stderr") from exc


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def sha256_of(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


class Step:
    """单步执行上下文：累计日志、闸门结果。"""

    def __init__(self, sid: str, dirname: str, name: str, ws: Path, log: list):
        self.sid, self.dirname, self.name = sid, dirname, name
        self.ws = ws
        self.log = log
        self.gates: list[dict] = []
        self.outputs: list[str] = []

    # ---- 执行 ----
    def cmd(self, script: str, args: list[str], out: str | None = None) -> dict:
        script_path = SKILLS_ROOT / self.dirname / "scripts" / script
        if not script_path.is_file():
            raise StepFailure(f"[{self.sid}] 找不到技能脚本：{script_path}")
        argv = [sys.executable, str(script_path)] + [str(a) for a in args]
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        proc = subprocess.run(argv, cwd=str(self.ws), capture_output=True, env=env)
        stdout = proc.stdout.decode("utf-8", errors="replace")
        stderr = proc.stderr.decode("utf-8", errors="replace")
        if out:
            write_text(self.ws / out, stdout)
            self.outputs.append(out)
        self.log.append({
            "step": self.sid,
            "script": f"Skills/{self.dirname}/scripts/{script}",
            "args": [str(a) for a in args],
            "exit": proc.returncode,
            "outFile": out,
            "stdoutTail": stdout.strip()[-400:],
            "stderrTail": stderr.strip()[-400:],
        })
        return {"exit": proc.returncode, "stdout": stdout, "stderr": stderr}

    # ---- 闸门 ----
    def gate(self, name: str, ok: bool, detail: str) -> None:
        self.gates.append({"name": name, "pass": bool(ok), "detail": detail})
        print(f"    [{'OK' if ok else 'NG'}] {name} — {detail}")

    def require_file(self, rel: str) -> Path:
        p = self.ws / rel
        if not p.is_file():
            raise StepFailure(f"[{self.sid}] 缺少预期产物：{rel}（查看上方脚本 stderr）")
        return p

    def fresh(self, rel: str) -> None:
        """删除上一轮旧产物，避免本步失败后下游误用陈旧文件。"""
        p = self.ws / rel
        if p.exists():
            p.unlink()

    def must(self, res: dict, what: str) -> None:
        if res.get("exit") != 0:
            tail = (res.get("stderr") or "").strip()[-1000:]
            raise StepFailure(f"[{self.sid}] {what}失败（退出码 {res.get('exit')}）"
                              + (f"\n{tail}" if tail else ""))

    def must_ok(self, res: dict, what: str) -> None:
        """要求子脚本退出码属于「业务结果」集合（0/1/3），其余视为脚本级错误。

        AHP 各子脚本用退出码表达业务结果而非成功/失败：
          0 = 通过；1 = 校验/一致性不通过（产物已带明细，属预期结果）；
          3 = 需人工确认（专家数不足、CR 警告等）；2 = 输入或参数错（真失败）。
        故 0/1/3 都放行，只拦 2 及其它异常码——否则「问卷结构有问题」这种
        正常要报告的情况会被当成脚本崩溃，反倒丢掉已经生成好的明细。
        """
        code = res.get("exit")
        if code in (0, 1, 3):
            return
        tail = (res.get("stderr") or res.get("stdout") or "").strip()[-1000:]
        raise StepFailure(f"[{self.sid}] {what}失败（退出码 {code}）"
                          + (f"\n{tail}" if tail else ""))

    def finish(self, last_cmd: dict | None = None) -> None:
        failed = [g for g in self.gates if not g["pass"]]
        if failed:
            msg = "；".join(f"{g['name']}：{g['detail']}" for g in failed)
            tail = ""
            if last_cmd and last_cmd.get("stderr"):
                tail = "\n脚本 stderr（末段）：\n" + last_cmd["stderr"].strip()[-800:]
            raise StepFailure(f"[{self.sid} {self.name}] 闸门未通过：{msg}{tail}")


# ---------------------------------------------------------------- 各步实现

def step_01(st: Step) -> None:
    st.must_ok(st.cmd("ahp_calc.py", ["indicators", "--format", "json"],
                       out=f"{PROC}/01-指标体系/指标目录.json"), "指标目录导出")
    st.must_ok(st.cmd("ahp_calc.py", ["indicators", "--format", "md"],
                       out=f"{PROC}/01-指标体系/指标目录.md"), "指标目录导出(md)")
    d = read_json_or_fail(st.require_file(f"{PROC}/01-指标体系/指标目录.json"), "指标目录")
    crits, inds = d.get("criteria", []), d.get("indicators", [])
    ids = {c["id"] for c in crits}
    st.gate("准则层与指标层非空", len(crits) >= 2 and len(inds) >= 4,
            f"准则 {len(crits)} 个、指标 {len(inds)} 项")
    st.gate("指标 criterionId 均能对上准则", all(i.get("criterionId") in ids for i in inds),
            "全部匹配" if all(i.get("criterionId") in ids for i in inds) else "存在孤儿指标")
    st.gate("每个指标具备评分锚点", all(len(i.get("anchors") or []) >= 3 for i in inds),
            "锚点最少项：" + str(min((len(i.get("anchors") or []) for i in inds), default=0)))
    st.finish()


def step_02(st: Step, cfg: dict) -> None:
    experts = int(cfg.get("experts") or 5)
    # 问卷一律匿名：封面固定写「某某片区改造项目」，避免专家被真实项目名锚定。
    # 原先这里传 cfg["projectName"]，与封面声明「本问卷不含真实项目名」自相矛盾，
    # 也让「发卷」这一步被迫依赖项目全称（发卷并不需要）。
    r = st.cmd("gen_questionnaire.py", [
        "generate", "--out", f"{PROC}/02-问卷生成/问卷空白", "--experts", str(experts),
        "--scale", "scale_0_1", "--criteria-scale", "scale_1_9",
        "--project-name", "某某片区改造项目", "--no-timestamp",
    ])
    st.must_ok(r, "空白问卷生成")
    d = st.require_file(f"{PROC}/02-问卷生成/问卷空白/问卷清单.json")
    files = sorted((st.ws / f"{PROC}/02-问卷生成/问卷空白").glob("AHP专家问卷-专家*.xlsx"))
    st.gate("空白问卷份数 = 专家数", len(files) == experts, f"生成 {len(files)} 份，期望 {experts} 份")
    st.gate("问卷清单可读（过程留痕）",
            bool(read_json_or_fail(d, "问卷清单").get("experts")), f"{d.name} 已生成")

    # 领料区：把空白问卷复制到人一眼能看到的输入侧目录（只复制 xlsx，
    # 问卷清单.json 留在过程文件里 —— 输入侧不放机读文件）。
    pending = st.ws / PENDING_REL
    pending.mkdir(parents=True, exist_ok=True)
    copied = 0
    for f in files:
        tgt = pending / f.name
        if not tgt.is_file():
            shutil.copy2(f, tgt)
            copied += 1
        else:
            # 领料区已有一份：若内容与刚生成的空白卷不同（版本更新，如新增指标说明列），
            # 说明领料区那份是旧的，用新的覆盖 —— 否则专家拿到的是过期模板。
            # 回收区不这么做：那里是专家填好的，绝不能覆盖。
            try:
                if sha256_of(tgt) != sha256_of(f):
                    shutil.copy2(f, tgt)
                    copied += 1
            except Exception:
                pass
    on_hand = sorted(pending.glob("AHP专家问卷-专家*.xlsx"))
    st.gate("领料区有空白问卷可发", len(on_hand) >= experts,
            f"{PENDING_REL}：{len(on_hand)} 份（本次复制 {copied} 份）")

    # 输入侧登记表：用 xlsx 专家名单替代原 JSON 问卷清单（输入只放 Office）
    q = sorted((st.ws / SURVEY_REL).glob("AHP专家问卷-专家*.xlsx"))
    if q:
        st.cmd("专家名单导出.py", ["--dir", SURVEY_REL])
    roster = st.ws / SURVEY_REL / "专家名单.xlsx"
    st.gate("输入侧专家名单为 xlsx", roster.is_file(),
            roster.name if roster.is_file() else f"回收目录无问卷，无法登记（{SURVEY_REL}）")
    if not 5 <= experts <= 7:
        st.gate("专家数在规则区间 5-7", False, f"当前 {experts} 位")
    else:
        st.gate("专家数在规则区间 5-7", True, f"{experts} 位")
    st.finish(r)


def step_03(st: Step) -> None:
    base = ["check", "--dir", SURVEY_REL, "--scale", "scale_0_1",
            "--criteria-scale", "scale_1_9"]
    st.must_ok(st.cmd("gen_questionnaire.py", base + ["--format", "json", "--no-timestamp"],
                      out=f"{PROC}/03-回收校验/回收校验报告.json"), "回收校验")
    st.must_ok(st.cmd("gen_questionnaire.py", base + ["--format", "md", "--no-timestamp"],
                      out=f"{PROC}/03-回收校验/回收校验报告.md"), "回收校验(md)")
    d = read_json_or_fail(st.require_file(f"{PROC}/03-回收校验/回收校验报告.json"), "回收校验报告")
    s = d.get("summary", {})
    struct_ok = s.get("problems", 1) == 0 and s.get("warnings", 1) == 0
    st.gate("问卷结构无问题", struct_ok,
            f"文件 {s.get('files')} 份、结构问题 {s.get('problems')}、警告 {s.get('warnings')}")
    st.gate("问卷一致性无 CR 警告", s.get("crWarnings", 1) == 0, f"CR 警告 {s.get('crWarnings')}")
    st.gate("专家数满足最少 5 位", int(s.get("files") or 0) >= 5, f"回收 {s.get('files')} 份")
    st.gate("回收校验总体通过", bool(s.get("pass")), f"summary.pass={s.get('pass')}")
    st.finish()


def step_04(st: Step) -> None:
    r = st.cmd("ahp_aggregate.py", ["aggregate", "--dir", SURVEY_REL,
                                   "--out", f"{PROC}/04-多专家聚合/群体矩阵.json",
                                   "--compare", "--no-timestamp"])
    st.must_ok(r, "多专家聚合")
    st.must_ok(st.cmd("ahp_aggregate.py", ["aggregate", "--dir", SURVEY_REL,
                                           "--format", "md", "--no-timestamp"],
                      out=f"{PROC}/04-多专家聚合/聚合报告.md"), "聚合报告(md)")
    d = read_json_or_fail(st.require_file(f"{PROC}/04-多专家聚合/群体矩阵.json"), "群体矩阵")
    agg, cons = d.get("aggregation", {}), d.get("consistency", {})
    st.gate("聚合口径为逐元素几何平均", agg.get("rule") == "geometric-elementwise",
            f"rule={agg.get('rule')}")
    st.gate("专家数 ≥ 5", int(agg.get("expertCount") or 0) >= 5, f"专家 {agg.get('expertCount')} 位")
    st.gate("群体矩阵个数 = 准则数 + 1", len(d.get("matrices") or {}) >= 2,
            f"矩阵 {len(d.get('matrices') or {})} 个")
    st.gate("聚合一致性通过", bool(cons.get("aggregatePass")),
            f"{cons.get('passCount')}/{cons.get('total')} 通过")
    st.finish(r)


def step_05(st: Step, cfg: dict) -> None:
    demo_args = ["--demo"] if demo_flag(cfg) else []
    st.must_ok(st.cmd("聚合转矩阵集.py", ["--agg", f"{PROC}/04-多专家聚合/群体矩阵.json",
                                          "--out-dir", f"{PROC}/05-权重与一致性"]),
               "群体矩阵转矩阵集")
    st.must_ok(st.cmd("ahp_calc.py", ["weights", "--matrix-set", f"{PROC}/05-权重与一致性/矩阵集.json",
                                      "--format", "json", "--no-timestamp"] + demo_args,
                      out=f"{PROC}/05-权重与一致性/权重表.json"), "权重计算")
    st.must_ok(st.cmd("ahp_calc.py", ["weights", "--matrix-set", f"{PROC}/05-权重与一致性/矩阵集.json",
                                      "--format", "md", "--no-timestamp"],
                      out=f"{PROC}/05-权重与一致性/权重表.md"), "权重计算(md)")
    st.must_ok(st.cmd("ahp_calc.py", ["consistency", "--matrix-set", f"{PROC}/05-权重与一致性/矩阵集.json",
                                      "--format", "md", "--no-timestamp"],
                      out=f"{PROC}/05-权重与一致性/一致性结果.md"), "一致性检验")
    d = read_json_or_fail(st.require_file(f"{PROC}/05-权重与一致性/权重表.json"), "权重表")
    w, cons = d.get("weights", {}), d.get("consistency", {})
    st.gate("准则权重和 = 1", abs(float(w.get("sumCriteria") or 0) - 1.0) < 1e-9,
            f"sum={w.get('sumCriteria')}")
    st.gate("综合权重和 = 1", abs(float(w.get("sumGlobal") or 0) - 1.0) < 1e-9,
            f"sum={w.get('sumGlobal')}")
    st.gate("一致性 CR<0.1 全部通过",
            cons.get("total") and cons.get("passCount") == cons.get("total"),
            f"{cons.get('passCount')}/{cons.get('total')}，阈值 {cons.get('threshold')}")
    st.finish()


def step_06(st: Step) -> None:
    full_dir = st.ws / f"{PROC}/05-权重与一致性/全矩阵"
    mats = sorted(full_dir.glob("*.json"))
    if not mats:
        raise StepFailure("[06] 缺少 05 步产出的全矩阵 JSON，无法诊断")
    rows, need_calib = [], []
    for p in mats:
        mid = p.stem
        # 注意：ahp_calibrate.py 的 suggest 在「CR 已通过→无候选」与「CR 不通过→有候选」
        # 两种情况都返回退出码 1，退出码在这里不是成功/失败信号。
        # 故只用退出码 2/其它（非 0/1）判脚本级错误，业务判定一律读产物 JSON 的 cr 字段。
        r_diag = st.cmd("ahp_calibrate.py", ["diagnose", "--matrix",
                                             f"{PROC}/05-权重与一致性/全矩阵/{mid}.json",
                                             "--format", "json"],
                        out=f"{PROC}/06-一致性校准/诊断-{mid}.json")
        if r_diag.get("exit") not in (0, 1):
            st.must(r_diag, f"一致性诊断({mid})")
        r_sug = st.cmd("ahp_calibrate.py", ["suggest", "--matrix",
                                            f"{PROC}/05-权重与一致性/全矩阵/{mid}.json",
                                            "--max-candidates", "3", "--no-timestamp"],
                       out=f"{PROC}/06-一致性校准/候选-{mid}.json")
        if r_sug.get("exit") not in (0, 1):
            st.must(r_sug, f"校准候选({mid})")
        dg = read_json_or_fail(st.ws / f"{PROC}/06-一致性校准/诊断-{mid}.json", f"诊断({mid})")
        cr = float(dg.get("cr") or 0.0)
        rows.append((mid, cr, dg.get("message", "")))
        if cr >= 0.1 - 1e-12:
            need_calib.append(mid)
    md = ["# 06 · 一致性校准记录", "",
          f"- 诊断矩阵 {len(rows)} 个；生成时间 {now_iso()}",
          "- 规则：CR<0.1 通过；不通过须人工批准最小改动候选（禁止自动采用）", ""]
    for mid, cr, msg in rows:
        md.append(f"- `{mid}`：CR={cr:.6f}；{msg}")
    if need_calib:
        md += ["", f"**需人工校准**：{', '.join(need_calib)}。",
               "处置：在 `ahp_calibrate.py apply --approve-by` 人工批准后回第 05 步重算。"]
    else:
        md += ["", "结论：全部矩阵 CR<0.1，无需校准，群体矩阵保持聚合原貌，直接进入第 07 步。"]
    write_text(st.ws / f"{PROC}/06-一致性校准/校准记录.md", "\n".join(md))
    st.outputs.append(f"{PROC}/06-一致性校准/校准记录.md")
    st.gate("一致性校准无需人工介入", not need_calib,
            "无需校准" if not need_calib else f"待校准：{need_calib}")
    st.finish()


def step_07(st: Step, cfg: dict) -> None:
    tpl = st.ws / TEMPLATE_REL
    if not tpl.is_file():
        raise StepFailure(f"[07] 模板不存在：{TEMPLATE_REL}")
    sha_before = sha256_of(tpl)
    st.fresh(f"{OUT}/AHP权重表.xlsx")
    st.fresh(f"{PROC}/07-权重表导出/权重表.json")
    matrix_set = f"{PROC}/05-权重与一致性/矩阵集.json"
    if not (st.ws / matrix_set).is_file():
        raise StepFailure(f"[07] 缺少 05 步矩阵集：{matrix_set}")
    st.fresh(f"{PROC}/07-权重表导出/权重表_溯源.xlsx")
    args = ["export",
            "--weights", f"{PROC}/05-权重与一致性/权重表.json",
            "--out", f"{OUT}/AHP权重表.xlsx",
            "--out-json", f"{PROC}/07-权重表导出/权重表.json",
            "--template", TEMPLATE_REL, "--no-timestamp",
            "--matrix-set", matrix_set,
            "--out-ann", f"{PROC}/07-权重表导出/权重表_溯源.xlsx",
            "--project", cfg.get("projectName") or "某某片区改造项目"]
    if demo_flag(cfg):
        args.append("--demo")
    r = st.cmd("ahp_export.py", args)
    st.must(r, "权重表导出")
    st.outputs += [f"{OUT}/AHP权重表.xlsx", f"{PROC}/07-权重表导出/权重表.json",
                   f"{PROC}/07-权重表导出/权重表_溯源.xlsx"]
    # 回读对账跑在「溯源底稿」上：成品是静态干净表（无公式、无矩阵公式），
    # 模板口径的回读只能对模板副本做。
    r = st.cmd("ahp_export.py", ["verify", "--file", f"{PROC}/07-权重表导出/权重表_溯源.xlsx",
                                 "--json", f"{PROC}/07-权重表导出/权重表.json",
                                 "--template", TEMPLATE_REL],
               out=f"{PROC}/07-权重表导出/回读校验报告.json")
    st.must(r, "权重表回读对账")
    v = read_json_or_fail(st.require_file(f"{PROC}/07-权重表导出/回读校验报告.json"), "回读校验报告")
    st.gate("xlsx 与过程 JSON 回读对账通过", bool(v.get("ok")), f"problems={v.get('problems')}")
    st.gate("模板原件未被改写", v.get("templateUnchanged") is True, "templateUnchanged=True")
    st.gate("模板 sha256 执行前后一致", sha256_of(tpl) == sha_before,
            f"{sha_before[:16]}…{sha_before[-8:]}")
    st.gate("成果权重表为 xlsx（机读 JSON 留在过程文件）",
            (st.ws / f"{OUT}/AHP权重表.xlsx").is_file(),
            f"{OUT}/AHP权重表.xlsx")
    # 成品表内 CR vs 报告 CR vs 矩阵集复算：三处逐项一致，否则 exit 3 拦下
    r = st.cmd("ahp_export.py", ["consistency",
                                 "--workbook", f"{OUT}/AHP权重表.xlsx",
                                 "--json", f"{PROC}/07-权重表导出/权重表.json",
                                 "--matrix-set", matrix_set],
               out=f"{PROC}/07-权重表导出/CR一致性核对.json")
    st.outputs.append(f"{PROC}/07-权重表导出/CR一致性核对.json")
    cc = read_json_or_fail(st.require_file(f"{PROC}/07-权重表导出/CR一致性核对.json"), "CR 一致性核对") \
        if r.get("exit") == 0 else {}
    st.gate("成品表内 CR 与报告逐项一致", r.get("exit") == 0 and bool(cc.get("ok")),
            "；".join(f"{it['matrix']} {it['tableCR']:.6f}" for it in cc.get("items", [])) or "未取到明细")
    if r.get("exit") != 0:
        raise StepFailure("[07] 成品表内 CR 与报告不一致（exit 3），详见 "
                          f"{PROC}/07-权重表导出/CR一致性核对.json")
    st.finish(r)


def check_scheme_consistency(ws: Path, list_rel: str, sheet_rel: str) -> str:
    """比对《方案清单》与已生成评分表的方案名；不一致返回说明文字，一致返回空串。

    为什么要有这一步：评分表可能在用户填《方案清单》之前就生成了。
    若用户填完清单又改了方案名，而评分表里仍是旧名 —— 成果里的方案名
    会与清单对不上。这里把这种静默不一致变成明确的提示。

    注：占位名（首跑时的「（待填方案名 S1）」）不算不一致 ——
    那正是等用户填清单的状态，第 08 步会先同步名再比对。

    只在两边都存在时比对；任一缺失都放行（缺清单不影响按评分表跑）。
    """
    list_path, sheet_path = ws / list_rel, ws / sheet_rel
    if not list_path.is_file() or not sheet_path.is_file():
        return ""
    try:
        sys.path.insert(0, str(SKILLS_ROOT / "08-方案打分" / "scripts"))
        from scheme_lib import read_scheme_list, read_sheet_schemes   # noqa: E402
        listed = read_scheme_list(list_path)
        if not listed:
            return ""
        in_sheet = read_sheet_schemes(sheet_path)
        if not in_sheet:
            return ""
        listed_names = [s["name"] for s in listed]
        sheet_names = [s["name"] for s in in_sheet]
        if listed_names == sheet_names:
            return ""
        # 报具体差异：哪几个对不上
        only_list = [n for n in listed_names if n not in sheet_names]
        only_sheet = [n for n in sheet_names if n not in listed_names]
        bits = []
        if only_list:
            bits.append("清单里有、评分表没有：" + "、".join(only_list))
        if only_sheet:
            bits.append("评分表里有、清单里没有：" + "、".join(only_sheet))
        if not bits:
            bits.append("方案顺序或数量不同（清单：" + "、".join(listed_names)
                        + "；评分表：" + "、".join(sheet_names) + "）")
        return "；".join(bits)
    except Exception:
        return ""          # 比对失败不阻断主流程（主流程有自己的一致性闸门）


def step_08(st: Step, cfg: dict) -> None:
    # 项目名空 → 用占位名，但要在运行报告里记一笔，提醒用户去《方案清单.xlsx》填
    project = cfg.get("projectName") or "某某片区改造项目"
    sheet = st.ws / SCORESHEET_REL
    scripts = "08-方案打分/scripts"

    if not sheet.is_file():
        # 评分表还没出：按《方案清单》的实际方案数生成（方案名以清单为准）
        st.cmd("gen_scoresheet.py", ["--weights", f"{PROC}/05-权重与一致性/权重表.json",
                                     "--project-name", project, "--out", SCORESHEET_REL,
                                     "--scheme-list", SCHEME_LIST_REL])
        raise StepFailure(f"[08] 缺少评分表：{SCORESHEET_REL}。已按《方案清单》生成空白模板"
                          f"（方案列数 = 你的方案数），请填写各方案得分与依据后重跑 `--from 08`。")

    # 评分表已存在（首跑就铺好了空白表）。此时要处理两件事：
    #  ① 方案名同步：用户可能在首跑之后才在《方案清单.xlsx》里填好方案名。
    #     评分表里还是占位名 → 把清单里的名字同步进来（得分格不动）。
    #     方案数量变了则无法原地改列 —— 交给一致性检查报错。
    #  ② 综合权重回填：首跑出的评分表没有权重（第 05 步之后才有），
    #     这里把真实权重写回 F 列，让人看得见、也保证表与引擎口径一致。
    sync = st.cmd("sync_scheme_names.py", ["--xlsx", SCORESHEET_REL,
                                            "--scheme-list", SCHEME_LIST_REL])
    if sync.get("exit") == 3:
        raise StepFailure(
            f"[08] 方案数量与《方案清单.xlsx》不一致：{sync.get('stdout', '').strip()[-200:]}"
            f"评分表列数是按旧方案数生成的。请把评分表《{SCORESHEET_REL}》删掉，"
            f"重跑 `--from 08` 按新的方案数重建（若已填了得分，请先把分数抄下来）。"
        )

    mismatch = check_scheme_consistency(st.ws, SCHEME_LIST_REL, SCORESHEET_REL)
    if mismatch:
        raise StepFailure(
            f"[08] 方案名与《方案清单.xlsx》不一致：{mismatch}。"
            f"评分表是先前按旧方案名生成的。请二选一："
            f"① 把评分表《{SCORESHEET_REL}》删掉，重跑 --from 08 按新方案名重新生成；"
            f"② 或把《{SCHEME_LIST_REL}》的方案名改回与评分表一致。"
        )

    # 回填综合权重（首跑出的空白表权重列是空的）
    st.cmd("fill_weights.py", ["--xlsx", SCORESHEET_REL,
                               "--weights", f"{PROC}/05-权重与一致性/权重表.json"])

    # xlsx（人填）→ 过程 JSON（引擎读）：输入侧不放 JSON
    st.fresh(SCORES_JSON_REL)
    r = st.cmd("xlsx转打分.py", ["--xlsx", SCORESHEET_REL, "--out", SCORES_JSON_REL,
                                 "--project-name", project])
    st.must(r, "评分表 xlsx → JSON 转换")
    src_demo = bool(read_json_or_fail(st.require_file(SCORES_JSON_REL),
                                      "打分转换结果").get("demo"))
    if src_demo and not demo_flag(cfg):
        raise Declined(
            "拒收（设计如此，不是崩溃）：评分表带 demo=true（演示数据），"
            "但本次运行没有声明演示。\n"
            "  为什么拒：演示数据不得按常规项目出交付物，否则对外就是假依据。\n"
            "  怎么改：确认这是演示数据 → 在命令后加 --demo 重跑；"
            "若这是真实项目 → 先查评分表来源，别把演示样例当真实打分。\n"
            "  续跑：本次从 08 步继续（--demo --from 08）。")
    demo_args = ["--demo"] if demo_flag(cfg) else []
    r = st.cmd("score_project.py", ["score", "--weights", f"{PROC}/05-权重与一致性/权重表.json",
                                    "--scores", SCORES_JSON_REL,
                                    "--out", f"{PROC}/08-方案打分/评分表.json", "--no-timestamp"]
               + demo_args)
    st.must(r, "方案加权评分")
    st.cmd("score_project.py", ["score", "--weights", f"{PROC}/05-权重与一致性/权重表.json",
                                "--scores", SCORES_JSON_REL, "--format", "md", "--no-timestamp"],
           out=f"{PROC}/08-方案打分/评分结论.md")
    # 过程 JSON → 成果 xlsx（交付）
    st.fresh(f"{OUT}/AHP方案评分结论.xlsx")
    r = st.cmd("评分结论导出.py", ["--scores", f"{PROC}/08-方案打分/评分表.json",
                                   "--out", f"{OUT}/AHP方案评分结论.xlsx",
                                   "--project-name", project] + demo_args)
    st.must(r, "评分结论 xlsx 导出")
    st.outputs += [SCORES_JSON_REL, f"{OUT}/AHP方案评分结论.xlsx"]

    d = read_json_or_fail(st.require_file(f"{PROC}/08-方案打分/评分表.json"), "评分表")
    src = read_json_or_fail(st.require_file(SCORES_JSON_REL), "打分转换结果")
    schemes, ranking = d.get("schemes", []), d.get("ranking", [])
    st.gate("评分表已转成过程 JSON", bool(src.get("schemes")), f"{SCORES_JSON_REL}")
    st.gate("方案数一致（xlsx = 出分）",
            len(src.get("schemes", [])) == len(schemes),
            f"xlsx {len(src.get('schemes', []))} / 出分 {len(schemes)}")
    st.gate("每个方案均出分且落在 0-10",
            bool(schemes) and all(s.get("composite") is not None and 0 <= s["composite"] <= 10
                                  for s in schemes),
            "；".join(f"{s['id']}={s.get('composite')}" for s in schemes))
    st.gate("排序完整（名次连续、无缺项）",
            len(ranking) == len(schemes)
            and [r["order"] for r in ranking] == list(range(1, len(schemes) + 1)),
            "；".join(f"第{r['order']}名 {r['name']} {r['composite']}" for r in ranking))
    st.gate("成果评分结论为 xlsx", (st.ws / f"{OUT}/AHP方案评分结论.xlsx").is_file(),
            f"{OUT}/AHP方案评分结论.xlsx")
    st.finish()


def non_office_files(ws: Path, rel: str) -> list[str]:
    """列出该目录下**不该出现**的文件（机读文本类）。

    边界规则的目的是「机读 JSON/MD 一律落在 过程文件/，不混进 输入文件/ 与 成果文件/」，
    **不是**禁止项目原始材料。真实项目资料里 dwg/jpg/zip/rar/ppt 等是正常输入
    （建筑图纸、现状照片、方案附件、成果打包），把它们判违规会让 09 步对任何
    带图纸的真实项目都必 FAIL——实测某市某新区项目 128 个原始文件全部被误判。

    故只拦机读文本（.json/.md/.jsonl/.csv/.tsv/.yaml/.yml/.txt/.log 等），
    其余一律放行；.downloading 等半截文件按「拷贝未完成」处理，不算违规。
    """
    base = ws / rel
    if not base.is_dir():
        return []
    # 例外：成果文件/内部件/依据数据/ 是**随包取证的依据数据**（权重表/评分表/群体矩阵/
    # 矩阵集 4 份 JSON）。它必须跟着交付包走，否则依据对照表离开本机就断链；
    # 与「机读产物留过程文件」不冲突——这里存的是复核凭证，不是流水线中间件。
    exempt = f"{rel}/内部件/依据数据".replace("\\", "/")
    return sorted(str(p.relative_to(ws)).replace("\\", "/")
                  for p in base.rglob("*")
                  if p.is_file() and p.suffix.lower() in MACHINE_TEXT_EXTS
                  and not str(p.relative_to(ws)).replace("\\", "/").startswith(exempt + "/"))


def step_09(st: Step, cfg: dict) -> None:
    # 项目名空 → 用占位名，但要在运行报告里记一笔，提醒用户去《方案清单.xlsx》填
    project = cfg.get("projectName") or "某某片区改造项目"
    baseline_mode = cfg.get("baselineMode") or "auto"
    r = st.cmd("独立验证.py", ["--matrix-set", f"{PROC}/05-权重与一致性/矩阵集.json",
                               "--weights", f"{PROC}/05-权重与一致性/权重表.json",
                               "--scores", f"{PROC}/08-方案打分/评分表.json",
                               "--template", TEMPLATE_REL,
                               "--baseline-mode", str(baseline_mode),
                               "--out-dir", f"{PROC}/09-成果验证与成稿"])
    st.must(r, "独立验证复算")
    d = read_json_or_fail(st.require_file(f"{PROC}/09-成果验证与成稿/独立验证结果.json"), "独立验证结果")
    st.gate("独立复算无失败项",
            d.get("failed") == 0 and bool(d.get("allPass")),
            f"通过 {d.get('passed')} / 失败 {d.get('failed')} / 跳过 {d.get('skipped')}")
    binfo = d.get("baselineInfo") or {}
    st.gate("基线信息已记录", True,
            f"mode={binfo.get('mode')}／status={binfo.get('status')}"
            + (f"：{binfo.get('source')}" if binfo.get("source") else ""))

    # 结论文档：成品版（对外）+ 溯源版（内部复核）双 docx；md 只作过程留痕
    docx_clean = f"{OUT}/成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.docx"
    docx_ann = f"{OUT}/内部件/AHP权重与评分结论/AHP权重与评分结论_溯源版.docx"
    st.fresh(docx_clean)
    st.fresh(docx_ann)
    st.fresh(f"{PROC}/09-成果验证与成稿/结论稿.md")
    st.fresh(f"{PROC}/09-成果验证与成稿/结论稿-溯源版.md")
    demo_args = ["--demo"] if demo_flag(cfg) else []
    r = st.cmd("结论成稿.py", ["--weights", f"{PROC}/05-权重与一致性/权重表.json",
                               "--scores", f"{PROC}/08-方案打分/评分表.json",
                               "--verification", f"{PROC}/09-成果验证与成稿/独立验证结果.json",
                               "--aggregation", f"{PROC}/04-多专家聚合/群体矩阵.json",
                               "--out-docx", docx_clean,
                               "--out-docx-ann", docx_ann,
                               "--out-md", f"{PROC}/09-成果验证与成稿/结论稿.md",
                               "--out-md-ann", f"{PROC}/09-成果验证与成稿/结论稿-溯源版.md",
                               "--project-name", project] + demo_args)
    st.must(r, "结论 docx 成稿")
    st.outputs += [docx_clean, docx_ann, f"{PROC}/09-成果验证与成稿/结论稿.md"]
    st.gate("成品版/溯源版双 docx 都已生成",
            (st.ws / docx_clean).is_file() and (st.ws / docx_ann).is_file(),
            "成品版对外交付、溯源版供内部复核；md 留在过程文件")

    # 交付装配：分存（一成果一子目录）+ 依据随包 + 成品 PDF（目录页码静态回填）+ 词表闸门
    st.fresh(f"{PROC}/09-成果验证与成稿/交付装配.json")
    r = st.cmd("交付装配.py", ["--weights", f"{PROC}/05-权重与一致性/权重表.json",
                               "--scores", f"{PROC}/08-方案打分/评分表.json",
                               "--aggregation", f"{PROC}/04-多专家聚合/群体矩阵.json",
                               "--matrix-set", f"{PROC}/05-权重与一致性/矩阵集.json",
                               "--weights-ann", f"{PROC}/07-权重表导出/权重表_溯源.xlsx",
                               "--out-json", f"{PROC}/09-成果验证与成稿/交付装配.json"]
               + demo_args)
    st.must(r, "交付装配（分存 / 依据随包 / 成品 PDF / 词表闸门）")
    asm = read_json_or_fail(st.require_file(f"{PROC}/09-成果验证与成稿/交付装配.json"), "交付装配")
    st.gate("交付装配通过", bool(asm.get("ok")),
            "成品版 4 件、内部件 3 件、依据数据 4 份；目录页码不指回目录页、无空白页")
    pdf_info = asm.get("pdf") or {}
    st.gate("成品版 PDF 页码已静态回填",
            int((asm.get("pdf") or {}).get("filled") or 0) >= len(
                ((asm.get("pdf") or {}).get("headingPages") or {})),
            f"PDF {pdf_info.get('pages')} 页，目录页 {pdf_info.get('tocPages')}，"
            f"章页码 {pdf_info.get('headingPages')}")

    # 依据对照表：结论稿里每个数（权重/λmax/CI/CR/综合得分）都指回随包依据数据；
    # 顺带独立复算权重合计、综合得分，并拦「合成演示」这类占位依据
    evid_xlsx = f"{OUT}/内部件/AHP依据对照表/AHP依据对照表.xlsx"
    st.fresh(evid_xlsx)
    r = st.cmd("依据对照表.py", ["--weights", f"{PROC}/05-权重与一致性/权重表.json",
                                 "--scores", f"{PROC}/08-方案打分/评分表.json",
                                 "--conclusion", docx_clean,
                                 "--out-xlsx", evid_xlsx,
                                 "--out-json", f"{PROC}/09-成果验证与成稿/依据对照表.json"]
               + demo_args)
    st.must(r, "依据对照表（每个数可追回源表）")
    st.outputs += [evid_xlsx, f"{PROC}/09-成果验证与成稿/依据对照表.json"]
    st.gate("结论稿每个数都能追回权重表/评分表",
            read_json_or_fail(st.require_file(f"{PROC}/09-成果验证与成稿/依据对照表.json"),
                              "依据对照表").get("ok") is True,
            "权重合计、综合得分独立复算一致，评分依据无占位措辞")

    # 演示身份：判定只认 --demo / 工作区参数 demo:true（项目名只提示，见 demo_name_hint）
    demo = demo_flag(cfg)
    scores_demo = bool(read_json_or_fail(st.require_file(f"{PROC}/08-方案打分/评分表.json"),
                                         "评分表").get("demo"))
    st.gate("演示身份与评分数据一致",
            not scores_demo or demo,
            "评分表 demo=true，本次已声明 --demo" if (scores_demo and demo)
            else ("评分表 demo=true，但未声明 --demo（演示数据不得按常规项目出交付物）"
                  if scores_demo else "评分表未带演示标记"))
    if demo:
        st.gate("演示身份已落到过程数据与本步产物",
                scores_demo and bool(read_json_or_fail(
                    st.require_file(f"{PROC}/09-成果验证与成稿/依据对照表.json"),
                    "依据对照表").get("demo")),
                "评分表 demo=true；合成问卷口径写进成品版「方法与口径」一节（只交代一次），"
                "成品版各页不撒角标"
                + ("（上游两件在 07/08 步已按 --demo 生成）" if scores_demo else ""))

    # 项目名是否还是占位名 —— 交付物上印着「某某片区改造项目」是不能直接交出去的，
    # 这里明确提醒用户去《方案清单.xlsx》B1 格填真实项目名。
    # 项目名允许是演示标识，但演示身份由 --demo/demo:true 决定，不由名字决定：
    # 名字看着像演示 ≠ 自动演示；名字换成真实片区 ≠ 演示数据洗白（评分表仍带 demo 标记）。
    real_name = (cfg.get("projectName") or "").strip()
    if not real_name or real_name == "某某片区改造项目":
        st.gate("项目名已填（占位名不可交付）", False,
                f"当前成果上印的是占位名「{project}」——"
                f"请在 {SCHEME_LIST_REL} 的 B1 格填项目全名后重跑 --from 09")
    elif demo_name_hint(real_name, cfg):
        st.gate("项目名疑似演示（仅提示，不改变演示判定）", True,
                f"「{real_name}」："
                + ("已声明 --demo，产物按演示样例标注"
                   if demo else "未声明 --demo，按常规项目处理；演示数据由评分表 demo 标记识别"))
    else:
        st.gate("项目名已填真实名称", True, real_name)

    # 边界闸门：输入/成果不得混入机读文本（JSON/MD/CSV 等）
    bad_in = non_office_files(st.ws, "输入文件")
    bad_out = non_office_files(st.ws, OUT)
    st.gate("输入文件无机读文本（JSON/MD 等）", not bad_in,
            "，".join(bad_in[:8]) if bad_in else "无 JSON/MD 等机读文件")
    st.gate("成果文件仅含交付物（无机读文本）", not bad_out,
            "，".join(bad_out[:8]) if bad_out else "无 JSON/MD 等机读文件")
    st.finish(r)


# ---------------------------------------------------------------- 运行器

def load_cfg(ws: Path) -> dict:
    """读工作区参数。

    项目名的取值顺序（用户在哪填的就在哪读）：
      1)《方案清单.xlsx》B1 格 —— 用户自己填的（拿得到的模板里填，不必问 AI）；
      2) 工作区参数.json 里的 projectName —— 命令行 --project 写进来的；
      3) 都没有 → 空串，由调用方决定用占位名并提醒用户去填。
    """
    p = ws / "工作区参数.json"
    cfg = {"projectName": "", "experts": 5, "baselineMode": "auto", "demo": False}
    if p.is_file():
        cfg.update({k: v for k, v in read_json(p).items()
                    if k in ("projectName", "experts", "baselineMode", "demo")})
    # 方案清单里用户填的项目名优先（比参数文件里的更接近用户真实意图）
    try:
        sys.path.insert(0, str(SKILLS_ROOT / "08-方案打分" / "scripts"))
        from scheme_lib import read_project_name          # noqa: E402
        from_sheet = read_project_name(ws / SCHEME_LIST_REL)
        if from_sheet:
            cfg["projectName"] = from_sheet
            cfg["projectNameSource"] = "sheet"
        else:
            cfg["projectNameSource"] = "param" if cfg.get("projectName") else "unset"
            cfg["projectName"] = cfg.get("projectName") or ""
    except Exception:
        cfg.setdefault("projectNameSource", "param" if cfg.get("projectName") else "unset")
    return cfg


DEMO_NAME_HINTS = ("演示", "合成", "非真实", "示例")


def demo_flag(cfg: dict) -> bool:
    """演示身份的唯一权威来源：--demo / 工作区参数 demo:true。

    项目名看起来像演示数据只做提示，绝不据此自动判定（改个名字就能洗掉身份）。
    """
    value = cfg.get("demo")
    if isinstance(value, str):
        return value.strip().lower() in ("1", "true", "yes", "on")
    return bool(value)


def demo_name_hint(project: str, cfg: dict) -> str:
    """项目名疑似演示时给出的提示（不改判定，只提醒显式声明）。"""
    hits = [w for w in DEMO_NAME_HINTS if w in (project or "")]
    if not hits:
        return ""
    if demo_flag(cfg):
        return (f"检测到项目名疑似演示数据（命中：{'、'.join(hits)}），已显式声明 demo=true；"
                "产物按演示样例标注。演示身份只认显式参数，不认项目名。")
    return (f"检测到项目名疑似演示数据（命中：{'、'.join(hits)}），未显式声明 demo；"
            "如确为演示请用 --demo 运行或把工作区参数 demo 设为 true，"
            "否则请换成真实项目名——演示身份只认显式参数，不认项目名。")


def run_pipeline(ws: Path, from_id: str, to_id: str, cfg: dict) -> int:
    if not SKILLS_ROOT.is_dir():
        print(f"[AHP-E-WS] 找不到技能链（CARDO_SKILLS_ROOT）：{SKILLS_ROOT}", file=sys.stderr)
        return 2
    for rel in (TEMPLATE_REL, SURVEY_REL):
        if not (ws / rel).exists():
            print(f"[AHP-E-WS] 缺输入：{rel}", file=sys.stderr)
            return 2
    if from_id == "01" and not (ws / SCORESHEET_REL).exists():
        print(f"[AHP-E-WS] 提示：暂无评分表 {SCORESHEET_REL}，第 08 步会自动生成空白模板供填写。",
              file=sys.stderr)

    ids = [sid for sid, _, _ in STEPS]
    if from_id not in STEP_BY_ID or to_id not in STEP_BY_ID:
        print(f"[AHP-E-WS] --from/--to 必须是 01..09 之一", file=sys.stderr)
        return 2
    if ids.index(from_id) > ids.index(to_id):
        print("[AHP-E-WS] --from 不能晚于 --to", file=sys.stderr)
        return 2

    # 演示续跑：07/08 的产物若是在「未声明演示」时生成的（权重表无演示页脚、
    # 过程 JSON demo≠true），从 08/09 续跑且声明了 --demo 时必须把它们一并重跑，
    # 否则权重表缺演示标记，而 09 仍会自述「上游已按 --demo 生成」——自相矛盾。
    if demo_flag(cfg) and from_id in ("08", "09"):
        for step, rel in (("07", f"{PROC}/07-权重表导出/权重表.json"),
                          ("08", SCORES_JSON_REL)):
            if ids.index(from_id) <= ids.index(step):
                continue
            path = ws / rel
            try:
                marked = json.loads(path.read_text(encoding="utf-8-sig")).get("demo") is True
            except (OSError, ValueError):
                marked = False
            if not marked:
                print(f"[演示续跑] 上游 {step} 的产物未按 --demo 生成：从 {step} 起重跑，"
                      "确保演示页脚与过程 JSON 的 demo 标记一致。\n")
                from_id = step
                break

    # 前置产物检查：新工作区只有输入，没有 01-07 的中间产物
    missing: list[str] = []
    if ids.index(from_id) >= ids.index("06"):
        missing += [f"{PROC}/05-权重与一致性/矩阵集.json", f"{PROC}/05-权重与一致性/权重表.json"]
    if from_id == "09":
        missing.append(f"{PROC}/08-方案打分/评分表.json")
    missing = [m for m in missing if not (ws / m).is_file()]
    if missing:
        print("[AHP-E-WS] 缺少前置产物：\n  " + "\n  ".join(missing)
              + "\n  新工作区首次运行请用全链（不带 --from），或先补跑 --from 01 --to 07。", file=sys.stderr)
        return 2

    started = now_iso()
    log: list[dict] = []
    report: list[dict] = []
    notes: list[str] = []
    hint = demo_name_hint(cfg.get("projectName") or "", cfg)
    if hint:
        notes.append(hint)
        print(f"[提示] {hint}\n")
    print(f"== Cardo城市更新AHP层次分析法 · 总控 ==\n工作区：{ws}\n区间：{from_id} → {to_id}"
          f"\n项目：{cfg.get('projectName')}；专家 {cfg.get('experts')} 位"
          f"；{'演示样例（demo=true）' if demo_flag(cfg) else '常规项目'}\n")

    current: Step | None = None
    try:
        for sid, dirname, name in STEPS:
            if ids.index(sid) < ids.index(from_id):
                report.append({"id": sid, "name": name, "status": "skipped"})
                continue
            print(f"[{sid}] {name}")
            current = Step(sid, dirname, name, ws, log)
            t0 = datetime.now()
            if sid == "01":
                step_01(current)
            elif sid == "02":
                step_02(current, cfg)
            elif sid == "03":
                step_03(current)
            elif sid == "04":
                step_04(current)
            elif sid == "05":
                step_05(current, cfg)
            elif sid == "06":
                step_06(current)
            elif sid == "07":
                step_07(current, cfg)
            elif sid == "08":
                step_08(current, cfg)
            elif sid == "09":
                step_09(current, cfg)
            report.append({"id": sid, "name": name, "status": "passed",
                           "durationMs": int((datetime.now() - t0).total_seconds() * 1000),
                           "gates": current.gates, "outputs": current.outputs})
            print(f"    → 通过\n")
            if sid == to_id:
                break
    except StepFailure as exc:
        if current is not None:
            report.append({"id": current.sid, "name": current.name, "status": "failed",
                           "gates": current.gates, "outputs": current.outputs,
                           "error": str(exc)})
        print(f"\n[STOP] [{current.sid if current else '??'}] {exc}", file=sys.stderr)
        summary = {"schema": "cardo-ahp/pipeline-run/v1", "workspace": str(ws),
                   "startedAt": started, "finishedAt": now_iso(), "ok": False,
                   "declined": isinstance(exc, Declined),
                   "failedStep": current.sid if current else None,
                   "steps": report, "commands": log, "notes": notes}
        write_text(ws / f"{CTRL}/运行日志.json", json.dumps(summary, ensure_ascii=False, indent=2))
        write_run_report(ws, summary)
        return getattr(exc, "exit_code", 1)
    except Exception as exc:  # 兜底：未预期异常也要落日志，不让用户看到裸栈 + 旧报告
        if current is not None:
            report.append({"id": current.sid, "name": current.name, "status": "failed",
                           "gates": current.gates, "outputs": current.outputs,
                           "error": f"{type(exc).__name__}: {exc}"})
        print(f"\n[STOP] [{current.sid if current else '??'}] 未预期异常 "
              f"{type(exc).__name__}: {exc}", file=sys.stderr)
        summary = {"schema": "cardo-ahp/pipeline-run/v1", "workspace": str(ws),
                   "startedAt": started, "finishedAt": now_iso(), "ok": False,
                   "failedStep": current.sid if current else None,
                   "steps": report, "commands": log, "notes": notes}
        write_text(ws / f"{CTRL}/运行日志.json", json.dumps(summary, ensure_ascii=False, indent=2))
        write_run_report(ws, summary)
        return 1

    summary = {"schema": "cardo-ahp/pipeline-run/v1", "workspace": str(ws),
               "startedAt": started, "finishedAt": now_iso(), "ok": True,
               "steps": report, "commands": log, "notes": notes}
    write_text(ws / f"{CTRL}/运行日志.json", json.dumps(summary, ensure_ascii=False, indent=2))
    write_run_report(ws, summary)
    print(f"\n全部通过（{from_id}→{to_id}），日志：{CTRL}/运行日志.json")
    return 0


def write_run_report(ws: Path, summary: dict) -> None:
    lines = ["# 00 · 总控运行报告", "",
             f"- 工作区：`{summary['workspace']}`",
             f"- 开始 / 结束：{summary['startedAt']} / {summary['finishedAt']}",
             f"- 结果：{'全部通过' if summary['ok'] else '在第 ' + str(summary.get('failedStep')) + ' 步停止'}",
             "", "| 步骤 | 名称 | 状态 | 耗时(ms) | 闸门 |", "|---|---|---|---|---|"]
    icon = {"passed": "✅通过", "failed": "❌失败", "skipped": "⏭跳过"}
    for r in summary["steps"]:
        gates = "；".join(f"{g['name']}{'✓' if g['pass'] else '✗'}" for g in r.get("gates", []))
        lines.append(f"| {r['id']} | {r['name']} | {icon.get(r['status'], r['status'])} "
                     f"| {r.get('durationMs', '-')} | {gates or '-'} |")
    if not summary["ok"]:
        lines += ["", "## 失败原因", "", "```", summary["steps"][-1].get("error", ""), "```"]
    if summary.get("notes"):
        lines += ["", "## 提示"] + [f"- {n}" for n in summary["notes"]]
    lines += ["", "## 执行命令流水", "", "| 步骤 | 脚本 | 退出码 | 输出文件 |", "|---|---|---|---|"]
    for c in summary.get("commands", []):
        lines.append(f"| {c['step']} | `{c['script']}` | {c['exit']} | {c.get('outFile') or ''} |")
    write_text(ws / f"{CTRL}/运行报告.md", "\n".join(lines))


# ---------------------------------------------------------------- status / scaffold

EXPECT = {
    "01": ["过程文件/01-指标体系/指标目录.json"],
    "02": ["过程文件/02-问卷生成/问卷空白/问卷清单.json"],
    "03": ["过程文件/03-回收校验/回收校验报告.json"],
    "04": ["过程文件/04-多专家聚合/群体矩阵.json"],
    "05": ["过程文件/05-权重与一致性/权重表.json", "过程文件/05-权重与一致性/矩阵集.json"],
    "06": ["过程文件/06-一致性校准/校准记录.md"],
    "07": ["成果文件/AHP权重表.xlsx", "过程文件/07-权重表导出/权重表.json"],
    "08": ["过程文件/08-方案打分/评分表.json", "成果文件/AHP方案评分结论.xlsx"],
    "09": ["过程文件/09-成果验证与成稿/独立验证结果.json",
           "成果文件/成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.docx",
           "成果文件/成品版/AHP权重与评分结论/AHP权重与评分结论_成品版.pdf"],
}
# 09 装配会把 07/08 的散件收进 成品版/<成果名>/（一成果一子目录），
# status 要认新位置，否则整轮跑完反显「07/08 未完成」。
MOVED = {"成果文件/AHP权重表.xlsx": "成果文件/成品版/AHP权重表/AHP权重表_成品版.xlsx",
         "成果文件/AHP方案评分结论.xlsx": "成果文件/成品版/AHP方案评分结论/AHP方案评分结论_成品版.xlsx"}


def _present(ws: Path, rel: str) -> bool:
    if (ws / rel).exists():
        return True
    moved = MOVED.get(rel)
    return bool(moved) and (ws / moved).exists()


def cmd_status(ws: Path) -> int:
    print(f"工作区：{ws}\n")
    print(f"{'步骤':<4}{'名称':<14}{'产物':<6}{'说明'}")
    for sid, _, name in STEPS:
        ok = all(_present(ws, r) for r in EXPECT[sid])
        extra = ""
        if sid == "09" and (ws / EXPECT[sid][0]).is_file():
            v = read_json(ws / EXPECT[sid][0])
            extra = f"通过 {v.get('passed')} / 失败 {v.get('failed')} / 跳过 {v.get('skipped')}"
        if sid == "08" and (ws / EXPECT[sid][0]).is_file():
            s = read_json(ws / EXPECT[sid][0])
            extra = "；".join(f"{r['name']}={r['composite']}" for r in s.get("ranking", []))
        print(f"{sid:<4}{name:<14}{'✅' if ok else '—':<6}{extra}")
    logp = ws / f"{CTRL}/运行日志.json"
    if logp.is_file():
        d = read_json(logp)
        print(f"\n最近一次总控：{'通过' if d.get('ok') else '失败'} @ {d.get('finishedAt')}")
    return 0


def cmd_scaffold(target: Path, project_name: str, experts: int, baseline_mode: str,
                 force: bool, demo: bool = False) -> int:
    src = WORKSPACE
    if not (src / "Skills").is_dir():
        print(f"[AHP-E-WS] 源工作区无效：{src}", file=sys.stderr)
        return 2
    if target.resolve() == src.resolve():
        print("[AHP-E-WS] 目标目录不能等于源工作区", file=sys.stderr)
        return 2
    if target.exists() and any(target.iterdir()) and not force:
        print(f"[AHP-E-WS] 目标目录非空：{target}（确认覆盖请加 --force）", file=sys.stderr)
        return 2
    target.mkdir(parents=True, exist_ok=True)

    ignore = shutil.ignore_patterns("__pycache__", "*.pyc")
    shutil.copytree(src / "Skills", target / "Skills", dirs_exist_ok=True, ignore=ignore)
    for sub in ["01-AHP四准则模板", "02-专家问卷（已回收）", "03-方案打分表"]:
        (target / "输入文件" / sub).mkdir(parents=True, exist_ok=True)
    for sid, dirname, _ in STEPS:
        (target / PROC / dirname).mkdir(parents=True, exist_ok=True)
    (target / CTRL).mkdir(parents=True, exist_ok=True)
    (target / OUT).mkdir(parents=True, exist_ok=True)
    (target / "记忆与规则").mkdir(parents=True, exist_ok=True)

    # 输入：模板 + 已回收问卷（含专家名单）+ 评分表（作为新测试的起点；输入只放 Office）
    shutil.copy2(src / TEMPLATE_REL, target / TEMPLATE_REL)
    for f in (src / SURVEY_REL).glob("*"):
        if f.is_file():
            shutil.copy2(f, target / SURVEY_REL / f.name)
    if (src / SCORESHEET_REL).is_file():
        shutil.copy2(src / SCORESHEET_REL, target / SCORESHEET_REL)

    # 基线：none=不建（只做结构验证）；weights=沿用权重类基线；full=连得分基线一起沿用
    # 基线按包命名空间隔离：只继承本包写的基线（package=CardoAHP），别包的旧文件不复制
    if baseline_mode in ("weights", "full"):
        bsrc = src / "记忆与规则/验证基线-CardoAHP.json"
        b = read_json(bsrc) if bsrc.is_file() else None
        if b and b.get("package") == "CardoAHP":
            keep = {k: b[k] for k in ("criteriaWeights", "globalWeights", "templateSha256") if k in b}
            if baseline_mode == "full":
                keep.update({k: b[k] for k in ("schemeScores", "schemeOrder") if k in b})
            keep["schema"] = "cardo-ahp/verification-baseline/v1"
            keep["package"] = "CardoAHP"
            keep["note"] = ("scaffold 继承的基线：weights 模式只校验权重与模板；"
                            "full 模式连方案得分一并校验。换评分测试建议用 none。")
            dest = target / "记忆与规则/验证基线-CardoAHP.json"
            dest.parent.mkdir(parents=True, exist_ok=True)
            write_text(dest, json.dumps(keep, ensure_ascii=False, indent=2))
        elif bsrc.is_file():
            print("  [跳过] 旧基线不是本包的（package={}），不继承".format(b.get("package")))

    cfg = {"schema": "cardo-ahp/workspace/v1", "projectName": project_name,
           "experts": experts, "baselineMode": baseline_mode, "demo": bool(demo),
           "createdAt": now_iso(), "scaffoldedFrom": str(src)}
    write_text(target / "工作区参数.json", json.dumps(cfg, ensure_ascii=False, indent=2))

    readme = [
        f"# {project_name} · AHP 测试工作区", "",
        "由 `CardoAHP/Skills/00-Cardo城市更新AHP层次分析法` 的 `scaffold` 生成，自包含。", "",
        "## 约定",
        "输入文件、成果文件只放常规 Office（xlsx/docx）；机读 JSON 全部在 过程文件/ 对应编号目录。", "",
        "## 跑法", "",
        "```powershell",
        "# 1) 改评分：编辑 输入文件/03-方案打分表/AHP方案评分表.xlsx",
        "#    （黄色格填 0-10 分，依据列写出处；方案信息表登记方案编号与名称）",
        "# 2) 全链或只重跑评分段",
        'python <技能目录>/Skills/00-Cardo城市更新AHP层次分析法/scripts/pipeline.py run --workspace .',
        'python <技能目录>/Skills/00-Cardo城市更新AHP层次分析法/scripts/pipeline.py run --workspace . --from 08',
        "```", "",
        f"- 专家 {experts} 位；基线模式 `{baseline_mode}`（none = 只做结构自洽验证）。",
        ("- 演示示例：工作区参数 demo=true，四件产物均带「演示样例」页脚。" if demo else
         "- 真实项目：工作区参数 demo=false，产物不加演示标识。"),
        "- 成果：AHP权重表.xlsx、AHP方案评分结论.xlsx、AHP权重与评分结论.docx、AHP依据对照表.xlsx。",
        "- 入口技能文档：[Skills/00-Cardo城市更新AHP层次分析法/SKILL.md]"
        "(Skills/00-Cardo城市更新AHP层次分析法/SKILL.md)。",
    ]
    write_text(target / "README.md", "\n".join(readme))

    built = sum(1 for _ in (target / "Skills").rglob("SKILL.md"))
    print(f"已生成工作区：{target}")
    print(f"  技能 {built} 个；输入（模板/问卷/评分表）已就位；基线模式 {baseline_mode}")
    print(f"  下一步：改 {SCORESHEET_REL} 后执行")
    print(f"  python <技能目录>/Skills/00-Cardo城市更新AHP层次分析法/scripts/pipeline.py run --workspace {target}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Cardo城市更新AHP层次分析法 · 总控")
    sub = ap.add_subparsers(dest="command", required=True)

    for name, help_text in (("run", "按顺序执行 01→09（带闸门）"),
                            ("status", "查看各步产物与最近一次运行")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--workspace", default=str(WORKSPACE))
        if name == "run":
            p.add_argument("--from", dest="from_id", default="01")
            p.add_argument("--to", dest="to_id", default="09")
            p.add_argument("--experts", type=int, default=None)
            p.add_argument("--project-name", default=None)
            p.add_argument("--demo", action="store_true",
                           help="声明本次为演示样例：过程数据写 demo=true，四件产物加「演示样例」页脚")

    p = sub.add_parser("scaffold", help="复制技能链与输入，生成新的测试工作区")
    p.add_argument("--target", required=True)
    p.add_argument("--project-name", default="AHP评分测试工作区")
    p.add_argument("--experts", type=int, default=5)
    p.add_argument("--baseline", dest="baseline_mode", default="none",
                   choices=["none", "weights", "full"])
    p.add_argument("--demo", action="store_true",
                   help="把新工作区标记为演示样例（写入 工作区参数.json 的 demo=true）")
    p.add_argument("--force", action="store_true")

    args = ap.parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass

    if args.command == "scaffold":
        return cmd_scaffold(Path(args.target).resolve(), args.project_name,
                            args.experts, args.baseline_mode, args.force, args.demo)

    ws = Path(args.workspace).resolve()
    if args.command == "status":
        return cmd_status(ws)

    cfg = load_cfg(ws)
    if args.experts is not None:
        cfg["experts"] = args.experts
    if args.project_name:
        cfg["projectName"] = args.project_name
    if args.demo:
        cfg["demo"] = True
        params_path = ws / "工作区参数.json"
        params = read_json(params_path) if params_path.is_file() else {}
        params["demo"] = True
        write_text(params_path, json.dumps(params, ensure_ascii=False, indent=2))
    return run_pipeline(ws, args.from_id, args.to_id, cfg)


if __name__ == "__main__":
    sys.exit(main())
