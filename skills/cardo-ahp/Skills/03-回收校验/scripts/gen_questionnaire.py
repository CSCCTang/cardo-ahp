#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_questionnaire.py -- cardo-ahp 专家问卷生成与回收校验（xlsx，openpyxl）。

generate：按标度生成 5-7 位专家各一份 xlsx（灰色输入区上三角 + 互反公式 + 自动校验页）。
check   ：回收校验（缺项/越界/非互反/CR 预估，逐格定位）。

用法（仓库根）：
  python .../gen_questionnaire.py generate --out "<输出目录>" --experts 5 --scale scale_0_1 \
      [--criteria-scale scale_1_9] [--project-name "某某片区改造项目"]
  python .../gen_questionnaire.py check --dir "<问卷目录>" [--no-timestamp]
  python .../gen_questionnaire.py check --file "<问卷.xlsx>" [--scale ...] [--criteria-scale ...]

退出码：0 通过；1 结构问题（缺项/越界/非互反）；2 参数/文件错；3 结构通过但 CR≥0.1（需人工确认）。
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402
import _survey_lib as survey  # noqa: E402

TZ = timezone(timedelta(hours=8))
PRODUCER = "@local/cardo-ahp"
ENGINE_VERSION = "0.1.1"
SURVEY_SCHEMA = "cardo/ahp-survey/v1"
CHECK_SCHEMA = "cardo/ahp-survey-check/v1"
SCALES = ("scale_1_9", "scale_0_1", "ratio")


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def now_iso() -> str:
    return datetime.now(TZ).isoformat(timespec="seconds")


def envelope(args, schema: str) -> dict:
    payload = {"schema": schema, "schemaVersion": 1, "producer": PRODUCER,
               "source": [], "gaps": []}
    if not getattr(args, "no_timestamp", False):
        payload["generatedAt"] = now_iso()
    payload["ruleVersion"] = int((lib.load_scale_rules().get("version") or 0))
    payload["engineVersion"] = ENGINE_VERSION
    return payload


# ---------------------------------------------------------------- generate
def cmd_generate(args) -> int:
    if not (5 <= int(args.experts) <= 7):
        raise lib.AhpError("AHP-E-INPUT", f"--experts 必须是 5-7（实际 {args.experts}）")
    for scale_id in (args.scale, args.criteria_scale):
        if scale_id not in SCALES:
            raise lib.AhpError("AHP-E-SCALE", f"未知标度 {scale_id}（可选 {SCALES}）")
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    project_name = args.project_name or "某某片区改造项目"
    manifest = envelope(args, SURVEY_SCHEMA)
    manifest["projectName"] = project_name
    manifest["experts"] = []
    for index in range(1, int(args.experts) + 1):
        expert_id = str(index)
        path = out_dir / f"AHP专家问卷-专家{expert_id}.xlsx"
        info = survey.write_questionnaire(path, expert_id, args.scale, args.criteria_scale, project_name)
        manifest["experts"].append({
            "id": expert_id,
            "file": path.name,
            "indicatorScale": args.scale,
            "criteriaScale": args.criteria_scale,
            "sheets": ["封面", "准则层", "指标层", "自动校验"],
            "inputCells": sum(len(m["inputs"]) for m in info["matrices"]),
            "matrixCount": len(info["matrices"]),
            "role": "请填写（项目负责人/投资专家/规划专家/财务专家）",
        })
        print(f"生成 {path.name}（输入格 {manifest['experts'][-1]['inputCells']}，矩阵 {len(info['matrices'])}）")
    manifest_path = out_dir / "问卷清单.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"out": str(out_dir), "experts": len(manifest["experts"]),
                      "manifest": manifest_path.name, "scale": args.scale,
                      "criteriaScale": args.criteria_scale}, ensure_ascii=False, indent=2))
    return 0


# ---------------------------------------------------------------- check
def check_one(path: Path, indicator_scale: str, criteria_scale: str) -> dict:
    result = {"file": path.name, "problems": [], "matrices": {}, "crWarnings": []}
    try:
        parsed = survey.read_questionnaire(path, indicator_scale, criteria_scale)
    except Exception as exc:  # 坏 xlsx / 非 xlsx
        raise lib.AhpError("AHP-E-INPUT", f"问卷不可读：{path.name}（{exc}）") from exc
    result["problems"] = parsed["problems"]
    result["warnings"] = parsed.get("warnings", [])
    for mid, item in parsed["matrices"].items():
        evaluation = lib.evaluate(item["matrix"])
        result["matrices"][mid] = {
            "scale": item["scale"],
            "lambdaMax": evaluation["lambdaMax"],
            "ci": evaluation["consistency"]["ci"],
            "ri": evaluation["consistency"]["ri"],
            "cr": evaluation["consistency"]["cr"],
            "pass": evaluation["consistency"]["pass"],
        }
        if not evaluation["consistency"]["pass"]:
            result["crWarnings"].append({"matrix": mid, "cr": evaluation["consistency"]["cr"]})
    return result


def load_manifest(directory: Path) -> dict | None:
    path = directory / "问卷清单.json"
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None


def cmd_check(args) -> int:
    targets: list[Path] = []
    if args.file:
        targets = [Path(args.file)]
    else:
        directory = Path(args.dir)
        if not directory.is_dir():
            raise lib.AhpError("AHP-E-INPUT", f"目录不存在：{args.dir}")
        targets = sorted(directory.glob("AHP专家问卷-专家*.xlsx"))
        if not targets:
            raise lib.AhpError("AHP-E-INPUT", f"目录下没有问卷文件：{args.dir}")
    for scale_id in (args.scale, args.criteria_scale):
        if scale_id not in SCALES:
            raise lib.AhpError("AHP-E-SCALE", f"未知标度 {scale_id}（可选 {SCALES}）")
    results = []
    total_problems = 0
    total_cr = 0
    total_warnings = 0
    for target in targets:
        if not target.is_file():
            raise lib.AhpError("AHP-E-INPUT", f"问卷不存在：{target}")
        indicator_scale = args.scale
        criteria_scale = args.criteria_scale
        manifest = load_manifest(target.parent)
        if manifest:
            entry = next((e for e in manifest.get("experts", []) if e.get("file") == target.name), None)
            if entry:
                indicator_scale = entry.get("indicatorScale", indicator_scale)
                criteria_scale = entry.get("criteriaScale", criteria_scale)
        one = check_one(target, indicator_scale, criteria_scale)
        total_problems += len(one["problems"])
        total_cr += len(one["crWarnings"])
        total_warnings += len(one["warnings"])
        results.append(one)

    payload = envelope(args, CHECK_SCHEMA)
    payload["files"] = results
    payload["summary"] = {
        "files": len(results),
        "problems": total_problems,
        "crWarnings": total_cr,
        "warnings": total_warnings,
        "pass": total_problems == 0 and total_cr == 0,
    }
    if args.format == "md":
        print(render_check_md(payload))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    if total_problems:
        return 1
    if total_cr:
        return 3
    return 0


def render_check_md(payload: dict) -> str:
    """渲染回收校验报告（人读）。

    除了诊断数据（哪位专家、哪个矩阵、CR 多少），还要给**行动指引** ——
    专家拿到报告要能照着改，而不是只知道自己"错了"。
    """
    summary = payload["summary"]
    lines = ["# AHP 问卷回收校验", "",
             f"- 文件 {summary['files']} 份；结构问题 {summary['problems']}；"
             f"CR 未通过 {summary['crWarnings']}", ""]

    # ---- 顶部行动指引：把「哪几份、哪几个矩阵要改」一次说清 ----
    bad_items = []
    for one in payload["files"]:
        bad_matrices = [mid for mid, m in one["matrices"].items() if not m["pass"]]
        if bad_matrices or one["problems"] or one.get("warnings"):
            bad_items.append((one["file"], bad_matrices, one["problems"], one.get("warnings")))
    if bad_items:
        lines += ["## 一、要改什么（先看这段）", ""]
        for name, bads, probs, warns in bad_items:
            if bads:
                lines.append(f"- **{name}**：以下矩阵 CR 超 0.1，需调整 —— " + "、".join(bads))
            if probs:
                lines.append(f"  - 结构问题 {len(probs)} 处（见下表）")
            if warns:
                lines.append(f"  - 警告 {len(warns)} 处（见下）")
        lines += ["", "### 怎么改（照这个做）", "",
                  "1. **只改标度数字，别动别的**：改「准则层」「指标层」页里灰色上三角的格子。",
                  "2. **找最不一致的那几格**：同一行里两个数差距过大时，把它往中间值靠",
                  "   （比如你填了 9 但同组其它比较都在 1~3 之间，就把它降到 3 或 5）。",
                  "3. **别凑数**：一致性是判断逻辑自洽与否的检验，不是让你填得好看。",
                  "   调整后如果 CR 还是超，请重新想想哪两个指标的相对关系填错了。",
                  "4. **改完把这份 xlsx 原样交回**（文件名别改），系统会重算 CR。",
                  "5. **不要改**：对角线（固定是 1）、下三角（自动取倒数的公式）、"
                  "「指标ID」「自动校验」页。",
                  "", "### 标度对照", "",
                  "| 填这个数 | 表示 |", "|---|---|",
                  "| 1 | 两项同等重要 |",
                  "| 3 | 一项比另一项稍重要 |",
                  "| 5 | 一项比另一项明显重要 |",
                  "| 7 | 一项比另一项强烈重要 |",
                  "| 9 | 一项比另一项极端重要 |",
                  "| 2 / 4 / 6 / 8 | 介于相邻两档之间 |",
                  "| 1/3、1/5、1/7、1/9 | 反过来：右边的项更重要（准则层） |",
                  "| 0.1~10 的小数 | 指标层用：填「本项是右项的几倍」，如 3 = 重要 3 倍 |",
                  "", "---", ""]

    lines += ["## 二、逐份诊断明细", ""]
    for one in payload["files"]:
        lines.append(f"### {one['file']}")
        if one["problems"]:
            lines += ["", "| 问题码 | 矩阵 | 单元格 | 说明 |", "|---|---|---|---|"]
            for p in one["problems"]:
                lines.append(f"| {p['code']} | {p.get('matrix', '')} | {p.get('cell', '')} | {p.get('what', '')} |")
        if one.get("warnings"):
            lines += ["", "> 警告：" + "；".join(f"{w['code']} {w.get('matrix', '')}!{w.get('cell', '')}"
                                                for w in one["warnings"])]
        lines += ["", "| 矩阵 | λmax | CI | RI | CR | 判定 |", "|---|---|---|---|---|---|"]
        for mid, m in one["matrices"].items():
            lines.append(f"| {mid} | {m['lambdaMax']:.6f} | {m['ci']:.6f} | {m['ri']:.2f} | "
                         f"{m['cr']:.6f} | {'通过' if m['pass'] else '需调整'} |")
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------- main
def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 专家问卷生成/回收校验")
    sub = ap.add_subparsers(dest="command", required=True)

    p_gen = sub.add_parser("generate", help="生成 5-7 位专家问卷")
    p_gen.add_argument("--out", required=True)
    p_gen.add_argument("--experts", type=int, default=5)
    p_gen.add_argument("--scale", default="scale_0_1")
    p_gen.add_argument("--criteria-scale", default="scale_1_9")
    p_gen.add_argument("--project-name", default=None)
    p_gen.add_argument("--no-timestamp", action="store_true")
    p_gen.set_defaults(func=cmd_generate)

    p_chk = sub.add_parser("check", help="回收校验")
    group = p_chk.add_mutually_exclusive_group(required=True)
    group.add_argument("--file", default=None)
    group.add_argument("--dir", default=None)
    p_chk.add_argument("--scale", default="scale_0_1")
    p_chk.add_argument("--criteria-scale", default="scale_1_9")
    p_chk.add_argument("--format", choices=["json", "md"], default="json")
    p_chk.add_argument("--no-timestamp", action="store_true")
    p_chk.set_defaults(func=cmd_check)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except lib.AhpError as exc:
        print(json.dumps({"error": {"code": exc.code, "message": exc.message}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return exc.exit_code
    except (RuntimeError, OSError) as exc:
        print(json.dumps({"error": {"code": "AHP-E-INTERNAL", "message": str(exc)}}, ensure_ascii=False, indent=2),
              file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())