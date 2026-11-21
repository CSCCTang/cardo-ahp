#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_templates.py -- 空工作区的「发料」一步：一次性铺好人要填的所有空白表。

**为什么有这一步**：默认场景就是空白文件夹 = 全新项目。用户手上什么都没有，
系统该做的是：把要人填的表**全部铺出来**，再把「要你定的事」列清楚，
等用户填完放回来继续算 —— 而不是先卡在「缺参数」上让人自己猜。

产出（都在输入侧，xlsx）：
  · 输入文件/02-专家问卷（待发放）/AHP专家问卷-专家N.xlsx   专家填
  · 输入文件/03-方案打分表/方案清单.xlsx                    用户填（定方案）
  · 输入文件/03-方案打分表/AHP方案评分表.xlsx               用户填（打分）
评分表此时**没有权重**（权重要等专家问卷回收后第 05 步才算出来），
综合权重列留空待回填，得分与依据列先填也不影响。

用法：
  python make_templates.py --workspace <工作区> [--experts 5] [--project-name 名称]
退出码：0 成功；2 参数/环境错。
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

from scheme_lib import read_scheme_list          # noqa: E402

PENDING_REL = "输入文件/02-专家问卷（待发放）"
RECV_REL = "输入文件/02-专家问卷（已回收）"
SCORE_DIR_REL = "输入文件/03-方案打分表"
SCHEME_LIST_REL = SCORE_DIR_REL + "/方案清单.xlsx"
SCORESHEET_REL = SCORE_DIR_REL + "/AHP方案评分表.xlsx"


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def run(cmd: list[str], cwd: Path) -> tuple[int, str]:
    p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def main() -> int:
    ap = argparse.ArgumentParser(description="为空白工作区铺好全部待填空白表")
    ap.add_argument("--workspace", required=True, help="工作区根目录")
    ap.add_argument("--experts", type=int, default=5)
    ap.add_argument("--project-name", default="",
                    help="预填的项目名称；留空则由用户在《方案清单》B1 格自己填")
    ap.add_argument("--preset", action="store_true",
                    help="预填建议候选方案名（用户改名用）；不给则全空白")
    args = ap.parse_args()

    ws = Path(args.workspace).resolve()
    if not ws.is_dir():
        print(f"[AHP-E-WS] 工作区不存在：{ws}", file=sys.stderr)
        return 2

    for rel in (PENDING_REL, RECV_REL, SCORE_DIR_REL):
        (ws / rel).mkdir(parents=True, exist_ok=True)

    # ---- 1. 方案清单（项目名 + 方案，都由用户填） ----
    list_path = ws / SCHEME_LIST_REL
    if not list_path.is_file():
        cmd = [sys.executable, str(HERE / "gen_scheme_list.py"),
               "--out", str(list_path), "--project-name", args.project_name]
        if args.preset:
            cmd.append("--preset")
        code, out = run(cmd, ws)
        if code != 0:
            print(out.strip()[-300:], file=sys.stderr)
            return 2
        mode = "预填建议方案（改名即可）" if args.preset else "全空白（自己填）"
        print(f"  [表 1/3] 方案清单（{mode}）：{SCHEME_LIST_REL}")
    else:
        print(f"  [表 1/3] 方案清单：已在位（{SCHEME_LIST_REL}）")

    # ---- 2. 方案评分表 ----
    # **不预生成**：评分表的列数 = 方案数，而方案数要等用户填《方案清单》才知道。
    # 首跑就按 3 列铺出来的话，用户填 4 个方案（或 2 个）就与表里的列数对不上，
    # 已填的得分会作废。所以评分表留到第 08 步 —— 清单填好后按实际方案数生成。
    # 用户想提前看表长什么样，看《怎么填.md》里的版式说明即可。
    sheet_path = ws / SCORESHEET_REL
    if sheet_path.is_file():
        print(f"  [表 2/3] 方案评分表：已在位（{SCORESHEET_REL}）")
    else:
        print(f"  [表 2/3] 方案评分表：等你填好方案清单后生成（列数按方案数定）")

    # ---- 3. 专家空白问卷（领料区） ----
    pending = sorted((ws / PENDING_REL).glob("AHP专家问卷-专家*.xlsx"))
    if len(pending) >= args.experts:
        print(f"  [表 3/3] 专家空白问卷：已在位（{len(pending)} 份，{PENDING_REL}）")
    else:
        code, out = run([sys.executable, str(HERE.parent.parent / "02-问卷生成" / "scripts" /
                                             "gen_questionnaire.py"),
                         "generate", "--out", str(ws / "过程文件" / "02-问卷生成" / "问卷空白"),
                         "--experts", str(args.experts),
                         "--scale", "scale_0_1", "--criteria-scale", "scale_1_9",
                         "--project-name", "某某片区改造项目",
                         "--no-timestamp"], ws)
        if code != 0:
            print(out.strip()[-300:], file=sys.stderr)
            return 2
        src_dir = ws / "过程文件" / "02-问卷生成" / "问卷空白"
        import shutil
        n = 0
        for f in sorted(src_dir.glob("AHP专家问卷-专家*.xlsx")):
            tgt = ws / PENDING_REL / f.name
            if not tgt.is_file():
                shutil.copy2(f, tgt)
                n += 1
        print(f"  [表 3/3] 专家空白问卷：{PENDING_REL}（本次复制 {n} 份）")

    # ---- 汇总：现在要人做什么 ----
    listed = read_scheme_list(list_path)
    from scheme_lib import read_project_name
    pname = read_project_name(list_path)
    print("")
    print("  ================= 现在轮到你（做完重跑开工.py，我接着算）=================")
    print("")
    print("  ① 定方案 —— 你来做")
    print(f"     打开：{SCHEME_LIST_REL}")
    if pname:
        print(f"     · B1 项目名已填：{pname}")
    else:
        print("     · B1 格填项目名称（会印在三份成果文件上）")
    if listed:
        print(f"     · 方案名已填 {len(listed)} 个：" + "、".join(s["name"] for s in listed))
    else:
        print("     · 「方案清单」页填方案名与说明"
              "（可翻「候选方案参考」页找思路；行数不限，一般 2~4 个）")
    print("     做完长什么样：B1 有项目名 + 方案名称列有 2~4 行有字")
    print("")
    print("  ② 发问卷 —— 你转给专家")
    print(f"     {PENDING_REL} 的 {args.experts} 份，一份一人，别改名别合并")
    print("     专家填完原样放回：" + RECV_REL)
    print(f"     做完长什么样：回收区里 ≥ {5} 份问卷")
    print("")
    print("  ③ 回收问卷 —— 专家交回后，跑：python <技能目录>/scripts/cardo.py --to 07")
    print("")
    print("  ④ 打分 —— 你来填")
    print(f"     评分表在方案定好、权重算完后自动生成：{SCORESHEET_REL}")
    print("     · 照「参考分档」列给分（1/3/5/7/9 各代表什么，都写好了）")
    print("     · 黄色格填 0–10 分，依据列写这一分的出处")
    print(f"     做完长什么样：所有方案的每个指标都填了分")
    print("")
    print("  ⑤ 出成果 —— 系统算：python <技能目录>/scripts/cardo.py --from 08")
    print("     产出三份：权重表 / 方案评分结论 / 权重与评分结论.docx")
    print("")
    print("  详细步骤看工作区根目录的 放料清单.md（每次运行都刷新状态）")
    print("  ===========================================================================")
    print("")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
