#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ahp_export.py -- cardo-ahp 权重表导出/校验 CLI（算法在 _export_lib.py）。

用法（仓库根）：
  python .../ahp_export.py export --weights "<权重.json>" --out "<权重表.xlsx>" [--out-json "<权重表.json>"]
      [--scores "<打分.json>"] [--template "<模板.xlsx>"] [--direction <方向ID>] [--no-timestamp]
  python .../ahp_export.py verify --file "<权重表.xlsx>" --json "<权重表.json>" [--template "<模板.xlsx>"]
  python .../ahp_export.py consistency --workbook "<成品权重表.xlsx>" --json "<权重表.json>" --matrix-set "<矩阵集.json>"
退出码：0 成功/一致；1 校验不一致；2 输入/参数错；3 表内 CR 与报告不一致（需人工核）。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True

import _ahp_lib as lib  # noqa: E402
import _export_lib as ex  # noqa: E402


def cmd_export(args) -> int:
    return ex.export_weights(args)


def cmd_verify(args) -> int:
    return ex.verify_weights(args)


def cmd_consistency(args) -> int:
    code, result = ex.check_consistency(Path(args.workbook), Path(args.json),
                                        Path(args.matrix_set))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return code


def main(argv: list[str] | None = None) -> int:
    ex.setup_stdout()
    ap = argparse.ArgumentParser(description="cardo-ahp 权重表导出/回读校验")
    sub = ap.add_subparsers(dest="command", required=True)
    p = sub.add_parser("export", help="导出 xlsx + 权重 JSON")
    p.add_argument("--weights", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--out-json", default=None)
    p.add_argument("--scores", default=None)
    p.add_argument("--template", default=str(ex.ROOT_TEMPLATE))
    p.add_argument("--direction", default=None)
    p.add_argument("--demo", action="store_true",
                   help="演示样例：权重表页脚标注「演示样例」，过程数据写 demo=true（不改数值）")
    p.add_argument("--no-timestamp", action="store_true")
    p.add_argument("--matrix-set", default=None,
                   help="05 的矩阵集.json（full 矩阵）；给了就按完整矩阵写入，导出侧与计算侧同口径")
    p.add_argument("--out-ann", default=None, help="溯源底稿 xlsx（模板副本，供回读对账）")
    p.add_argument("--project", default=None, help="封面用项目名称")
    p.add_argument("--org", default=None, help="封面用编制单位")
    p.add_argument("--version", default=None, help="封面用版本号")
    p.set_defaults(func=cmd_export)
    v = sub.add_parser("verify", help="回读 xlsx 与 JSON 对账")
    v.add_argument("--file", required=True)
    v.add_argument("--json", required=True)
    v.add_argument("--template", default=str(ex.ROOT_TEMPLATE))
    v.set_defaults(func=cmd_verify)
    c = sub.add_parser("consistency", help="成品表内 CR 与报告 JSON、矩阵集复算逐项比对")
    c.add_argument("--workbook", required=True)
    c.add_argument("--json", required=True)
    c.add_argument("--matrix-set", required=True)
    c.set_defaults(func=cmd_consistency)
    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except lib.AhpError as exc:
        print(json.dumps({"error": {"code": exc.code, "message": exc.message}},
                         ensure_ascii=False, indent=2), file=sys.stderr)
        return exc.exit_code
    except Exception as exc:  # 兜底：未预期异常也不抛 Traceback
        print(json.dumps({"error": {"code": "AHP-E-INTERNAL",
                                    "message": f"{type(exc).__name__}: {exc}"}},
                         ensure_ascii=False, indent=2), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
