import argparse
import json
import sys
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from .audit import run
from .document import format_docx


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="format-research-word", description="中文科研 Word 报告排版工具"
    )
    commands = parser.add_subparsers(dest="command", required=True)
    audit = commands.add_parser("audit", help="查看文献映射与长小数候选")
    audit.add_argument("input", type=Path)
    audit.add_argument("--json", type=Path, help="导出完整结果到新的 JSON 文件")
    fmt = commands.add_parser("format", help="将可编辑正文、题注及表格排版到新 DOCX")
    fmt.add_argument("input", type=Path)
    fmt.add_argument("output", type=Path)
    fmt.add_argument(
        "--renumber-citations",
        action="store_true",
        help="按事先复核的静态方括号引文重排文献",
    )
    args = parser.parse_args(argv)
    try:
        if args.command == "format":
            stats = format_docx(args.input, args.output, args.renumber_citations)
            print(args.output, json.dumps(stats, ensure_ascii=False))
        else:
            report = run(args.input)
            if args.json:
                with args.json.open("x", encoding="utf-8") as f:
                    json.dump(report, f, ensure_ascii=False, indent=2)
            print(
                "文献条目 {文献条目}，候选引文 {静态候选引用处} 处，涉及 {已引用条目} 条。".format(
                    **report
                )
            )
            print("未见静态引用的原号", report["未见静态引用的原号"])
            print("长小数候选", len(report["长小数候选"]))
            for hit in report["长小数候选"][:5]:
                print(hit["原值"], "→", hit["建议显示"])
            for msg in report["范围提示"]:
                print(msg)
        return 0
    except (OSError, ValueError, KeyError, BadZipFile, SyntaxError, ParseError) as e:
        print(f"处理失败：{e}", file=sys.stderr)
        return 1
