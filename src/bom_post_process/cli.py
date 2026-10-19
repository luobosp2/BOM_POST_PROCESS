"""Independent BOM commands with a manual review step before export."""

from __future__ import annotations

import argparse
from dataclasses import asdict
import json
from pathlib import Path
import sys

from . import __version__
from .errors import BomError


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="AD23 BOM 检验、参数解析与标准化 Word 导出")
    parser.add_argument("--version", action="version", version=__version__)
    commands = parser.add_subparsers(dest="command", required=True)
    descriptions = {
        "validate": "检查电阻电容封装冲突及同料号拆行，并检查所有元件的位号和数量",
        "gen_value": "保留原表和样式，追加 name、note、not_bought、not_welded 四列",
        "gen_soldering": "从人工修改后的参数表导出焊接表，排除 not_welded 非空行",
        "gen_bom": "从人工修改后的参数表导出采购 BOM，排除 not_bought 非空行",
        "gen_bom_standard": "从采购 BOM Excel 生成标准化 Word、PDF 和分页核验记录",
    }
    for name, description in descriptions.items():
        command = commands.add_parser(name, help=description, description=description)
        command.add_argument("input", type=Path, help="输入 .xlsx 文件")
        command.add_argument("--sheet", help="输入工作表名；默认自动识别")
        command.add_argument("-o", "--output", type=Path,
                             help="输出路径；validate 为可选 JSON 报告，其他命令默认写入 output/")
        if name == "gen_bom_standard":
            command.add_argument("--template", type=Path, help="自定义两页 Word 模板")
            command.add_argument("--config", type=Path, help="表尾文字 JSON 配置")
    return parser


def _available(paths: list[Path], source: Path) -> None:
    for path in paths:
        if path.resolve() == source.resolve():
            raise BomError("输出路径不能与输入文件相同")
        if path.exists():
            raise BomError(f"输出文件已存在，请指定新路径：{path}")


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(data, ensure_ascii=False, indent=2, default=str) + "\n"
    with path.open("x", encoding="utf-8") as stream:
        try:
            stream.write(content)
        except OSError:
            stream.close()
            path.unlink(missing_ok=True)
            raise


def _show_validation(report) -> None:
    print(f"工作表：{report.sheet}；完成封装比较 {report.checked} 行，未比较 {report.skipped} 行；"
          f"错误 {report.errors}，提示 {report.warnings}。")
    errors = [issue for issue in report.issues if issue.severity == "error"]
    warnings = [issue for issue in report.issues if issue.severity != "error"]
    merge_warnings = [issue for issue in warnings if issue.code == "same_comment_different_footprints"]
    other_warnings = [issue for issue in warnings if issue.code != "same_comment_different_footprints"]
    for issue in errors + merge_warnings + other_warnings[:8]:
        label = "错误" if issue.severity == "error" else "提示"
        print(f"{label}：第 {issue.row} 行 [{issue.designator}] {issue.message}", file=sys.stderr)
    if len(other_warnings) > 8:
        print(f"另有 {len(other_warnings) - 8} 条提示；可通过 validate -o 把完整结果保存为 JSON。", file=sys.stderr)


def _validation_json(report) -> dict:
    result = asdict(report)
    result.update(ok=report.ok, errors=report.errors, warnings=report.warnings,
                  checked=report.checked, skipped=report.skipped)
    return result


def _show_values(report) -> None:
    print(f"参数解析：电容 {report.capacitors} 行，电阻 {report.resistors} 行；"
          f"共 {report.total} 条物料。")
    for issue in report.issues:
        print(f"提示：第 {issue.row} 行，{issue.model!r}：{issue.reason}", file=sys.stderr)
    print(f"待人工调整的参数表：{report.output}")
    print("name 已预填可识别元件名称；新建 not_welded 列时，NC 条目会自动标记为 NC。")
    print("请编辑 name、note；不采购的条目填写 not_bought，不焊接的条目填写 not_welded。")
    print("保存后分别运行 gen_soldering 和 gen_bom；再对 gen_bom 的结果运行 gen_bom_standard。")


def _show_standard(report) -> None:
    print(f"标准化 BOM：{report.items} 条物料，{report.pages} 页。")
    print(f"Word：{report.output}\nPDF：{report.pdf}\n分页核验：{report.audit}")


def _run(args: argparse.Namespace) -> int:
    if args.command == "validate":
        from .validation import validate_workbook

        if args.output:
            if args.output.suffix.lower() != ".json":
                raise BomError("校验报告输出必须为 .json")
            _available([args.output], args.input)
        report = validate_workbook(args.input, sheet=args.sheet)
        _show_validation(report)
        if args.output:
            _write_json(args.output, _validation_json(report))
            print(f"校验报告：{args.output.resolve()}")
        return 0 if report.ok else 1

    if args.command == "gen_value":
        from .values import gen_value

        output = args.output or Path("output") / f"{args.input.stem}_已解析.xlsx"
        _show_values(gen_value(args.input, output, sheet=args.sheet))
        return 0

    if args.command == "gen_soldering":
        from .soldering import gen_soldering

        stem = args.input.stem.removesuffix("_已解析")
        output = args.output or Path("output") / f"{stem}_焊接表.xlsx"
        report = gen_soldering(args.input, output, sheet=args.sheet)
        print(f"焊接表 {report.written} 行；按 not_welded 排除 {report.excluded} 行。")
        print(f"焊接表：{report.output}")
        return 0

    if args.command == "gen_bom":
        from .bom import gen_bom

        stem = args.input.stem.removesuffix("_已解析")
        output = args.output or Path("output") / f"{stem}_BOM.xlsx"
        report = gen_bom(args.input, output, sheet=args.sheet)
        print(f"采购 BOM {report.written} 行；按 not_bought 排除 {report.excluded} 行。")
        if report.pcb_placeholder:
            print("已添加 PCB 制板占位行，数量留空，可在 BOM Excel 中补充。")
        print(f"采购 BOM：{report.output}")
        return 0

    if args.command == "gen_bom_standard":
        from .standard import export_standard_bom

        stem = args.input.stem.removesuffix("_BOM")
        output = args.output or Path("output") / f"{stem}_标准化.docx"
        _show_standard(export_standard_bom(args.input, output, sheet=args.sheet,
                                          template=args.template, config=args.config))
        return 0

def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        return _run(args)
    except (BomError, OSError, ValueError, RuntimeError) as error:
        print(f"错误：{error}", file=sys.stderr)
        return 1
