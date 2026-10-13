"""Export the reviewed BOM verbatim, separating mounted and NC components."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from .components import component_name, is_nc
from .errors import BomError
from .workbook import (
    BOM_HEADERS, cell_text, copy_table_layout, copy_table_row, open_workbook, prepare_output,
    read_ad_table, save_workbook_exclusive, select_bom_sheet, select_note_sheet, set_text, split_designators,
)


@dataclass(frozen=True)
class SolderingReport:
    output: Path
    written: int
    excluded_nc: int
    excluded_pcb: int
    nc_designators: int
    source_sheet: str
    bom_sheet: str


def _map_reviewed_notes(workbook, source, headers):
    """Build the six export columns in memory; never calculate or replace notes."""
    if source.merged_cells:
        raise BomError("待导出 BOM 包含合并单元格，请取消合并后再导出焊接表")
    rows = read_ad_table(workbook, source.title).rows
    mapped = workbook.create_sheet("_焊接导出映射")
    columns = [headers[name] for name in ("comment", "comment", "libref", "quantity", "designator", "note")]
    copy_table_layout(source, mapped)
    copy_table_row(source, mapped, 1, 1, columns)
    for index, label in enumerate(BOM_HEADERS, start=1):
        set_text(mapped.cell(1, index), label)
    for output_column, input_column in enumerate(columns, start=1):
        width = source.sheet_format.defaultColWidth or 13
        for dimension in source.column_dimensions.values():
            if (dimension.min or input_column) <= input_column <= (dimension.max or input_column):
                width = dimension.width
                break
        mapped.column_dimensions[get_column_letter(output_column)].width = width
    for index, row in enumerate(rows, start=1):
        note = source.cell(row.row, headers["note"])
        if note.data_type in {"f", "e"}:
            raise BomError(f"note 单元格 {note.coordinate} 包含公式或 Excel 错误，请先修正")
        copy_table_row(source, mapped, row.row, index + 1, columns)
        mapped.cell(index + 1, 1, index)
        set_text(mapped.cell(index + 1, 3), component_name(row.comment, row.libref))
    return mapped


def gen_soldering(
    input_path: Path | str, output_path: Path | str | None = None, sheet: str | None = None,
    *, ad_sheet: str | None = None,
) -> SolderingReport:
    """Copy reviewed values; optionally select which original AD sheet defines NC."""
    source = Path(input_path).expanduser().resolve()
    default = source.with_name(source.stem.removesuffix("_已解析") + "_焊接表.xlsx")
    output = prepare_output(source, output_path or default, ".xlsx")
    workbook = open_workbook(source)
    result = Workbook()
    mounted = result.active
    mounted.title = "焊接表"
    unmounted = result.create_sheet("未焊接元件")
    excluded_pcb = 0
    try:
        reviewed = select_note_sheet(workbook, sheet)
        if reviewed is not None:
            note_sheet, headers = reviewed
            original = read_ad_table(workbook, ad_sheet or note_sheet.title)
            bom_sheet = note_sheet.title
            bom = _map_reviewed_notes(workbook, note_sheet, headers)
        else:
            bom = select_bom_sheet(workbook, sheet)
            if bom is None:
                raise BomError("缺少 note 列；请先运行 gen_value 并人工修改备注后再导出焊接表")
            original = read_ad_table(workbook, ad_sheet)
            bom_sheet = bom.title
        if not original.rows:
            raise BomError("AD 原始表没有物料数据")
        if bom.merged_cells:
            raise BomError("整理后 BOM 包含合并单元格，请取消合并后再处理")
        nc_refs: set[str] = set()
        for row in original.rows:
            if is_nc(row.comment):
                references = split_designators(row.designator)
                if not references:
                    raise BomError(f"AD 原始表第 {row.row} 行标记为 NC，但 Designator 位号为空")
                nc_refs.update(references)
        result.loaded_theme = workbook.loaded_theme
        result.epoch = workbook.epoch
        for target in (mounted, unmounted):
            copy_table_layout(bom, target)
            copy_table_row(bom, target, 1, 1)
        total = 0
        for cells in bom.iter_rows(min_row=2, max_col=6):
            if all(cell.value is None or cell.value == "" for cell in cells):
                continue
            total += 1
            for cell in cells:
                if cell.data_type in {"f", "e"}:
                    raise BomError(f"整理后 BOM 的 {cell.coordinate} 包含公式或 Excel 错误，请先修正")
            if any(cell_text(cell.value).strip() == "PCB制板" for cell in cells[1:3]):
                excluded_pcb += 1
                continue
            references = split_designators(cells[4].value)
            matched = references & nc_refs
            if matched and matched != references:
                raise BomError(f"整理后 BOM 第 {cells[0].row} 行混有贴装和不贴装位号；不贴装位号："
                               + ", ".join(sorted(matched)) + "。请拆分该行后重试")
            target = unmounted if matched else mounted
            copy_table_row(bom, target, cells[0].row, target.max_row + 1)
        if not total:
            raise BomError("整理后 BOM 没有物料数据")
        for target in (mounted, unmounted):
            target.auto_filter.ref = f"A1:F{target.max_row}"
            target.print_title_rows = "1:1"
            target.print_area = f"A1:F{target.max_row}"
        save_workbook_exclusive(result, output)
        return SolderingReport(output, mounted.max_row - 1, unmounted.max_row - 1,
                               excluded_pcb, len(nc_refs), original.sheet, bom_sheet)
    finally:
        workbook.close()
        result.close()
