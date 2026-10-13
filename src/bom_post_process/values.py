"""Append a note column to the original AD table without restyling it."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl.cell.cell import MergedCell

from .components import component_kind, parse_capacitor, parse_resistor
from .errors import BomError
from .workbook import (
    open_workbook, prepare_output, read_ad_table, save_bytes_exclusive, select_ad_sheet,
)
from .xlsx_notes import update_notes


@dataclass(frozen=True)
class ParseIssue:
    row: int
    model: str
    reason: str
    sheet: str = ""


@dataclass(frozen=True)
class ValueReport:
    output: Path
    source_sheet: str
    sheet: str
    capacitors: int
    resistors: int
    skipped: int
    issues: tuple[ParseIssue, ...]
    total: int

    @property
    def parsed(self) -> int:
        return self.capacitors + self.resistors


def gen_value(
    input_path: Path | str, output_path: Path | str | None = None, sheet: str | None = None,
) -> ValueReport:
    """Copy the workbook and fill empty note cells for supported components."""
    source = Path(input_path).expanduser().resolve()
    output = prepare_output(source, output_path or source.with_name(source.stem + "_已解析.xlsx"), ".xlsx")
    workbook = open_workbook(source)
    counts = {"capacitor": 0, "resistor": 0}
    issues: list[ParseIssue] = []
    skipped = 0
    try:
        worksheet, headers = select_ad_sheet(workbook, sheet)
        original = read_ad_table(workbook, sheet)
        if not original.rows:
            raise BomError("AD 原始表没有物料数据")
        created = "note" not in headers
        column = worksheet.max_column + 1 if created else headers["note"]
        if column > 16384:
            raise BomError("工作表已达到 Excel 最大列数，无法追加 note 列")
        updates = {1: "note"} if created else {}
        for row in original.rows:
            cell = worksheet.cell(row.row, column)
            if isinstance(cell, MergedCell) or any(cell.coordinate in area for area in worksheet.merged_cells.ranges):
                raise BomError(f"note 单元格 {cell.coordinate} 处于合并区域，请先取消该区域合并")
            if cell.value is not None and cell.value != "":
                skipped += 1
                continue
            if created:
                updates[row.row] = ""
            kind = component_kind(row.comment, row.libref)
            if kind is None:
                skipped += 1
                continue
            parser = parse_capacitor if kind == "capacitor" else parse_resistor
            try:
                remark = parser(row.comment)
            except BomError as error:
                issues.append(ParseIssue(row.row, row.comment, str(error), original.sheet))
                continue
            updates[row.row] = remark
            counts[kind] += 1
        data = update_notes(source.read_bytes(), original.sheet, column, updates,
                            headers["description"], new_column=created)
        save_bytes_exclusive(data, output)
        return ValueReport(output, original.sheet, original.sheet, counts["capacitor"], counts["resistor"],
                           skipped, tuple(issues), len(original.rows))
    finally:
        workbook.close()
