"""Append editable review columns to the AD table without restyling it."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl.cell.cell import MergedCell

from .components import component_kind, component_name, is_nc, parse_capacitor, parse_resistor
from .errors import BomError
from .workbook import (
    open_workbook, prepare_output, read_ad_table, save_bytes_exclusive, select_ad_sheet,
)
from .xlsx_notes import update_review_columns


REVIEW_HEADERS = ("name", "note", "not_bought", "not_welded")


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
    """Preserve the original table and add fields for manual review."""
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
        present = set(REVIEW_HEADERS).intersection(headers)
        legacy_note = None
        if not present:
            first_column = worksheet.max_column + 1
            created = set(REVIEW_HEADERS)
        elif present == {"note"} and headers["note"] == worksheet.max_column:
            first_column = headers["note"]
            legacy_note = first_column
            created = {"name", "not_bought", "not_welded"}
        elif present == set(REVIEW_HEADERS):
            first_column = headers["name"]
            if any(headers[name] != first_column + offset for offset, name in enumerate(REVIEW_HEADERS)):
                raise BomError("审核列顺序应为 name、note、not_bought、not_welded，请先整理这些列")
            created = set()
        else:
            raise BomError("审核列不完整或旧版 note 不是末列，请保留完整的 name、note、not_bought、not_welded 四列")
        columns = {name: first_column + offset for offset, name in enumerate(REVIEW_HEADERS)}
        if columns["not_welded"] > 16384:
            raise BomError("工作表已达到 Excel 最大列数，无法追加审核列")
        updates = {columns[name]: {1: name} for name in created}
        note_updates = updates.setdefault(columns["note"], {})
        for row in original.rows:
            for name in REVIEW_HEADERS:
                source_column = legacy_note if name == "note" and legacy_note else columns[name]
                cell = worksheet.cell(row.row, source_column)
                if isinstance(cell, MergedCell) or any(cell.coordinate in area for area in worksheet.merged_cells.ranges):
                    raise BomError(f"{name} 单元格 {cell.coordinate} 处于合并区域，请先取消该区域合并")
                if name in created:
                    value = ""
                    if name == "name":
                        value = component_name(row.comment, row.libref)
                    elif name == "not_welded" and is_nc(row.comment):
                        value = "NC"
                    updates[columns[name]][row.row] = value
            note_cell = worksheet.cell(row.row, legacy_note or columns["note"])
            if note_cell.value is not None and note_cell.value != "":
                skipped += 1
                continue
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
            note_updates[row.row] = remark
            counts[kind] += 1
        widths = {columns[name]: width for name, width in zip(REVIEW_HEADERS, (22, 42, 16, 16))
                  if name in created}
        data = update_review_columns(source.read_bytes(), original.sheet, updates,
                                     headers["description"], widths=widths, legacy_note=legacy_note)
        save_bytes_exclusive(data, output)
        return ValueReport(output, original.sheet, original.sheet, counts["capacitor"], counts["resistor"],
                           skipped, tuple(issues), len(original.rows))
    finally:
        workbook.close()
