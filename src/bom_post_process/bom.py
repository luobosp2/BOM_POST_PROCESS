"""Export the manually reviewed rows selected for purchasing."""

from pathlib import Path

from .reviewed_export import ExportReport, export_reviewed


BomReport = ExportReport


def gen_bom(
    input_path: Path | str, output_path: Path | str | None = None, sheet: str | None = None,
) -> BomReport:
    """Exclude only rows whose not_bought cell contains a value, adding a PCB placeholder if absent."""
    return export_reviewed(input_path, output_path, sheet, flag="not_bought", title="BOM", add_pcb=True)
