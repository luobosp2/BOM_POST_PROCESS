"""Export the manually reviewed rows selected for soldering."""

from pathlib import Path

from .reviewed_export import ExportReport, export_reviewed


SolderingReport = ExportReport


def gen_soldering(
    input_path: Path | str, output_path: Path | str | None = None, sheet: str | None = None,
) -> SolderingReport:
    """Exclude only rows whose not_welded cell contains a value."""
    return export_reviewed(input_path, output_path, sheet, flag="not_welded", title="焊接表")
