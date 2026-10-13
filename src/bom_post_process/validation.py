"""Validate AD BOM data and package compatibility for supported passive models.

Package checks use the project's complete resistor/capacitor naming rules. They
do not infer a package from arbitrary digits inside a manufacturer's part number.
Unknown models and footprint aliases remain explicit warnings, not verified rows.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

from .components import CAPACITOR_RE, RESISTOR_RE, normalize_model
from .errors import BomError
from .workbook import AdRow, AdTable, read_ad_workbook, tokenize_designators


# Exact aliases only: a substring such as M1206, QFN0402 or R0402/C0603 is not
# enough to assert the package of a component. L is the legacy capacitor land
# pattern suffix in names such as "C 0402L".
_PREFIX_FOOTPRINT_RE = re.compile(
    r"(?P<kind>CAP|RES|C|R)[\s_-]*(?P<package>[0-9]{4})(?P<land>L?)"
)
_SUFFIX_FOOTPRINT_RE = re.compile(r"(?P<package>[0-9]{4})[\s_-]*(?P<kind>C|R)?")


@dataclass(frozen=True)
class ValidationIssue:
    severity: str
    code: str
    row: int
    designator: str
    comment: str
    footprint: str
    message: str


@dataclass(frozen=True)
class ValidationReport:
    """A JSON-serializable report; ``ok`` only means no detected errors.

    ``checked`` counts rows whose Comment and Footprint could be compared,
    including rows with detected conflicts. ``skipped`` counts every other row.
    Warnings describe the unverified scope and must not be treated as approval
    of the complete BOM. All rows receive designator and quantity checks.
    """

    sheet: str
    checked: int
    skipped: int
    issues: tuple[ValidationIssue, ...]
    errors: int = field(init=False)
    warnings: int = field(init=False)
    ok: bool = field(init=False)

    def __post_init__(self) -> None:
        errors = sum(issue.severity == "error" for issue in self.issues)
        warnings = sum(issue.severity == "warning" for issue in self.issues)
        object.__setattr__(self, "errors", errors)
        object.__setattr__(self, "warnings", warnings)
        object.__setattr__(self, "ok", errors == 0)


def _footprint_package(footprint: str) -> tuple[str, str | None] | None:
    value = footprint.strip().upper()
    match = _PREFIX_FOOTPRINT_RE.fullmatch(value)
    if match is not None:
        kind = "C" if match["kind"] in {"C", "CAP"} else "R"
        if match["land"] and kind != "C":
            return None
        return match["package"], kind
    match = _SUFFIX_FOOTPRINT_RE.fullmatch(value)
    if match is not None:
        return match["package"], match["kind"]
    return None


def _quantity(value: object) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        quantity = Decimal(str(value).strip())
    except InvalidOperation:
        return None
    if not quantity.is_finite() or quantity <= 0 or quantity != quantity.to_integral_value():
        return None
    return int(quantity)


def validate_table(table: AdTable) -> ValidationReport:
    """Validate an already loaded AD table without writing any files."""
    if not table.rows:
        raise BomError(f"工作表“{table.sheet}”没有 BOM 数据行")
    issues: list[ValidationIssue] = []
    checked = 0
    seen_designators: dict[str, int] = {}

    def issue(row: AdRow, severity: str, code: str, message: str) -> None:
        issues.append(ValidationIssue(
            severity=severity,
            code=code,
            row=row.row,
            designator=row.designator,
            comment=row.comment,
            footprint=row.footprint,
            message=message,
        ))

    for row in table.rows:
        designators = tokenize_designators(row.designator)
        if not designators:
            issue(row, "error", "missing_designator", "Designator 位号为空")
        duplicates: set[str] = set()
        for designator in designators:
            key = designator.casefold()
            if key in seen_designators and key not in duplicates:
                issue(row, "error", "duplicate_designator",
                      f"位号 {designator} 重复，首次出现在第 {seen_designators[key]} 行")
                duplicates.add(key)
            seen_designators.setdefault(key, row.row)

        quantity = _quantity(row.quantity)
        if quantity is None:
            issue(row, "error", "invalid_quantity", "Quantity 必须是正整数")
        elif designators and quantity != len(designators):
            issue(row, "error", "quantity_mismatch",
                  f"Quantity 为 {quantity}，实际列出 {len(designators)} 个位号")

        if not row.footprint.strip():
            issue(row, "error", "missing_footprint", "Footprint 封装为空")
        if not row.comment.strip():
            issue(row, "error", "missing_comment", "Comment 料号为空")
            continue

        model = normalize_model(row.comment)
        match = CAPACITOR_RE.fullmatch(model)
        kind = "C"
        if match is None:
            match = RESISTOR_RE.fullmatch(model)
            kind = "R"
        if match is None:
            issue(row, "warning", "unsupported_comment",
                  "Comment 不符合已支持的标准电阻/电容命名规则，未验证封装一致性")
            continue

        if not row.footprint.strip():
            continue
        footprint = _footprint_package(row.footprint)
        if footprint is None:
            issue(row, "warning", "unknown_footprint",
                  "Footprint 不属于已支持的封装别名，未验证封装一致性")
            continue

        checked += 1
        package, footprint_kind = footprint
        if match["package"] != package:
            issue(row, "error", "package_mismatch",
                  f"Comment 指定 {match['package']} 封装，Footprint 指定 {package} 封装")
        if footprint_kind is not None and footprint_kind != kind:
            names = {"C": "电容", "R": "电阻"}
            issue(row, "error", "component_type_mismatch",
                  f"Comment 为{names[kind]}，Footprint 别名表示{names[footprint_kind]}")

    return ValidationReport(table.sheet, checked, len(table.rows) - checked, tuple(issues))


def validate_workbook(
    input_path: Path | str, *, sheet: str | None = None
) -> ValidationReport:
    """Read an AD23 export and report conflicts without changing the workbook."""
    return validate_table(read_ad_workbook(input_path, sheet=sheet))
