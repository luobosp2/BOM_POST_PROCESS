"""Parse the project's resistor and MLCC naming conventions.

K/L/U/M in resistor names are the legacy temperature-coefficient codes.
The legacy project convention supplies 1% tolerance; those letters are not
interpreted as an industry-standard tolerance code.
"""

from __future__ import annotations

from decimal import Decimal
import re

from .errors import BomError


CAPACITOR_RE = re.compile(
    r"(?P<package>[0-9]{4})-"
    r"(?P<material>X5R|X7R|C0G|COG|NP0)"
    r"(?P<capacity>[0-9]{3}|[0-9]R[0-9])"
    r"(?P<tolerance>[JKM])"
    r"(?P<voltage>[0-9]{3}|[0-9]R[0-9])"
)
RESISTOR_RE = re.compile(
    r"(?P<package>[0-9]{4})[KLUM](?P<resistance>[0-9]{4}|[0-9]+R[0-9]*)"
)
NC_SUFFIX_RE = re.compile(r"(?:/NC|\(NC\))$", re.IGNORECASE)
CAPACITOR_TOLERANCES = {"J": "5%", "K": "10%", "M": "20%"}


def normalize_model(value: object) -> str:
    if not isinstance(value, str) or not value.strip():
        raise BomError("Comment/规格型号为空或不是文本")
    return NC_SUFFIX_RE.sub("", value.strip().upper()).rstrip()


def is_nc(value: object) -> bool:
    return isinstance(value, str) and bool(NC_SUFFIX_RE.search(value.strip()))


def format_number(value: Decimal) -> str:
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def format_scaled(value: Decimal, units: tuple[tuple[int, str], ...]) -> str:
    for divisor, suffix in units:
        if value >= divisor:
            return format_number(value / Decimal(divisor)) + suffix
    return format_number(value) + units[-1][1]


def parse_capacitor(model: object) -> str:
    match = CAPACITOR_RE.fullmatch(normalize_model(model))
    if match is None:
        raise BomError("不符合标准电容命名规则，未推断厂商专用编码的参数")
    capacity_code = match["capacity"]
    capacity_pf = (
        Decimal(capacity_code.replace("R", "."))
        if "R" in capacity_code
        else Decimal(int(capacity_code[:2]) * 10 ** int(capacity_code[2]))
    )
    voltage_code = match["voltage"]
    voltage = (
        Decimal(voltage_code.replace("R", "."))
        if "R" in voltage_code else Decimal(voltage_code) / Decimal(10)
    )
    material = match["material"]
    if material in {"COG", "NP0"}:
        material = "C0G"
    capacity = format_scaled(capacity_pf, ((1_000_000, "uF"), (1_000, "nF"), (1, "pF")))
    tolerance = CAPACITOR_TOLERANCES[match["tolerance"]]
    return f"{capacity}/{format_number(voltage)}V/{match['package']}/{material}/{tolerance}"


def parse_resistor(model: object) -> str:
    match = RESISTOR_RE.fullmatch(normalize_model(model))
    if match is None:
        raise BomError("不符合标准电阻命名规则")
    code = match["resistance"]
    resistance = (
        Decimal(code.replace("R", "."))
        if "R" in code else Decimal(int(code[:3]) * 10 ** int(code[3]))
    )
    value = format_scaled(resistance, ((1_000_000, "M"), (1_000, "k"), (1, "")))
    return f"{value}/{match['package']}/1%"


def component_kind(comment: object, libref: str = "", name: str = "") -> str | None:
    """Recognize supported categories, without treating every four digits as a package."""
    if name.strip() == "贴片电容":
        return "capacitor"
    if name.strip() == "贴片电阻":
        return "resistor"
    try:
        model = normalize_model(comment)
    except BomError:
        model = ""
    if CAPACITOR_RE.fullmatch(model):
        return "capacitor"
    if RESISTOR_RE.fullmatch(model):
        return "resistor"
    normalized_libref = libref.strip().casefold()
    if normalized_libref in {"cap", "capacitor"}:
        return "capacitor"
    if normalized_libref in {"res", "res1", "resistor"}:
        return "resistor"
    return None


def component_name(comment: str, libref: str) -> str:
    kind = component_kind(comment, libref)
    if kind == "capacitor":
        return "贴片电容"
    if kind == "resistor":
        return "贴片电阻"
    normalized = libref.strip().casefold()
    if normalized.startswith("cap pol"):
        return "极性电容"
    if normalized.startswith("ferrite bead"):
        return "磁珠"
    if normalized.startswith("osc"):
        return "晶振"
    return {
        "led": "发光二极管", "inductor": "电感", "npn": "三极管",
        "bitvs": "TVS二极管", "1-ch esd": "ESD保护器件",
        "schottkydiode_0": "肖特基二极管", "res varistor": "压敏电阻",
        "markpoint": "定位点", "sw-pb": "按键", "sw-4": "拨码开关",
    }.get(normalized, libref.strip() or "其他器件")
