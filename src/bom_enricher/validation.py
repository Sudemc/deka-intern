from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re

from .component_data import normalize_part_number
from .models import CheckIssue, ComponentData, ComponentLookupResult, CustomerBomRow, RowResult


def validate_row(row: CustomerBomRow, lookup: ComponentLookupResult) -> RowResult:
    """Validate one customer row without inventing missing component facts."""
    issues: list[CheckIssue] = []
    _validate_quantity(row, issues)

    component = _verified_component(row, lookup)
    if component is None:
        detail = lookup.message or f"Lookup status is {lookup.status}."
        issues.append(
            CheckIssue(
                field="part_number",
                status="unverified",
                original=row.part_number,
                proposed=row.part_number,
                message=detail,
            )
        )
        issues.append(
            CheckIssue(
                field="description",
                status="unverified",
                original=row.value,
                proposed=row.value,
                message="Description was retained because component identity is unverified.",
            )
        )
        issues.append(
            CheckIssue(
                field="mounting",
                status="unverified",
                original="",
                proposed="",
                message="Mounting is unresolved because component identity is unverified.",
            )
        )
        return RowResult(source=row, description=row.value, mounting="", issues=issues)

    description = _validate_description(row.value, component, issues)
    mounting = _validate_mounting(row.value, component, issues)
    return RowResult(source=row, description=description, mounting=mounting, issues=issues)


def _verified_component(
    row: CustomerBomRow, lookup: ComponentLookupResult
) -> ComponentData | None:
    component = lookup.component
    if lookup.status != "matched" or component is None:
        return None
    if normalize_part_number(row.part_number) != normalize_part_number(component.part_number):
        return None
    if row.manufacturer.strip() and _normalize_text(row.manufacturer) != _normalize_text(
        component.manufacturer
    ):
        return None
    return component


def _validate_quantity(row: CustomerBomRow, issues: list[CheckIssue]) -> None:
    references = [reference.strip() for reference in row.part.split(",") if reference.strip()]
    try:
        quantity = Decimal(str(row.qty).strip())
    except (InvalidOperation, ValueError):
        quantity = None

    if quantity is None or not quantity.is_finite() or quantity <= 0:
        issues.append(
            CheckIssue(
                field="quantity",
                status="invalid",
                original=str(row.qty),
                proposed=str(row.qty),
                message="Qty must be a positive numeric value.",
            )
        )
        return
    if not references:
        issues.append(
            CheckIssue(
                field="part",
                status="invalid",
                original=row.part,
                proposed=row.part,
                message="Part must contain at least one board reference.",
            )
        )
        return
    if quantity != Decimal(len(references)):
        issues.append(
            CheckIssue(
                field="quantity",
                status="conflict",
                original=str(row.qty),
                proposed=str(len(references)),
                message="Qty does not match the count of comma-separated board references.",
            )
        )


def _validate_description(
    source_description: str, component: ComponentData, issues: list[CheckIssue]
) -> str:
    proposed = component.short_description.strip()
    if not _has_deterministic_description_evidence(component):
        issues.append(
            CheckIssue(
                field="description",
                status="unverified",
                original=source_description,
                proposed=source_description,
                message="Description was retained because structured attributes are insufficient.",
            )
        )
        return source_description

    if not source_description.strip():
        issues.append(
            CheckIssue(
                field="description",
                status="added",
                original=source_description,
                proposed=proposed,
                message="Description was added from an exact match with supported attributes.",
            )
        )
        return proposed

    if _description_agrees(source_description, component):
        return source_description

    issues.append(
        CheckIssue(
            field="description",
            status="corrected",
            original=source_description,
            proposed=proposed,
            message="Description was corrected from an exact match with supported attributes.",
        )
    )
    return proposed


def _has_deterministic_description_evidence(component: ComponentData) -> bool:
    description = component.short_description.strip()
    attributes = component.attributes
    lowered_description = description.casefold()
    if "capacitor" in lowered_description:
        required = ("capacitance", "voltage")
    elif "resistor" in lowered_description:
        required = ("resistance", "tolerance")
    elif "connector" in lowered_description:
        required = ("positions",)
    else:
        return False
    return bool(description) and all(
        attributes.get(name, "").strip()
        and attributes[name].casefold() in lowered_description
        for name in required
    )


def _description_agrees(source_description: str, component: ComponentData) -> bool:
    lowered_source = source_description.casefold()
    lowered_component = component.short_description.casefold()
    if "capacitor" in lowered_component:
        required = ("capacitance", "voltage")
        family = "capacitor"
    elif "resistor" in lowered_component:
        required = ("resistance", "tolerance")
        family = "resistor"
    elif "connector" in lowered_component:
        required = ("positions",)
        family = "connector"
    else:
        return False
    if not all(component.attributes[name].casefold() in lowered_source for name in required):
        return False
    return not _has_explicit_optional_attribute_conflict(
        source_description, component, family
    )


def _has_explicit_optional_attribute_conflict(
    source_description: str, component: ComponentData, family: str
) -> bool:
    if family == "capacitor":
        return _has_conflicting_token(
            source_description, component.attributes.get("dielectric", ""), r"\bX[0-9][A-Z]\b"
        ) or _has_conflicting_package(
            source_description, _component_package(component)
        )
    if family == "resistor":
        return _has_conflicting_package(
            source_description, _component_package(component)
        ) or _has_conflicting_power(
            source_description,
            component.attributes.get("power", ""),
        )
    return False


def _has_conflicting_token(source_description: str, expected: str, pattern: str) -> bool:
    if not expected.strip():
        return False
    source_tokens = re.findall(pattern, source_description, flags=re.IGNORECASE)
    return bool(source_tokens) and any(
        token.casefold() != expected.casefold() for token in source_tokens
    )


def _component_package(component: ComponentData) -> str:
    return component.attributes.get("package") or component.attributes.get(
        "case_package", ""
    )


def _has_conflicting_package(source_description: str, expected_display: str) -> bool:
    expected_tokens = set(re.findall(r"\b[0-9]{4}\b", expected_display))
    if not expected_tokens:
        return False
    source_tokens = set(re.findall(r"\b[0-9]{4}\b", source_description))
    return bool(source_tokens) and not source_tokens.issubset(expected_tokens)


def _has_conflicting_power(source_description: str, expected: str) -> bool:
    expected_power = _parse_power(expected)
    if expected_power is None:
        return False
    source_tokens = re.findall(
        r"\b(?:[0-9]+(?:\.[0-9]+)?|[0-9]+/[0-9]+)\s*W\b",
        source_description,
        flags=re.IGNORECASE,
    )
    return any(
        (source_power := _parse_power(token)) is not None
        and source_power != expected_power
        for token in source_tokens
    )


def _parse_power(value: str) -> Decimal | None:
    match = re.fullmatch(
        r"\s*([0-9]+(?:\.[0-9]+)?|[0-9]+/[0-9]+)\s*W\s*",
        value,
        flags=re.IGNORECASE,
    )
    if match is None:
        return None
    numeric_value = match.group(1)
    try:
        if "/" in numeric_value:
            numerator, denominator = numeric_value.split("/", maxsplit=1)
            return Decimal(numerator) / Decimal(denominator)
        return Decimal(numeric_value)
    except (InvalidOperation, ZeroDivisionError):
        return None


def _validate_mounting(
    source_description: str, component: ComponentData, issues: list[CheckIssue]
) -> str:
    mounting = _canonical_mounting(component.attributes.get("mounting", ""))
    if not mounting:
        issues.append(
            CheckIssue(
                field="mounting",
                status="unverified",
                original="",
                proposed="",
                message="Mounting is unresolved because no clear supported mounting attribute exists.",
            )
        )
        return ""

    source_clue = _mounting_clue(source_description)
    if source_clue and source_clue != mounting:
        issues.append(
            CheckIssue(
                field="mounting",
                status="conflict",
                original=source_clue,
                proposed=mounting,
                message="Value contains an explicit mounting clue that conflicts with component data.",
            )
        )
    return mounting


def _canonical_mounting(value: str) -> str:
    normalized = " ".join(value.casefold().replace("-", " ").split())
    if normalized in {"th", "through hole", "throughhole"}:
        return "TH"
    if normalized in {"smd", "surface mount", "surfacemount"}:
        return "SMD"
    return ""


def _mounting_clue(value: str) -> str:
    lowered_value = value.casefold()
    if re.search(r"\b(smd|surface[ -]?mount)\b", lowered_value):
        return "SMD"
    if re.search(r"\b(th|through[ -]?hole)\b", lowered_value):
        return "TH"
    return ""


def _normalize_text(value: str) -> str:
    return " ".join(value.split()).casefold()
