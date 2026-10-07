import json

import pytest

from bom_enricher.component_data import JsonCatalogProvider
from bom_enricher.models import ComponentData, ComponentLookupResult, CustomerBomRow
from bom_enricher.validation import validate_row


def matched_capacitor() -> ComponentLookupResult:
    return ComponentLookupResult(
        status="matched",
        component=ComponentData(
            part_number="SYN-CAP-001",
            manufacturer="Example Components Ltd. (fictional)",
            short_description="10 uF 16 V ceramic capacitor",
            attributes={
                "capacitance": "10 uF",
                "voltage": "16 V",
                "dielectric": "X7R",
                "package": "0805",
                "mounting": "SMD",
            },
        ),
    )


def matched_resistor() -> ComponentLookupResult:
    return ComponentLookupResult(
        status="matched",
        component=ComponentData(
            part_number="SYN-RES-002",
            manufacturer="Imaginary Passive Works (fictional)",
            short_description="1 kOhm 1% resistor",
            attributes={
                "resistance": "1 kOhm",
                "tolerance": "1%",
                "power": "0.1 W",
                "package": "0603",
                "mounting": "SMD",
            },
        ),
    )


def matched_capacitor_with_case_package() -> ComponentLookupResult:
    return ComponentLookupResult(
        status="matched",
        component=ComponentData(
            part_number="SYN-CAP-001",
            manufacturer="Example Components Ltd. (fictional)",
            short_description="10 uF 16 V ceramic capacitor",
            attributes={
                "capacitance": "10 uF",
                "voltage": "16 V",
                "dielectric": "X7R",
                "case_package": "0805 (2012 Metric)",
                "mounting": "SMD",
            },
        ),
    )


def test_unknown_and_ambiguous_part_number(tmp_path):
    catalog_path = tmp_path / "catalog.json"
    catalog_path.write_text(
        json.dumps(
            {
                "SYN-CAP-001": {
                    "manufacturer": "Example Components Ltd. (fictional)",
                    "short_description": "10 uF 16 V ceramic capacitor",
                    "attributes": {"capacitance": "10 uF", "voltage": "16 V"},
                }
            }
        ),
        encoding="utf-8",
    )
    provider = JsonCatalogProvider(catalog_path)

    unknown = provider.lookup_many(["unknown-part"])["unknown-part"]
    assert unknown.status == "not_found"
    unknown_result = validate_row(
        CustomerBomRow("C1", 1, "", "unknown-part"), unknown
    )
    assert unknown_result.description == ""
    assert any(
        issue.field == "part_number" and issue.status == "unverified"
        for issue in unknown_result.issues
    )

    ambiguous_result = validate_row(
        CustomerBomRow("C1", 1, "", "SYN-CAP-001"),
        ComponentLookupResult(status="ambiguous", component=None, message="Two exact matches"),
    )
    assert ambiguous_result.description == ""
    assert any(
        issue.field == "part_number" and issue.status == "unverified"
        for issue in ambiguous_result.issues
    )


def test_not_found_empty_description_is_unverified_for_review():
    result = validate_row(
        CustomerBomRow("C1", 1, "", "UNKNOWN"),
        ComponentLookupResult(status="not_found", component=None),
    )

    assert result.description == ""
    assert any(
        issue.field == "description"
        and issue.status == "unverified"
        and issue.original == ""
        and issue.proposed == ""
        for issue in result.issues
    )


def test_description_add_and_correction_require_match():
    matched = matched_capacitor()

    missing = validate_row(CustomerBomRow("C1", 1, "", "SYN-CAP-001"), matched)
    assert missing.description == "10 uF 16 V ceramic capacitor"
    assert any(
        issue.field == "description"
        and issue.status == "added"
        and issue.proposed == "10 uF 16 V ceramic capacitor"
        for issue in missing.issues
    )

    incorrect = validate_row(
        CustomerBomRow("C1", 1, "1 uF 50 V ceramic capacitor", "SYN-CAP-001"),
        matched,
    )
    assert incorrect.description == "10 uF 16 V ceramic capacitor"
    assert any(
        issue.field == "description" and issue.status == "corrected"
        for issue in incorrect.issues
    )

    unverified = validate_row(
        CustomerBomRow("C1", 1, "1 uF 50 V ceramic capacitor", "SYN-CAP-001"),
        ComponentLookupResult(status="unverified", component=None, message="Offline error"),
    )
    assert unverified.description == "1 uF 50 V ceramic capacitor"
    assert any(
        issue.field == "description" and issue.status == "unverified"
        for issue in unverified.issues
    )


def test_description_preserves_correct_extended_capacitor_attributes():
    source_description = "10 uF 16 V X7R 0805 ceramic capacitor"

    result = validate_row(
        CustomerBomRow("C1", 1, source_description, "SYN-CAP-001"),
        matched_capacitor(),
    )

    assert result.description == source_description
    assert not any(issue.field == "description" for issue in result.issues)


def test_description_corrects_explicit_conflicting_capacitor_dielectric_and_package():
    result = validate_row(
        CustomerBomRow(
            "C1", 1, "10 uF 16 V X5R 0603 ceramic capacitor", "SYN-CAP-001"
        ),
        matched_capacitor(),
    )

    assert result.description == "10 uF 16 V ceramic capacitor"
    assert any(
        issue.field == "description"
        and issue.status == "corrected"
        and issue.proposed == "10 uF 16 V ceramic capacitor"
        for issue in result.issues
    )


def test_description_validates_nexar_case_package_alias():
    conflicting = validate_row(
        CustomerBomRow("C1", 1, "10 uF 16 V X7R 0603 ceramic capacitor", "SYN-CAP-001"),
        matched_capacitor_with_case_package(),
    )
    assert conflicting.description == "10 uF 16 V ceramic capacitor"
    assert any(
        issue.field == "description" and issue.status == "corrected"
        for issue in conflicting.issues
    )

    matching_description = "10 uF 16 V X7R 0805 ceramic capacitor"
    matching = validate_row(
        CustomerBomRow("C1", 1, matching_description, "SYN-CAP-001"),
        matched_capacitor_with_case_package(),
    )
    assert matching.description == matching_description
    assert not any(issue.field == "description" for issue in matching.issues)

    full_equivalent_description = "10 uF 16 V X7R 0805 (2012 Metric) ceramic capacitor"
    full_equivalent = validate_row(
        CustomerBomRow("C1", 1, full_equivalent_description, "SYN-CAP-001"),
        matched_capacitor_with_case_package(),
    )
    assert full_equivalent.description == full_equivalent_description
    assert not any(issue.field == "description" for issue in full_equivalent.issues)


def test_description_handles_explicit_resistor_power_and_package_without_guessing():
    equivalent_power = "1 kOhm 1% 0.1W 0603 resistor"
    preserved = validate_row(
        CustomerBomRow("R1", 1, equivalent_power, "SYN-RES-002"), matched_resistor()
    )
    assert preserved.description == equivalent_power
    assert not any(issue.field == "description" for issue in preserved.issues)

    conflicting = validate_row(
        CustomerBomRow("R1", 1, "1 kOhm 1% 0.25 W 0402 resistor", "SYN-RES-002"),
        matched_resistor(),
    )
    assert conflicting.description == "1 kOhm 1% resistor"
    assert any(
        issue.field == "description" and issue.status == "corrected"
        for issue in conflicting.issues
    )


def test_quantity_matches_reference_count():
    matching = validate_row(CustomerBomRow("C1, C2", "2", "", "SYN-CAP-001"), matched_capacitor())
    assert not any(issue.field == "quantity" for issue in matching.issues)

    mismatch = validate_row(CustomerBomRow("C1, C2", 1, "", "SYN-CAP-001"), matched_capacitor())
    assert any(
        issue.field == "quantity" and issue.status == "conflict"
        for issue in mismatch.issues
    )

    invalid = validate_row(CustomerBomRow("C1", 0, "", "SYN-CAP-001"), matched_capacitor())
    assert any(
        issue.field == "quantity" and issue.status == "invalid"
        for issue in invalid.issues
    )


def test_mounting_conflict_and_unknown():
    conflict = validate_row(
        CustomerBomRow("C1", 1, "10 uF 16 V SMD capacitor", "SYN-CONN-003"),
        ComponentLookupResult(
            status="matched",
            component=ComponentData(
                part_number="SYN-CONN-003",
                manufacturer="Placeholder Devices (fictional)",
                short_description="2-position through-hole connector",
                attributes={"positions": "2", "mounting": "TH"},
            ),
        ),
    )
    assert conflict.mounting == "TH"
    assert any(
        issue.field == "mounting" and issue.status == "conflict"
        for issue in conflict.issues
    )

    unknown = validate_row(
        CustomerBomRow("C1", 1, "connector", "SYN-CONN-003"),
        ComponentLookupResult(
            status="matched",
            component=ComponentData(
                part_number="SYN-CONN-003",
                manufacturer="Placeholder Devices (fictional)",
                short_description="2-position connector",
                attributes={"positions": "2"},
            ),
        ),
    )
    assert unknown.mounting == ""
    assert any(
        issue.field == "mounting" and issue.status == "unverified"
        for issue in unknown.issues
    )


def test_json_catalog_rejects_duplicate_normalized_part_numbers(tmp_path):
    catalog_path = tmp_path / "catalog.json"
    catalog_path.write_text(
        json.dumps(
            {
                "SYN-CAP-001": {"manufacturer": "A", "short_description": "A", "attributes": {}},
                " syn-cap-001 ": {"manufacturer": "B", "short_description": "B", "attributes": {}},
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="duplicate normalized part number"):
        JsonCatalogProvider(catalog_path)
