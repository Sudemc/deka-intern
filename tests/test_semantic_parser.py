"""Tests for semantic electrical component specification parser and comparison engine."""

from decimal import Decimal
import pytest

from bom_enricher.semantic_parser import (
    ElectricalSpecs,
    MultiElectricalSpecs,
    compare_descriptions,
    parse_specs,
    parse_specs_set,
)
from bom_enricher.comparison import split_multi_code
from bom_enricher.excel_io import is_valid_part_number


def test_multi_spec_extraction_ferrite_bead():
    text = "220 Ohms @ 100 MHz 1 Power Line Ferrite Bead 0805 (2012 Metric) 2A 45mOhm"
    specs = parse_specs_set(text)

    assert isinstance(specs, MultiElectricalSpecs)
    assert Decimal("220") in specs.resistances
    assert Decimal("0.045") in specs.resistances
    assert "0805" in specs.packages


def test_multi_spec_extraction_tantalum_cap():
    text = "10 \u00b5F Molded Tantalum Capacitors 16 V 1210 (3225 Metric) 800mOhm @ 100kHz"
    specs = parse_specs_set(text)

    assert isinstance(specs, MultiElectricalSpecs)
    assert Decimal("1e-5") in specs.capacitances
    assert Decimal("16") in specs.voltages
    assert "1210" in specs.packages
    assert Decimal("0.8") in specs.resistances


def test_backward_compatible_parse_specs():
    text = "100nF 50V 0805"
    legacy_specs = parse_specs(text)

    assert isinstance(legacy_specs, ElectricalSpecs)
    assert legacy_specs.capacitance_farad == Decimal("1e-7")
    assert legacy_specs.voltage_volt == Decimal("50")
    assert legacy_specs.package == "0805"


def test_ferrite_bead_multi_value_compatibility():
    cust = "220 Ohms @ 100 MHz 1 Power Line Ferrite Bead 0805 (2012 Metric) 2A 45mOhm"
    erp = "SMD 0805 Package, F.Bead: 220R, 2A, 45mOhm, Murata"

    issues = compare_descriptions(cust, erp)
    assert issues == [], f"Expected no contradiction but got: {issues}"


def test_tantalum_cap_compatibility():
    cust = "10 \u00b5F Molded Tantalum Capacitors 16 V 1210 (3225 Metric) 800mOhm @ 100kHz"
    erp = "10uF 16V 1210"

    issues = compare_descriptions(cust, erp)
    assert issues == [], f"Expected no contradiction but got: {issues}"


def test_genuine_resistance_mismatch_flagged():
    cust = "Resistor 100R 0805"
    erp = "Resistor 10k 0805"

    issues = compare_descriptions(cust, erp, mpn_matched=False)
    assert any("Direnç farklı" in issue for issue in issues)

    # Even with mpn_matched=True, severe contradiction (100R vs 10k) MUST still be flagged
    issues_shielded = compare_descriptions(cust, erp, mpn_matched=True)
    assert any("Direnç farklı" in issue for issue in issues_shielded)


def test_mpn_matched_shield_voltage_clamp_ratings():
    cust = "TVS Diode 5V Standoff"
    erp = "TVS Diode 9.2V Clamping"

    issues_unmatched = compare_descriptions(cust, erp, mpn_matched=False)
    assert any("Voltaj farklı" in issue for issue in issues_unmatched)

    issues_matched = compare_descriptions(cust, erp, mpn_matched=True)
    assert issues_matched == [], f"Expected shield to suppress clamp voltage variation, got: {issues_matched}"


def test_mpn_matched_shield_tolerance_variation():
    cust = "Resistor 10k 1%"
    erp = "Resistor 10k 5%"

    issues_unmatched = compare_descriptions(cust, erp, mpn_matched=False)
    assert any("Tolerans farklı" in issue for issue in issues_unmatched)

    issues_matched = compare_descriptions(cust, erp, mpn_matched=True)
    assert issues_matched == [], f"Expected shield to suppress tolerance difference, got: {issues_matched}"


def test_zero_ohm_resistor_comparison():
    # 0R vs 0R should be fully compatible
    issues_same = compare_descriptions("Resistor 0R 0805", "Resistor 0R 0805")
    assert issues_same == []

    # 0R vs 10R should flag a contradiction without DivisionByZero
    issues_diff = compare_descriptions("Resistor 0R 0805", "Resistor 10R 0805")
    assert any("Direnç farklı" in issue for issue in issues_diff)
    assert any("0R vs değerli direnç farkı" in issue for issue in issues_diff)


def test_exposed_helper_functions():
    assert callable(split_multi_code)
    assert callable(is_valid_part_number)
    codes = split_multi_code("RC0603FR-0710KL / BLM21PG221SN1D")
    assert "RC0603FR-0710KL" in codes
    assert "BLM21PG221SN1D" in codes
    assert is_valid_part_number("RC0603FR-0710KL") is True
    assert is_valid_part_number("SEE NOTE") is False


def test_jedec_diodes_and_cap_rcodes_do_not_parse_as_inductance():
    from bom_enricher.semantic_parser import parse_specs_set, compare_descriptions
    s_diode = parse_specs_set("1N4148")
    assert not s_diode.capacitances
    assert not s_diode.inductances
    s_cap = parse_specs_set("4u7 16V 0603")
    assert s_cap.capacitances and not s_cap.inductances
    assert compare_descriptions("1N4148", "1N4007 Rectifier Diode", mpn_matched=False) == []


def test_megaohm_vs_milliohm_and_resistor_arrays():
    from decimal import Decimal
    from bom_enricher.semantic_parser import parse_specs_set
    assert Decimal("10000000") in parse_specs_set("10M").resistances
    assert Decimal("10000000") in parse_specs_set("10MOhm").resistances
    assert Decimal("0.01") in parse_specs_set("10mOhm").resistances
    assert Decimal("10000") in parse_specs_set("4x10K 0603").resistances


def test_cross_category_color_polarity_frequency_and_mpn_matched_value_conflicts():
    from bom_enricher.semantic_parser import compare_descriptions
    assert compare_descriptions("100nF", "RES SMD 100R 1% 0402", mpn_matched=False)
    assert compare_descriptions("Red LED 0603", "Yellow LED 0603", mpn_matched=False)
    assert compare_descriptions("NPN SOT-23", "PNP SOT-23", mpn_matched=False)
    assert compare_descriptions("8MHz", "16MHz Crystal", mpn_matched=False)
    assert compare_descriptions("4.7K 0603", "SMD 0603, 10k, %1, Yageo", mpn_matched=True)
    assert compare_descriptions("100nF 0402", "SMD 0805 Package, 100nF", mpn_matched=True)

