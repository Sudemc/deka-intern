"""Tests for datasheet PDF analysis and ERP property mapping."""

import io
from pathlib import Path

from openpyxl import load_workbook
import pytest
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as rl_canvas

from bom_enricher.comparison import load_rulebook
from bom_enricher.datasheet_analyzer import (
    analyze_single_datasheet,
    classify_material_type,
    decode_manufacturer_pn,
    extract_component_specs,
    extract_part_numbers_from_text,
    extract_pdf_data,
    generate_datasheet_excel_report,
    is_automotive_qualified,
    is_catalog_document_name,
    map_specs_to_erp,
)


@pytest.fixture(scope="module")
def rulebook():
    rules_path = Path("data/rules.xlsx")
    return load_rulebook(rules_path)


def _create_sample_pdf(title: str, pn: str, bullet_points: list[str]) -> bytes:
    buffer = io.BytesIO()
    c = rl_canvas.Canvas(buffer, pagesize=letter)
    c.drawString(100, 750, title)
    c.drawString(100, 730, f"Part Number: {pn}")
    y = 700
    for point in bullet_points:
        c.drawString(100, y, f"- {point}")
        y -= 20
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def test_decode_samsung_capacitor_part_number():
    specs = decode_manufacturer_pn("CL21B104KBCNNNC")
    assert specs["manufacturer"] == "Samsung"
    assert specs["package"] == "0805"
    assert specs["capacitance_str"] == "100nF"
    assert specs["tolerance"] == "10%"
    assert specs["voltage"] == "50V"
    assert specs["special_def"] == "X7R"


def test_decode_yageo_capacitor_part_number():
    specs = decode_manufacturer_pn("CC0805KRX7R9BB104")
    assert specs["manufacturer"] == "Yageo"
    assert specs["package"] == "0805"
    assert specs["tolerance"] == "10%"
    assert specs["voltage"] == "50V"
    assert specs["special_def"] == "X7R"
    assert specs["capacitance_str"] == "100nF"


def test_decode_yageo_resistor_part_number():
    specs = decode_manufacturer_pn("RC0805FR-0710KL")
    assert specs["manufacturer"] == "Yageo"
    assert specs["package"] == "0805"
    assert specs["tolerance"] == "1%"
    assert specs["resistance_str"] == "10k"


def test_classify_material_types():
    assert classify_material_type("Some text", "CL21B104KBCNNNC")[0] == "CAP"
    assert classify_material_type("Some text", "RC0805FR-0710KL")[0] == "RES"
    assert classify_material_type("Some text", "BAT54S")[0] == "DIO"
    assert classify_material_type("Some text", "2N7002")[0] == "FET"
    assert classify_material_type("Multilayer Ceramic Capacitor 100nF 50V", "XYZ-123")[0] == "CAP"
    assert classify_material_type("Thick Film Chip Resistor 10k 1%", "XYZ-123")[0] == "RES"
    assert classify_material_type("Schottky Barrier Diode 30V 200mA", "XYZ-123")[0] == "DIO"


def test_pdf_extraction_and_analysis_end_to_end(rulebook):
    pdf_bytes = _create_sample_pdf(
        "SAMSUNG ELECTRO-MECHANICS MULTILAYER CERAMIC CAPACITOR",
        "CL21B104KBCNNNC",
        [
            "SMD Surface Mount Capacitor",
            "Size: 0805 (2012 metric)",
            "Capacitance: 100nF (0.1uF)",
            "Rated Voltage: 50V",
            "Tolerance: +/-10%",
            "Dielectric: X7R",
            "Operating Temperature: -55C to +125C",
            "Manufacturer: Samsung Electro-Mechanics",
        ],
    )

    result = analyze_single_datasheet(
        filename="CL21B104KBCNNNC.pdf",
        file_bytes=pdf_bytes,
        rulebook=rulebook,
    )

    assert result.material_type == "CAP"
    assert result.inferred_mounting == "SMD"
    assert result.confidence >= 0.9
    assert len(result.properties) >= 4
    # Check that suggested description contains key ERP values
    assert "0805 Package" in result.suggested_description
    assert "100nF" in result.suggested_description
    assert "50V" in result.suggested_description
    assert "X7R" in result.suggested_description
    assert "Samsung" in result.suggested_description


def test_excel_report_generation(rulebook, tmp_path):
    pdf_bytes = _create_sample_pdf(
        "Yageo Chip Resistor",
        "RC0805FR-0710KL",
        [
            "Thick Film Chip Resistor",
            "Size: 0805",
            "Resistance: 10k ohm",
            "Tolerance: 1%",
            "Power Rating: 1/8W",
            "Manufacturer: Yageo",
        ],
    )

    res = analyze_single_datasheet("RC0805FR-0710KL.pdf", pdf_bytes, rulebook)
    output_excel = tmp_path / "Datasheet_Rapor.xlsx"
    generate_datasheet_excel_report([res], output_excel)

    assert output_excel.exists()
    wb = load_workbook(output_excel)
    assert "Datasheet ERP Eşleme Özeti" in wb.sheetnames
    assert "Özellik Eşleme Detayları" in wb.sheetnames
    ws = wb["Datasheet ERP Eşleme Özeti"]
    assert ws.max_row == 2
    assert ws.cell(row=2, column=3).value == "RC0805FR-0710KL"
    wb.close()


def test_decode_royal_ohm_resistor_part_number():
    specs = decode_manufacturer_pn("0805W8F1001T5E")
    assert specs["manufacturer"] == "Royal Ohm"
    assert specs["package"] == "0805"
    assert specs["wattage"] == "1/8W"
    assert specs["tolerance"] == "1%"
    assert specs["resistance_str"] == "1k"
    assert specs["composition"] == "Thick Film"
    assert specs["special_def"] == "General"


def test_classify_molex_connector():
    m_type, conf, _ = classify_material_type("Molex KK 254 Crimp Housing 3 Circuits", "2697053")
    assert m_type == "CON"
    assert conf >= 0.70


def test_ceramic_capacitor_marks_electrolytic_attributes_as_not_applicable(rulebook):
    specs = {
        "mounting": "SMD",
        "package": "0805",
        "capacitance_str": "100nF",
        "voltage": "50V",
        "tolerance": "10%",
        "special_def": "X7R",
        "temperature": "-55°C~+125°C",
        "manufacturer": "Samsung",
    }
    mappings, desc, _ = map_specs_to_erp(rulebook, "CAP", specs)
    # Check that Load Life, Pitch, Size are NOT included for ceramic capacitor
    prop_by_name = {p.attribute_name: p for p in mappings}
    assert "Load Life" not in prop_by_name
    assert "Pitch" not in prop_by_name
    assert "Size Capacitor" not in prop_by_name
    # Description should be clean
    assert "0805 Package" in desc
    assert "100nF" in desc


def test_is_catalog_document_name():
    # True for catalog/document codes
    assert is_catalog_document_name("PYU-RC_51_ROHS_P") is True
    assert is_catalog_document_name("PYU-RC_GROUP_51_ROHS_L.pdf") is True
    assert is_catalog_document_name("RC_SERIES") is True
    assert is_catalog_document_name("RC0805") is True
    assert is_catalog_document_name("1234567") is True
    assert is_catalog_document_name("DOC-987654") is True

    # False for specific genuine MPNs
    assert is_catalog_document_name("RC0805JR-07300RL") is False
    assert is_catalog_document_name("RC0805FR-0710KL") is False
    assert is_catalog_document_name("CL21B104KBCNNNC") is False
    assert is_catalog_document_name("0805W8F1001T5E") is False
    assert is_catalog_document_name("BAT54S") is False


def test_is_automotive_qualified_prevents_false_positive_on_rc_series():
    # Yageo RC series is strictly General Purpose even if text mentions AEC-Q200 testing or AC series
    text_with_footnote = (
        "Yageo General Purpose Chip Resistors RC series.\n"
        "Note: Tested using AEC-Q200 test methods. For automotive grade, please see AC series."
    )
    assert is_automotive_qualified("RC0805JR-07300RL", text_with_footnote) is False
    assert is_automotive_qualified("PYU-RC_51_ROHS_P", text_with_footnote) is False

    # Automotive AC series should be recognized as automotive
    assert is_automotive_qualified("AC0805JR-07300RL", text_with_footnote) is True

    # Murata GRM is general purpose even if AEC-Q200 appears in notes
    assert is_automotive_qualified("GRM188R71H104KA93D", text_with_footnote) is False


def test_yageo_rc0805_exact_decoding_and_mapping(rulebook):
    specs = decode_manufacturer_pn("RC0805JR-07300RL")
    assert specs["manufacturer"] == "Yageo"
    assert specs["mounting"] == "SMD"
    assert specs["package"] == "0805"
    assert specs["tolerance"] == "5%"
    assert specs["resistance_str"] == "300R"
    assert specs["wattage"] == "1/8W"
    assert specs["composition"] == "Thick Film"
    assert specs["special_def"] == "General"

    mappings, desc, code = map_specs_to_erp(rulebook, "RES", specs)
    # Check all 8 key attributes matched
    matched_names = {p.attribute_name: p.matched_erp_value for p in mappings if p.status == "TAM EŞLEŞTİ"}
    assert "Mounting" in matched_names
    assert "Package" in matched_names and "0805" in matched_names["Package"]
    assert "Value Resistors" in matched_names and "300R" in matched_names["Value Resistors"]
    assert "Tolerance" in matched_names
    assert "Wattage" in matched_names and "1/8W" in matched_names["Wattage"]
    assert "Composition - Resistors" in matched_names and "Thick Film" in matched_names["Composition - Resistors"]
    assert "Special Definition - Resistors" in matched_names and "General" in matched_names["Special Definition - Resistors"]
    assert "Q" in matched_names and "Yageo" in matched_names["Q"]

    # Verify suggested description structure
    assert "SMD" in desc
    assert "0805 Package" in desc
    assert "300R" in desc
    assert "1/8W" in desc
    assert "Thick Film" in desc
    assert "General" in desc
    assert "Yageo" in desc
    assert "Automotive AEC-Q200" not in desc


def test_yageo_ac_automotive_decoding():
    specs = decode_manufacturer_pn("AC0805JR-07300RL")
    assert specs["special_def"] == "Automotive AEC-Q200"
    assert specs["package"] == "0805"
    assert specs["wattage"] == "1/8W"


def test_catalog_datasheet_avoids_greedy_property_fabrication(rulebook):
    pdf_bytes = _create_sample_pdf(
        "Yageo RC Series General Purpose Chip Resistors (0075 to 2512)",
        "PYU-RC_51_ROHS_P",
        [
            "Table 1: Dimensions 0075, 0100, 0201, 0402, 0603, 0805, 1206, 2512",
            "Rated power: 0201: 1/20W, 0402: 1/16W, 0603: 1/10W, 0805: 1/8W",
            "Jumper tolerance: 0% +30%",
            "Standard tolerance: 1%, 5%",
            "Tested in compliance with AEC-Q200 test specifications",
            "For automotive qualified parts refer to AC series",
        ],
    )

    result = analyze_single_datasheet(
        filename="PYU-RC_51_ROHS_P.pdf",
        file_bytes=pdf_bytes,
        rulebook=rulebook,
    )

    assert result.is_catalog is True
    # Crucial: Must NOT greedily pick 0201 as the package
    assert result.raw_specs.get("package") != "0201"
    # Crucial: Must NOT pick 1/20W as the wattage
    assert result.raw_specs.get("wattage") != "1/20W"
    # Crucial: Must NOT pick 0% +30% as tolerance
    assert result.raw_specs.get("tolerance") != "0% +30%"
    # Crucial: Must NOT be Automotive AEC-Q200
    assert result.raw_specs.get("special_def") != "Automotive AEC-Q200"
    assert result.raw_specs.get("special_def") == "General"
    assert "Katalog/Seri Dokümanı Tespit Edildi" in result.summary_text


def test_extract_part_numbers_from_text():
    text = (
        "Ordering Information:\n"
        "Example part number: RC0805JR-07300RL taped on 7-inch reel.\n"
        "Also available: CL21B104KBCNNNC and 0805W8F1001T5E.\n"
    )
    mpns = extract_part_numbers_from_text(text)
    assert "RC0805JR-07300RL" in mpns
    assert "CL21B104KBCNNNC" in mpns
    assert "0805W8F1001T5E" in mpns


def test_decode_ttaf_connector_200g1_04r(rulebook):
    specs = decode_manufacturer_pn("200G1-04R")
    assert specs["manufacturer"] == "Ttaf"
    assert specs["pitch"] == "2mm"
    assert specs["pin_number"] == "4 pin"
    assert specs["type_connector"] == "Wire to Board Konnektör"
    assert specs["pin_type"] == "Erkek"
    assert specs["mounting"] == "TH"
    assert specs["angle"] == "90°(Yatık)"

    mappings, desc, code = map_specs_to_erp(rulebook, "CON", specs)
    assert desc == "TH Wire to Board Konnektör, Erkek, 2mm, 4 pin, 90°(Yatık), Ttaf"
    assert code == "CONTADAB41-TT"
    assert len(mappings) == 7
    assert all(m.status == "TAM EŞLEŞTİ" for m in mappings)


def test_decode_royal_ohm_resistor_with_parentheses(rulebook):
    specs = decode_manufacturer_pn("0603(S/W)AJ0221T5E")
    assert specs["package"] == "0603"
    assert specs["resistance_str"] == "220R"
    assert specs["tolerance"] == "5%"
    assert specs["wattage"] == "1/10W"
    assert specs["manufacturer"] == "Royal Ohm"

    mappings, desc, code = map_specs_to_erp(rulebook, "RES", specs)
    assert "0603" in desc
    assert "220R" in desc
    assert "Thick Film" in desc
    assert "General" in desc
    assert "Royal Ohm" in desc
    assert code == "RESSS22200JC8GN-RO"
    assert len(mappings) == 8


def test_exact_pitch_and_pin_number_matching(rulebook):
    specs = {
        "mounting": "TH",
        "type_connector": "Wire to Board Konnektör",
        "pin_type": "Erkek",
        "pitch": "2mm",
        "pin_number": "4 pin",
        "angle": "90°(Yatık)",
        "manufacturer": "Ttaf",
    }
    mappings, desc, code = map_specs_to_erp(rulebook, "CON", specs)
    pitch_m = next(m for m in mappings if m.attribute_name == "Pitch Connector")
    assert pitch_m.matched_erp_value == "2mm,"
    assert pitch_m.erp_code == "B"
    pin_m = next(m for m in mappings if m.attribute_name == "Pin Number")
    assert pin_m.matched_erp_value == "4 pin,"
    assert pin_m.erp_code == "4"



