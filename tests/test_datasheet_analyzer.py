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
    extract_pdf_data,
    generate_datasheet_excel_report,
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

