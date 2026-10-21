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
    assert code == "CONTADAB0410000XX-TT"
    assert len(mappings) == 8


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
    assert pin_m.erp_code == "04"


def test_yageo_ac0603fr_07255kl_automotive_resistor(rulebook):
    pn = "AC0603FR-07255KL"
    mat_type, conf, reason = classify_material_type("AUTOMOTIVE GRADE AC series THICK FILM CHIP RESISTORS", pn)
    assert mat_type == "RES"

    pdf_bytes = _create_sample_pdf(
        "YAGEO AUTOMOTIVE GRADE AC SERIES THICK FILM CHIP RESISTORS",
        pn,
        [
            "Product specification - March 19, 2021 V.8",
            "Sizes 0201/0402/0603/0805/1206/1210/1218/2010/2512",
            "AEC-Q200 qualified",
            "AC0603: 1/10W, 75V",
            "Ordering example: AC0402FR-07100KL",
        ],
    )
    result = analyze_single_datasheet("AC_series.pdf", pdf_bytes, rulebook, part_number_override=pn)
    assert result.material_type == "RES"
    assert result.suggested_description == "SMD 0603 Package, 255K, 0.01 1/10W, Thick Film, Automotive AEC-Q200, Yageo"
    assert result.suggested_erp_code == "RESSS2K2550FC8AU-YG"
    assert len(result.properties) == 8
    assert all(p.status == "TAM EŞLEŞTİ" for p in result.properties)


def test_classify_from_cover_page_titles():
    """Verify that cover page titles accurately classify component material type."""
    test_cases = [
        ("RES", "--- PAGE 1 ---\nDATA SHEET\nTHICK FILM CHIP RESISTORS\nAUTOMOTIVE GRADE\n--- PAGE 2 ---"),
        ("CAP", "--- PAGE 1 ---\nSURFACE MOUNT MULTILAYER CERAMIC CAPACITORS\nCC Series 6.3V to 50V\n--- PAGE 2 ---"),
        ("DIO", "--- PAGE 1 ---\nSCHOTTKY BARRIER RECTIFIER\n3.0 AMPERES 40 VOLTS SS34\n--- PAGE 2 ---"),
        ("FET", "--- PAGE 1 ---\nN-Channel Enhancement Mode Power MOSFET\n30V 5.8A SI2302\n--- PAGE 2 ---"),
        ("TRA", "--- PAGE 1 ---\nNPN General Purpose Transistor\nBC847 SOT-23\n--- PAGE 2 ---"),
        ("IND", "--- PAGE 1 ---\nSMD Power Inductor\nNR4018 Series 10uH Shielded\n--- PAGE 2 ---"),
        ("ICR", "--- PAGE 1 ---\nLow Dropout Linear Regulator\nAMS1117 3.3V 1A\n--- PAGE 2 ---"),
        ("ICS", "--- PAGE 1 ---\nStep-Down Switching Regulator\nLM2596 3A Buck Converter\n--- PAGE 2 ---"),
        ("MCU", "--- PAGE 1 ---\n32-Bit ARM Cortex-M0 Microcontroller\nSTM32F030F4P6\n--- PAGE 2 ---"),
        ("OPT", "--- PAGE 1 ---\nPHOTOCOUPLER\nPC817 Phototransistor Optocoupler\n--- PAGE 2 ---"),
        ("REL", "--- PAGE 1 ---\nMiniature Power Relay\n12VDC 10A 250VAC HF32F\n--- PAGE 2 ---"),
        ("FUS", "--- PAGE 1 ---\nSurface Mount Resettable PTC Fuse\n1206 MF-MSMF series\n--- PAGE 2 ---"),
        ("VAR", "--- PAGE 1 ---\nRadial Lead Metal Oxide Varistor\n14D series 275VAC 430V\n--- PAGE 2 ---"),
        ("TER", "--- PAGE 1 ---\nNTC Thermistor Temperature Sensor\n10K ohm 3950K\n--- PAGE 2 ---"),
        ("CRY", "--- PAGE 1 ---\nQuartz Crystal Unit\n32.768kHz 12.5pF SMD 3.2x1.5mm\n--- PAGE 2 ---"),
        ("SWT", "--- PAGE 1 ---\nTactile Switch\n6x6mm SMD Push Button\n--- PAGE 2 ---"),
        ("FER", "--- PAGE 1 ---\nFerrite Chip Bead\n0805 120 ohm at 100MHz\n--- PAGE 2 ---"),
        ("POT", "--- PAGE 1 ---\nTrimmer Potentiometer\n3296 Series 10K Multi Turn\n--- PAGE 2 ---"),
        ("BUZ", "--- PAGE 1 ---\nPiezo Buzzer\n12V 85dB Resonant Frequency 2.4kHz\n--- PAGE 2 ---"),
        ("TRF", "--- PAGE 1 ---\nCurrent Transformer\n1:1000 50A Isolation\n--- PAGE 2 ---"),
        ("AMP", "--- PAGE 1 ---\nOperational Amplifier\nLM358 Dual Op Amp\n--- PAGE 2 ---"),
    ]
    for expected_type, sample_text in test_cases:
        mtype, conf, reason = classify_material_type(sample_text)
        assert mtype == expected_type, f"Failed for {expected_type}: got {mtype} ({reason})"
        assert conf >= 0.95


def test_inductor_spec_extraction_and_mapping(rulebook):
    """Verify inductance extraction and mapping to rules.xlsx."""
    sample_text = (
        "--- PAGE 1 ---\n"
        "SMD Power Inductor NR4018 Series\n"
        "Dimensions: 4.0x4.0x1.8mm\n"
        "Inductance: 10uH\n"
        "Mounting: Surface Mount SMD\n"
        "--- PAGE 2 ---\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "NR4018-100M")
    assert mat_type == "IND"

    specs = extract_component_specs("IND", sample_text, "NR4018-100M")
    assert specs.get("mounting") == "SMD"
    assert specs.get("inductance_str") == "10uH"

    mappings, desc, code = map_specs_to_erp(rulebook, "IND", specs)
    # Value Inductor attribute should match 10uH
    val_m = next(m for m in mappings if "value" in m.attribute_name.casefold())
    assert val_m.matched_erp_value == "10uH,"
    assert val_m.status in ("TAM EŞLEŞTİ", "YAKIN EŞLEŞTİ")


def test_te_connectivity_antenna_2108857_1(rulebook):
    """Verify that TE Connectivity antenna 2108857-1 is classified as ANT and maps to ANTC3-TC."""
    sample_text = (
        "--- PAGE 1 ---\n"
        "Wi-Fi 6/ 6E TRIPLE BAND EMBEDDED ANTENNA\n"
        "Part Numbers: 2108857-1, 2108857-2\n"
        "TE Connectivity\n"
        "FEATURES & BENEFITS\n"
        "• WiFi 6, also including WiFi 6E new 6GHz band, Bluetooth, ZigBee\n"
        "• Adhesive backing on the PCB simplifies mounting within the device\n"
        "• Global Frequency coverage\n"
        "• Wide selection of micro-coaxial cable lengths\n"
        "SPECIFICATIONS\n"
        "Frequency Range (MHz) 2400-2500 5150-5875 5925-7125\n"
        "Mounting Adhesive Tape\n"
        "Cable (Coaxial) 0.81mm and 1.37mm Dia.\n"
        "--- PAGE 2 ---\n"
    )
    pn = "2108857-1"
    mat_type, conf, reason = classify_material_type(sample_text, pn)
    assert mat_type == "ANT"
    assert conf >= 0.95

    specs = extract_component_specs("ANT", sample_text, pn)
    assert specs.get("mounting") == "TH"
    assert specs.get("type_antenna") == "RF"
    assert specs.get("manufacturer") == "TE Connectivity"
    assert specs.get("manufacturer_code") == "TC"

    mappings, desc, code = map_specs_to_erp(rulebook, "ANT", specs)
    assert desc == "TH RF, TE Connectivity"
    assert code == "ANTT30000XX-TC"
    assert len(mappings) == 4

    # With known sequential counter from DEKA ERP
    specs["serial_counter"] = "000023"
    _, _, code_known = map_specs_to_erp(rulebook, "ANT", specs)
    assert code_known == "ANTT3000023-TC"

    from bom_enricher.datasheet_analyzer import build_rich_description
    specs["cable_length"] = "50mm"
    specs["mating_connector"] = "MHF"
    rich_desc = build_rich_description("ANT", specs, text=sample_text)
    assert rich_desc == "RF Board Antenna, WiFi, 50mm MHF Cable, TE Connectivity"


def test_pdfplumber_2d_table_extraction_antenna():
    from bom_enricher.datasheet_analyzer import match_part_number_in_tables, extract_pdf_tables
    from pathlib import Path
    pdf_path = Path(r"C:\Users\sudem\.gemini\antigravity\brain\8be84be6-1095-43d5-b848-ce9fe57a8347\.user_uploaded\media_1792147543521.pdf")
    if pdf_path.exists():
        pdf_bytes = pdf_path.read_bytes()
        tables = extract_pdf_tables(pdf_bytes)
        assert len(tables) >= 5

        # Variant 1: 50mm MHF
        res1 = match_part_number_in_tables(pdf_bytes, "2108857-1")
        assert res1.get("cable_length") == "50mm"
        assert res1.get("mating_connector") == "MHF"

        # Variant 2: 100mm MHF
        res2 = match_part_number_in_tables(pdf_bytes, "2108857-2")
        assert res2.get("cable_length") == "100mm"
        assert res2.get("mating_connector") == "MHF"

        # Variant 5: 50mm MHF4
        res5 = match_part_number_in_tables(pdf_bytes, "2108857-5")
        assert res5.get("cable_length") == "50mm"
        assert res5.get("mating_connector") == "MHF4"


def test_pdfplumber_graceful_fallback():
    from bom_enricher.datasheet_analyzer import match_part_number_in_tables, extract_pdf_tables
    # Empty bytes
    assert extract_pdf_tables(b"") == []
    assert match_part_number_in_tables(b"", "2108857-1") == {}
    assert match_part_number_in_tables(None, "2108857-1") == {}
    # Non-PDF random bytes
    assert extract_pdf_tables(b"random-bytes-not-a-pdf") == []
    assert match_part_number_in_tables(b"random-bytes-not-a-pdf", "2108857-1") == {}


def test_no_false_positive_te_connectivity():
    from bom_enricher.datasheet_analyzer import extract_component_specs
    text = "Operating Temperature: -40 to +85 C, rated voltage 5V, white housing, byte format, test system."
    specs = extract_component_specs(text, "GENERIC-123", "RES")
    assert specs.get("manufacturer") != "TE Connectivity"
    assert specs.get("manufacturer_code") != "TC"


def test_buzzer_spec_extraction_fuet_and_huaneng(rulebook):
    folder = Path("C:/Users/sudem/Desktop/datasheets")
    fuet_path = folder / "FUET-5020.pdf"
    if fuet_path.exists():
        res = analyze_single_datasheet(fuet_path.name, fuet_path.read_bytes(), rulebook)
        assert res.material_type == "BUZ"
        assert res.raw_specs.get("buzzer_voltage") == "3V"
        assert res.raw_specs.get("buzzer_frequency") == "4kHz"
        assert res.raw_specs.get("sound_level") == "75dB"
        assert res.raw_specs.get("manufacturer") == "FUET"
        assert "BUZSA33" in res.suggested_erp_code
        assert "-FT" in res.suggested_erp_code

    hnb_path = folder / "HNB09A05.pdf"
    if hnb_path.exists():
        res = analyze_single_datasheet(hnb_path.name, hnb_path.read_bytes(), rulebook)
        assert res.material_type == "BUZ"
        assert res.raw_specs.get("buzzer_voltage") == "5V"
        assert res.raw_specs.get("buzzer_frequency") == "2.7kHz"
        assert res.raw_specs.get("sound_level") == "85dB"
        assert res.raw_specs.get("manufacturer") == "Huaneng"
        assert "-HUA" in res.suggested_erp_code


def test_diode_m7_extraction(rulebook):
    folder = Path("C:/Users/sudem/Desktop/datasheets")
    m7_path = folder / "M7_HT.pdf"
    if m7_path.exists():
        res = analyze_single_datasheet(m7_path.name, m7_path.read_bytes(), rulebook)
        assert res.material_type == "DIO"
        assert res.raw_specs.get("package") == "SMA"
        assert res.raw_specs.get("voltage") == "1000V"
        assert res.raw_specs.get("current") == "1A"
        assert res.raw_specs.get("manufacturer") == "Hottech"
        assert "DIO" in res.suggested_erp_code
        assert "-HO" in res.suggested_erp_code


def test_varistor_b722_extraction(rulebook):
    folder = Path("C:/Users/sudem/Desktop/datasheets")
    var_path = folder / "B72207S0271K111.pdf"
    if var_path.exists():
        res = analyze_single_datasheet(var_path.name, var_path.read_bytes(), rulebook)
        assert res.material_type == "VAR"
        assert res.raw_specs.get("varistor_size") == "7mm"
        assert res.raw_specs.get("manufacturer") in ("Epcos", "TDK")
        assert res.raw_specs.get("voltage_ac") == "275VAC"
        assert res.inferred_mounting == "TH"


def test_battery_holder_kls_extraction(rulebook):
    res = analyze_single_datasheet("L-KLS5-SP-073-R.pdf", b"", rulebook)
    assert res.material_type == "BAH"
    assert res.raw_specs.get("manufacturer") == "Kls"


def test_degson_terminal_block_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, classify_material_type
    pn = "DG8HS-1.2-12P-17-100SH/AH"
    specs = decode_manufacturer_pn(pn)
    assert specs.get("manufacturer") == "Degson"
    assert specs.get("pin_number") == "12 pin"
    assert specs.get("pitch") == "8.2mm"
    assert specs.get("pin_type") == "Dişi"
    assert specs.get("type_connector") == "Wire to Wire Klemens"
    assert specs.get("mounting") == "Cable"

    mat_type, conf, reason = classify_material_type("", pn, "")
    assert mat_type == "CON"


def test_empty_text_does_not_fallback_to_capacitor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    mat_type, conf, reason = classify_material_type("", "UNKNOWN-PART-XYZ", "")
    assert mat_type == "BİLİNMİYOR"
    assert conf == 0.0


def test_analyze_datasheet_cad_zero_text_uses_vision_fusion():
    from unittest.mock import patch
    import io
    from reportlab.pdfgen import canvas
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    # Empty canvas simulating vector CAD drawing without extractable text
    c.save()
    cad_pdf_bytes = buf.getvalue()

    mock_vlm_output = {
        "material_type": "CON",
        "confidence": 0.90,
        "manufacturer": "Degson",
        "specs": {
            "type_connector": "Wire to Wire Klemens",
            "pitch": "8.2mm",
            "pin_number": "12 pin",
            "pin_type": "Dişi",
            "mounting": "Cable",
            "angle": "180°(Dik)",
        },
    }

    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images", return_value=mock_vlm_output):
        result = analyze_datasheet(cad_pdf_bytes, part_number="DG8HS-1.2-12P-17-100SH/AH")
        assert result["material_type"] == "CON"
        assert result["specs"].get("pin_number") == "12 pin"
        assert result["specs"].get("pitch") == "8.2mm"
        # Deterministic ERP code must be generated matching CON rules
        assert result.get("erp_code") in ("CONCAIB4120000XX-DE", "CONCAIB41220000XX-DE")
        assert result.get("analysis_method") in ("hybrid_vision", "fast_path")


def test_analyze_datasheet_fast_path_when_pn_has_core_specs(rulebook):
    from unittest.mock import patch
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    pdf_bytes = _create_sample_pdf(
        "YAGEO CHIP RESISTORS RC SERIES",
        "RC0805FR-0710KL",
        ["SMD Resistor 10k 1% 0805", "Resistance: 10k", "Tolerance: 1%"],
    )
    # Even if vision client is available, complete text + clean PN doesn't trigger vision
    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images") as mock_extract:
        result = analyze_datasheet(pdf_bytes, part_number="RC0805FR-0710KL", rulebook=rulebook)
        assert result["material_type"] == "RES"
        assert result["specs"].get("resistance_str") == "10k"
        assert result["analysis_method"] == "fast_path"
        mock_extract.assert_not_called()


def test_analyze_datasheet_deterministic_text_method(rulebook):
    from unittest.mock import patch
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    # Custom PN that cannot be decoded by decode_manufacturer_pn
    pdf_bytes = _create_sample_pdf(
        "GENERAL PURPOSE RECTIFIER DIODE",
        "GENERIC-RECTIFIER-XYZ",
        ["Surface Mount Rectifier", "Voltage: 400V", "Current: 1A", "Package: SMA"],
    )
    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images") as mock_extract:
        result = analyze_datasheet(pdf_bytes, part_number="GENERIC-RECTIFIER-XYZ", rulebook=rulebook)
        assert result["material_type"] == "DIO"
        assert result["analysis_method"] == "deterministic_text"
        mock_extract.assert_not_called()


def test_analyze_datasheet_vision_unavailable_graceful_fallback(rulebook):
    from unittest.mock import patch
    import io
    from reportlab.pdfgen import canvas
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.save()
    cad_pdf_bytes = buf.getvalue()

    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=False):
        result = analyze_datasheet(cad_pdf_bytes, part_number="DG8HS-1.2-12P-17-100SH/AH", rulebook=rulebook)
        assert result["material_type"] == "CON"
        assert result["specs"].get("pin_number") == "12 pin"
        # Fast path was used since vision was unavailable
        assert result["analysis_method"] == "fast_path"


def test_analyze_datasheet_adopts_vlm_material_type_for_unknown():
    from unittest.mock import patch
    import io
    from reportlab.pdfgen import canvas
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.save()
    empty_pdf = buf.getvalue()

    mock_vlm_output = {
        "material_type": "DIO",
        "confidence": 0.92,
        "manufacturer": "Vishay",
        "specs": {
            "package": "SMA",
            "voltage": "600V",
            "current": "2A",
        },
    }

    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images", return_value=mock_vlm_output):
        result = analyze_datasheet(empty_pdf, part_number="UNKNOWN-CUSTOM-PN-99")
        assert result["material_type"] == "DIO"
        assert result["confidence"] == 0.92
        assert result["specs"].get("voltage") == "600V"
        assert result["analysis_method"] == "hybrid_vision"


def test_analyze_datasheet_normalizes_vlm_material_type_synonyms():
    from unittest.mock import patch
    import io
    from reportlab.pdfgen import canvas
    from bom_enricher.datasheet_analyzer import analyze_datasheet, NORMALIZE_VLM_TYPE

    # Direct dictionary verification
    assert NORMALIZE_VLM_TYPE["DIODE"] == "DIO"
    assert NORMALIZE_VLM_TYPE["CONNECTOR"] == "CON"
    assert NORMALIZE_VLM_TYPE["TRANSISTOR"] == "TRA"
    assert NORMALIZE_VLM_TYPE["RESISTOR"] == "RES"

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.save()
    empty_pdf = buf.getvalue()

    # Case 1: VLM outputs "DIODE" -> normalized to "DIO"
    vlm_diode_output = {
        "material_type": "DIODE",
        "confidence": 0.88,
        "specs": {"voltage": "1000V", "current": "1A"},
    }
    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images", return_value=vlm_diode_output):
        result_diode = analyze_datasheet(empty_pdf, part_number="UNKNOWN-1N4007-CUSTOM")
        assert result_diode["material_type"] == "DIO"
        assert result_diode["material_type_name"] == "Diyot (DIO)"

    # Case 2: VLM outputs "connector" (lowercase) -> normalized to "CON"
    vlm_conn_output = {
        "material_type": "connector",
        "confidence": 0.91,
        "specs": {"pin_number": "4 pin", "pitch": "2.54mm"},
    }
    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images", return_value=vlm_conn_output):
        result_conn = analyze_datasheet(empty_pdf, part_number="UNKNOWN-HDR-CUSTOM")
        assert result_conn["material_type"] == "CON"
        assert result_conn["material_type_name"] == "Konnektör (CON)"


@pytest.mark.slow
def test_real_degson_cad_datasheet_e2e(rulebook):
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    pdf_path = Path(r"C:\Users\sudem\Desktop\datasheets\DG8-10-12-15HS-1.2.pdf")
    if not pdf_path.exists():
        pytest.skip("Local Degson PDF not found on Desktop")

    with open(pdf_path, "rb") as f:
        file_bytes = f.read()

    pn = "DG8HS-1.2-12P-17-100SH/AH"
    result = analyze_datasheet(file_bytes, part_number=pn, rulebook=rulebook)

    # Must be classified as CON
    assert result["material_type"] == "CON"
    assert result["specs"].get("pin_number") == "12 pin"
    assert result["specs"].get("pitch") == "8.2mm"
    assert result["specs"].get("type_connector") == "Wire to Wire Klemens"
    assert result["specs"].get("pin_type") == "Dişi"
    assert result["specs"].get("mounting") == "Cable"

    # Deterministic ERP code verification
    assert result.get("erp_code") == "CONCAIB4120000XX-DE"


@pytest.mark.slow
def test_real_cad_pdf_without_pn_does_not_fallback_to_cap():
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    pdf_path = Path(r"C:\Users\sudem\Desktop\datasheets\DG8-10-12-15HS-1.2.pdf")
    if not pdf_path.exists():
        pytest.skip("Local Degson PDF not found on Desktop")

    with open(pdf_path, "rb") as f:
        file_bytes = f.read()

    result = analyze_datasheet(file_bytes, part_number=None, filename="datasheet.pdf")
    assert result["material_type"] in ("CON", "BİLİNMİYOR")
    assert result["material_type"] != "CAP"


def test_crystal_with_esr_not_misclassified_as_resistor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SURFACE MOUNT MICROPROCESSOR CRYSTAL\n"
        "RH100-8.000-16-3030-TR-NS1\n"
        "RALTRON ELECTRONICS\n"
        "NOMINAL FREQUENCY 8.000 MHz\n"
        "LOAD CAPACITANCE 16 pF\n"
        "EQUIVALENT SERIES RESISTANCE 400 ohm max\n"
        "SHUNT CAPACITANCE 2 pF max\n"
        "INSULATION RESISTANCE 500 Mohm min\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "RH100-8.000-16-3030-TR-NS1", "")
    assert mat_type == "CRY"
    assert conf >= 0.95


def test_inductor_with_dcr_not_misclassified_as_resistor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SMD POWER INDUCTOR\n"
        "CD54-100MC\n"
        "INDUCTANCE 10 uH\n"
        "DC RESISTANCE 0.15 OHM MAX\n"
        "RATED CURRENT 1.44 A\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "CD54-100MC", "")
    assert mat_type == "IND"
    assert conf >= 0.95


def test_diode_with_junction_capacitance_not_misclassified_as_capacitor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SURFACE MOUNT SCHOTTKY BARRIER RECTIFIER\n"
        "SS34\n"
        "FORWARD VOLTAGE 0.5 V\n"
        "REVERSE LEAKAGE CURRENT 0.5 mA\n"
        "JUNCTION CAPACITANCE 150 pF\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "SS34", "")
    assert mat_type == "DIO"
    assert conf >= 0.95


def test_parasitic_filtering_pure_keyword_scoring():
    from bom_enricher.datasheet_analyzer import classify_material_type
    # Crystal with generic PN and no title header
    text_cry = (
        "SPECIFICATION\n"
        "CUSTOM-CRY-999\n"
        "Crystal Unit\n"
        "Load Capacitance 18 pF\n"
        "Equivalent Series Resistance 50 ohm max\n"
        "Shunt Capacitance 7 pF max\n"
    )
    mat_type, conf, reason = classify_material_type(text_cry, "CUSTOM-CRY-999", "")
    assert mat_type == "CRY"
    assert conf >= 0.70

    # Inductor with generic PN and no title header
    text_ind = (
        "SPECIFICATION\n"
        "CUSTOM-IND-999\n"
        "Choke Coil Inductance 22 uH\n"
        "DC Resistance 0.05 ohm\n"
    )
    mat_type, conf, reason = classify_material_type(text_ind, "CUSTOM-IND-999", "")
    assert mat_type == "IND"
    assert conf >= 0.70

    # Diode with generic PN and no title header
    text_dio = (
        "SPECIFICATION\n"
        "CUSTOM-DIO-999\n"
        "Rectifier Diode\n"
        "Forward Voltage 0.7 V\n"
        "Junction Capacitance 50 pF\n"
    )
    mat_type, conf, reason = classify_material_type(text_dio, "CUSTOM-DIO-999", "")
    assert mat_type == "DIO"
    assert conf >= 0.70


def test_crystal_spec_extraction_and_rich_description():
    from bom_enricher.datasheet_analyzer import extract_component_specs, build_rich_description
    sample_text = (
        "SURFACE MOUNT MICROPROCESSOR CRYSTAL\n"
        "RH100-8.000-16-3030-TR-NS1\n"
        "RALTRON ELECTRONICS\n"
        "NOMINAL FREQUENCY 8.000 MHz\n"
        "LOAD CAPACITANCE 16 pF\n"
        "EQUIVALENT SERIES RESISTANCE 400 ohm max\n"
        "SHUNT CAPACITANCE 2 pF max\n"
        "4x1.0 4x1.4 SOLDERING PATTERN TOP VIEW PIN CONNECTION GND 4 3 2 GND 1 3.2x2.5\n"
    )
    specs = extract_component_specs(sample_text, "CRY", "RH100-8.000-16-3030-TR-NS1")
    assert specs.get("frequency_str") == "8MHz"
    assert specs.get("load_capacitance") == "16 pF"
    assert specs.get("esr") in ("400Ohm", "400 ohm")
    assert specs.get("mounting") == "SMD"
    assert specs.get("package") == "4-SMD"
    assert specs.get("manufacturer") == "Raltron"

    desc = build_rich_description("CRY", specs, sample_text)
    assert desc == "SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron"


def test_real_raltron_crystal_datasheet_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Synthetic multi-attribute Raltron crystal datasheet
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "SURFACE MOUNT MICROPROCESSOR CRYSTAL")
    c.drawString(100, 730, "RALTRON ELECTRONICS - RH100-8.000-16-3030-TR-NS1")
    c.drawString(100, 710, "NOMINAL FREQUENCY: 8.000 MHz")
    c.drawString(100, 690, "LOAD CAPACITANCE: 16 pF")
    c.drawString(100, 670, "EQUIVALENT SERIES RESISTANCE: 400 ohm max")
    c.drawString(100, 650, "MECHANICAL SPECIFICATION: 3.2 x 2.5 mm, 4 PADS, TOP VIEW PIN CONNECTION GND 4, 3, 2 GND, 1")
    c.save()
    pdf_bytes = buf.getvalue()

    pn = "RH100-8.000-16-3030-TR-NS1"
    result = analyze_datasheet(pdf_bytes, part_number=pn)

    assert result["material_type"] == "CRY"
    assert result["specs"].get("frequency_str") == "8MHz"
    assert result["specs"].get("load_capacitance") == "16 pF"
    assert result["specs"].get("esr") == "400Ohm"
    assert result["specs"].get("package") == "4-SMD"
    assert result["specs"].get("mounting") == "SMD"
    assert result["specs"].get("manufacturer") == "Raltron"

    assert result["rich_description"] == "SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron"
    assert result["erp_code"].startswith("CRYSA2")
    assert result["erp_code"].endswith("-RA")
    assert result["erp_code"] == "CRYSA20000XX-RA"

    # Also verify when serial_counter="076" is passed
    result_with_counter = analyze_datasheet(pdf_bytes, part_number=pn, serial_counter="076")
    assert result_with_counter["erp_code"] == "CRYSA2076-RA"


def test_dale_resistor_not_misattributed_to_raltron():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, extract_component_specs
    specs = decode_manufacturer_pn("RH-50")
    assert specs.get("manufacturer") != "Raltron"

    specs_ext = extract_component_specs("POWER RESISTOR RH-50 50W", "RES", "RH-50")
    assert specs_ext.get("manufacturer") != "Raltron"


def test_standalone_esr_caught_by_parasitic_regex_without_inflating_res():
    import re
    from bom_enricher.datasheet_analyzer import PARASITIC_RESISTANCE_RE, classify_material_type

    # 1. Standalone ESR/DCR patterns (without the word 'resistance') match PARASITIC_RESISTANCE_RE
    assert re.search(PARASITIC_RESISTANCE_RE, "ESR: 400 ohm max", re.IGNORECASE) is not None
    assert re.search(PARASITIC_RESISTANCE_RE, "DCR: 0.15 ohm max", re.IGNORECASE) is not None
    assert re.search(PARASITIC_RESISTANCE_RE, "ESR = 25", re.IGNORECASE) is not None
    assert re.search(PARASITIC_RESISTANCE_RE, "esr(typ) 10", re.IGNORECASE) is not None
    assert re.search(PARASITIC_RESISTANCE_RE, "DCR <= 0.05", re.IGNORECASE) is not None

    # 2. Text with standalone ESR: 400 ohm max (without the word 'resistance')
    # and an unrelated resistance mention (e.g. Solder Resistance 260 C)
    # when non_resistor_dominant is NOT set (e.g. for a varistor or capacitor).
    # Standalone ESR must trigger parasitic filtering, dropping bare 'resistance' from RES scoring,
    # so RES does not get inflated (+48 pts) and does not falsely beat the actual component type.
    text_var = (
        "DEVICE SPECIFICATION\n"
        "PART NO: VAR-CUSTOM-99\n"
        "Clamping Voltage: 50V\n"
        "ESR: 400 ohm max\n"
        "Solder Resistance: 260C\n"
    )
    mat_type, conf, reason = classify_material_type(text_var, "VAR-CUSTOM-99", "")
    assert mat_type == "VAR"
    assert mat_type != "RES"

    # Also test capacitor where non_resistor_dominant is not set:
    text_cap = (
        "SPECIFICATION\n"
        "PART NO: CUST-CAP-01\n"
        "Capacitance: 100 uF\n"
        "ESR: 400 ohm max\n"
        "Solder Resistance: 260C\n"
    )
    mat_type, conf, reason = classify_material_type(text_cap, "CUST-CAP-01", "")
    assert mat_type == "CAP"
    assert mat_type != "RES"


def test_dsp_classification_from_title_and_pn():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "NINGBO FORYARD OPTOELECTRONICS CO.,LTD\n"
        "PRODUCT SPECIFICATION\n"
        "Model No.: FYQ-5643C/DX-XX\n"
        "Descriptions & Features:\n"
        "0.56Inch (14.20mm) digit height Four Digits Display\n"
        "Case mold type.\n"
        "Easy mounting on P.C. board or socket.\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "FYQ-5643CUW-21", "")
    assert mat_type == "DSP"
    assert conf >= 0.95

    # Test PN prefix high-confidence classification without text
    mat_type_pn, conf_pn, _ = classify_material_type("", "FYQ-5643CUW-21", "")
    assert mat_type_pn == "DSP"
    assert conf_pn >= 0.95


def test_foryard_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn
    # Common Cathode White 4-digit
    specs = decode_manufacturer_pn("FYQ-5643CUW-21")
    assert specs.get("manufacturer") == "Foryard"
    assert specs.get("mounting") == "TH"
    assert specs.get("type_display") == "7 Segment Display"
    assert specs.get("color") == "Beyaz"
    assert specs.get("polarity") == "COM-"
    assert specs.get("digits") == "4 Digit"

    # Common Anode Red 1-digit
    specs_anode = decode_manufacturer_pn("FYS-5611DUR-11")
    assert specs_anode.get("polarity") == "COM+"
    assert specs_anode.get("color") == "Kırmızı"
    assert specs_anode.get("digits") == "1 Digit"

    # 2-Digit Green Common Cathode
    specs_dual = decode_manufacturer_pn("FYD-5621CUG-21")
    assert specs_dual.get("digits") == "2 Digit"
    assert specs_dual.get("polarity") == "COM-"
    assert specs_dual.get("color") == "Yeşil"

    # 3-Digit Blue Common Anode
    specs_triple = decode_manufacturer_pn("FYT-5631DUB-11")
    assert specs_triple.get("digits") == "3 Digit"
    assert specs_triple.get("polarity") == "COM+"
    assert specs_triple.get("color") == "Mavi"

    # Yellow color
    specs_yellow = decode_manufacturer_pn("FYD-5621DUY-11")
    assert specs_yellow.get("color") == "Sarı"


def test_universal_rule_matcher_dsp_attributes():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp
    rulebook = load_rulebook("data/rules.xlsx")
    specs = {
        "mounting": "TH",
        "type_display": "7 Segment Display",
        "color": "Beyaz",
        "special_def_display": "4 Digit, COM-",
        "serial_counter": "0028",
        "manufacturer": "Foryard",
    }
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "DSP", specs)
    assert erp_code == "DSPTE4BH0028-FO"
    attr_map = {m.attribute_name: (m.matched_erp_value, m.erp_code) for m in mappings}
    assert attr_map["Mounting"][1] == "T"
    assert attr_map["Type Display"][1] == "E"
    assert attr_map["Color Display LED"][1] == "4"
    sp_def = attr_map.get("Special Definition - Display") or attr_map.get("Special Definition - Display ")
    assert sp_def is not None and sp_def[1] == "BH"
    assert attr_map["Serial Display Number"][1] == "0028"
    assert attr_map["Q"][1] == "FO"


def test_dsp_rich_description():
    from bom_enricher.datasheet_analyzer import build_rich_description
    specs = {
        "mounting": "TH",
        "type_display": "7 Segment Display",
        "color": "Beyaz",
        "special_def_display": "4 Digit, COM-",
        "manufacturer": "Foryard",
    }
    desc = build_rich_description("DSP", specs, "")
    assert desc == "TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard"


def test_dsp_extract_component_specs():
    from bom_enricher.datasheet_analyzer import extract_component_specs
    text = (
        "NINGBO FORYARD OPTOELECTRONICS CO.,LTD\n"
        "Model No.: FYQ-5643CUW-21\n"
        "0.56Inch Four Digits Display\n"
        "Ultra White Common Cathode\n"
        "Through hole type\n"
    )
    specs = extract_component_specs("DSP", text, "FYQ-5643CUW-21")
    assert specs["mounting"] == "TH"
    assert specs["type_display"] == "7 Segment Display"
    assert specs["color"] == "Beyaz"
    assert specs["digits"] == "4 Digit"
    assert specs["polarity"] == "COM-"
    assert specs["special_def_display"] == "4 Digit, COM-"
    assert specs["manufacturer"] == "Foryard"


def test_dsp_multilingual_semantic_matching():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp
    rulebook = load_rulebook("data/rules.xlsx")
    # Provide English / alternate synonym terms
    specs = {
        "mounting": "TH",
        "type_display": "7-Segment Display",
        "color": "White",  # Should match Beyaz via MULTILINGUAL_SYNONYMS
        "special_def_display": "4 Digit, Common Cathode",  # Should match 4 Digit, COM- via synonyms
        "serial_counter": "0028",
        "manufacturer": "Foryard",
    }
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "DSP", specs)
    assert erp_code == "DSPTE4BH0028-FO"
    attr_map = {m.attribute_name: (m.matched_erp_value, m.erp_code) for m in mappings}
    assert attr_map["Color Display LED"][1] == "4"
    sp_def = attr_map.get("Special Definition - Display") or attr_map.get("Special Definition - Display ")
    assert sp_def is not None and sp_def[1] == "BH"


def test_match_spec_to_entries_partial_compound_no_false_match():
    from bom_enricher.datasheet_analyzer import _match_spec_to_entries, RuleEntry
    e1 = RuleEntry("DSP", "ATTR", "Special Definition - Display", "BG", "4 Digit, COM+,", "TR", "4", "", "")
    e2 = RuleEntry("DSP", "ATTR", "Special Definition - Display", "BH", "4 Digit, COM-,", "TR", "4", "", "")
    entries = (e1, e2)

    # 1. Partial value (digit only) must NOT falsely match first compound entry
    matched, status = _match_spec_to_entries("4 Digit", entries)
    assert matched is None
    assert status == "BULUNAMADI"

    # 2. Partial value (polarity only) must NOT falsely match
    matched, status = _match_spec_to_entries("Common Cathode", entries)
    assert matched is None
    assert status == "BULUNAMADI"

    # 3. Full value with polarity matches correct polarity entry
    matched_neg, status_neg = _match_spec_to_entries("4 Digit, COM-", entries)
    assert matched_neg is not None
    assert matched_neg.value == "BH"
    assert status_neg == "TAM EŞLEŞTİ"

    # 4. Synonym expansion for full multi-dimensional value
    matched_syn_neg, _ = _match_spec_to_entries("4 Digit, Common Cathode", entries)
    assert matched_syn_neg is not None
    assert matched_syn_neg.value == "BH"

    matched_syn_pos, _ = _match_spec_to_entries("4 Digit, Common Anode", entries)
    assert matched_syn_pos is not None
    assert matched_syn_pos.value == "BG"


def test_real_foryard_display_datasheet_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Synthetic Foryard 4-Digit Display datasheet
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "NINGBO FORYARD OPTOELECTRONICS CO.,LTD")
    c.drawString(100, 730, "PRODUCT SPECIFICATION - Model No.: FYQ-5643C/DX-XX")
    c.drawString(100, 710, "0.56Inch (14.20mm) digit height Four Digits Display")
    c.drawString(100, 690, "Descriptions & Features: Case mold type. Easy mounting on P.C. board or socket.")
    c.drawString(100, 670, "Mechanical Dimensions: All pins are Φ0.51mm, 2.54*6=15.24")
    c.drawString(100, 650, "Internal Circuit Diagrams: FYQ-5643Cx Common Cathode, FYQ-5643Dx Common Anode")
    c.drawString(100, 630, "Electrical-Optical: UW Ultra White InGaN CCT:9500K Forward Voltage 2.80V")
    c.save()
    pdf_bytes = buf.getvalue()

    pn = "FYQ-5643CUW-21"

    # 1. Default counter test
    result = analyze_datasheet(pdf_bytes, part_number=pn)
    assert result["material_type"] == "DSP"
    assert result["specs"].get("mounting") == "TH"
    assert result["specs"].get("type_display") == "7 Segment Display"
    assert result["specs"].get("color") == "Beyaz"
    assert result["specs"].get("polarity") == "COM-"
    assert result["specs"].get("digits") == "4 Digit"
    assert result["specs"].get("manufacturer") == "Foryard"
    assert result["rich_description"] == "TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard"
    assert result["erp_code"] == "DSPTE4BH0000XX-FO"

    # 2. Injected internal counter test (0028)
    result_counter = analyze_datasheet(pdf_bytes, part_number=pn, serial_counter="0028")
    assert result_counter["erp_code"] == "DSPTE4BH0028-FO"


def test_panjit_schottky_diode_ss0520_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn
    # SS0520: 0.5A (500mA), 20V Schottky
    specs = decode_manufacturer_pn("SS0520")
    assert specs.get("type_diode") == "Schottky"
    assert specs.get("mounting") == "SMD"
    assert specs.get("voltage") == "20V"
    assert specs.get("current") == "500mA"
    assert specs.get("value_diode") == "SS0520"

    # SS0540: 0.5A (500mA), 40V
    specs40 = decode_manufacturer_pn("SS0540")
    assert specs40.get("voltage") == "40V"
    assert specs40.get("current") == "500mA"

    # SS34: 3A, 40V
    specs34 = decode_manufacturer_pn("SS34")
    assert specs34.get("voltage") == "40V"
    assert specs34.get("current") == "3A"


def test_dio_map_specs_to_erp_and_rich_description():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp, build_rich_description
    rulebook = load_rulebook("data/rules.xlsx")
    specs = {
        "mounting": "SMD",
        "package": "SOD123",
        "type_diode": "Schottky",
        "voltage": "20V",
        "current": "500mA",
        "value_diode": "SS0520",
        "manufacturer": "Panjit",
        "manufacturer_code": "PA",
    }
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "DIO", specs)
    assert erp_code == "DIOSZ41S1000-PA"
    rich_desc = build_rich_description("DIO", specs, "")
    assert rich_desc == "SMD SOD123, Schottky, 20V, 500mA,, SS0520,  Panjit"


def test_real_panjit_ss0520_datasheet_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Synthetic Panjit SS0520 datasheet
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "PANJIT SEMICONDUCTOR")
    c.drawString(100, 730, "SS0520~SS0540 SURFACE MOUNT SCHOTTKY BARRIER")
    c.drawString(100, 710, "VOLTAGE 20 to 40 Volt CURRENT 0.5 Ampere")
    c.drawString(100, 690, "Case: SOD-123, Plastic")
    c.drawString(100, 670, "PARAMETER: Maximum Recurrent Peak Reverse Voltage VRRM: SS0520=20V, SS0530=30V, SS0540=40V")
    c.drawString(100, 650, "Maximum Average Forward Current at TA=75C IAV: 0.5 A")
    c.drawString(100, 630, "Panjit International Inc.")
    c.save()
    pdf_bytes = buf.getvalue()

    result = analyze_datasheet(pdf_bytes, part_number="SS0520")
    assert result["material_type"] == "DIO"
    assert result["specs"].get("mounting") == "SMD"
    assert result["specs"].get("package") == "SOD123"
    assert result["specs"].get("type_diode") == "Schottky"
    assert result["specs"].get("voltage") == "20V"
    assert result["specs"].get("current") == "500mA"
    assert result["specs"].get("manufacturer") == "Panjit"
    assert result["rich_description"] == "SMD SOD123, Schottky, 20V, 500mA,, SS0520,  Panjit"
    assert result["erp_code"] == "DIOSZ41S1000-PA"


def test_dynamic_manufacturer_loading_from_rules():
    from bom_enricher.comparison import load_rulebook
    from bom_enricher.datasheet_analyzer import extract_component_specs
    rulebook = load_rulebook("data/rules.xlsx")
    assert len(rulebook.manufacturers) > 100

    # 1. Unlisted manufacturer present in rules.xlsx (e.g. Qualitek -> QA, Kepo -> KE, Piher -> PI)
    text_qualitek = "Qualitek Delta Solder Wire Specifications"
    specs_q = extract_component_specs("RES", text_qualitek, "WIRE-01", rulebook=rulebook)
    assert specs_q.get("manufacturer") == "Qualitek"
    assert specs_q.get("manufacturer_code") == "QA"

    # 2. Collision test: code DI ('Diodes') must NOT match generic diode description
    text_generic_diode = "SURFACE MOUNT SCHOTTKY BARRIER DIODES VOLTAGE 40V"
    specs_dio = extract_component_specs("DIO", text_generic_diode, "GENERIC-01", rulebook=rulebook)
    assert specs_dio.get("manufacturer") != "Diodes"


def test_fast_path_never_suppresses_pdf_text():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # PDF containing package and manufacturer in text, with PN having electrical specs
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "PANJIT SEMICONDUCTOR")
    c.drawString(100, 730, "SS0520 SURFACE MOUNT SCHOTTKY BARRIER")
    c.drawString(100, 710, "Case: SOD-123, Plastic")
    c.save()
    pdf_bytes = buf.getvalue()

    result = analyze_datasheet(pdf_bytes, part_number="SS0520")
    # Even though SS0520 decoded electrical specs on fast-path, text MUST be read
    assert result["specs"].get("package") == "SOD123"
    assert result["specs"].get("manufacturer") == "Panjit"

    # Also verify that parts on fast-path (pn_has_core=True) retain full text extraction
    buf2 = io.BytesIO()
    c2 = canvas.Canvas(buf2)
    c2.drawString(100, 750, "SAMSUNG ELECTRO-MECHANICS")
    c2.drawString(100, 730, "CL21B104KBCNNNC Multilayer Ceramic Capacitor")
    c2.drawString(100, 710, "Operating Temperature: -55C to +125C")
    c2.save()
    pdf_bytes2 = buf2.getvalue()

    result2 = analyze_datasheet(pdf_bytes2, part_number="CL21B104KBCNNNC")
    assert result2["analysis_method"] == "fast_path"
    assert result2["specs"].get("temperature") == "-55°C~+125°C"


def test_universal_default_special_definition_fallback():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp, build_rich_description
    rulebook = load_rulebook("data/rules.xlsx")

    # 1. Diode (Special Definition - ALL has 2-char siblings AU, AC, AE -> formats as '00')
    specs_dio = {
        "mounting": "SMD",
        "package": "SOD123",
        "type_diode": "Schottky",
        "value_diode": "SS0520",
        "manufacturer": "Panjit",
        "manufacturer_code": "PA",
    }
    mappings_d, desc_d, code_d = map_specs_to_erp(rulebook, "DIO", specs_dio)
    assert code_d == "DIOSZ41S1000-PA"
    assert "Standart" not in desc_d

    # 2. Core attributes (like Package) must NEVER default to zero if unextracted
    specs_dio_no_pkg = {"mounting": "SMD", "type_diode": "Schottky", "value_diode": "SS0520"}
    mappings_np, _, code_np = map_specs_to_erp(rulebook, "DIO", specs_dio_no_pkg)
    pkg_m = next((m for m in mappings_np if m.attribute_name == "Package"), None)
    assert pkg_m is not None
    assert pkg_m.erp_code == "—"
    assert pkg_m.status == "BULUNAMADI"


def test_diverse_dynamic_manufacturers_from_rules_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Test 1: Qualitek (QA) - PDF canvas with QUALITEK INTERNATIONAL and 10K Resistor SMD 0603, part number RC0603FR-0710KL
    buf1 = io.BytesIO()
    c1 = canvas.Canvas(buf1)
    c1.drawString(100, 750, "QUALITEK INTERNATIONAL")
    c1.drawString(100, 730, "10K Resistor SMD 0603")
    c1.save()
    res1 = analyze_datasheet(buf1.getvalue(), part_number="RC0603FR-0710KL")
    assert res1["specs"].get("manufacturer") == "Qualitek"
    assert res1["erp_code"].endswith("-QA")

    # Test 2: Panjit Schottky Diode SS0520 - PDF canvas with PANJIT SEMICONDUCTOR, SURFACE MOUNT SCHOTTKY BARRIER RECTIFIER, SOD-123, Voltage 20V Current 0.5A, part number SS0520
    buf2 = io.BytesIO()
    c2 = canvas.Canvas(buf2)
    c2.drawString(100, 750, "PANJIT SEMICONDUCTOR")
    c2.drawString(100, 730, "SURFACE MOUNT SCHOTTKY BARRIER RECTIFIER")
    c2.drawString(100, 710, "SOD-123")
    c2.drawString(100, 690, "Voltage 20V Current 0.5A")
    c2.save()
    res2 = analyze_datasheet(buf2.getvalue(), part_number="SS0520")

    assert res2["erp_code"] == "DIOSZ41S1000-PA"
    assert res2["specs"]["manufacturer"] == "Panjit"
    assert res2["specs"]["package"] == "SOD123"
    for expected_token in ["SMD SOD123", "Schottky", "20V", "500mA", "SS0520", "Panjit"]:
        assert expected_token in res2["rich_description"]
    # Ensure "Standart" and "None" are NOT present in rich_description or suggested_description
    assert "Standart" not in res2["rich_description"]
    assert "None" not in res2["rich_description"]
    assert "Standart" not in res2["suggested_description"]
    assert "None" not in res2["suggested_description"]


def test_lmbr_and_mbr_schottky_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn

    # 1. LMBR0520T1G: LRC, 500mA, 20V, SOD-123
    specs_l = decode_manufacturer_pn("LMBR0520T1G")
    assert specs_l.get("manufacturer") == "LRC"
    assert specs_l.get("manufacturer_code") == "L0"
    assert specs_l.get("type_diode") == "Schottky"
    assert specs_l.get("mounting") == "SMD"
    assert specs_l.get("current") == "500mA"
    assert specs_l.get("voltage") == "20V"
    assert specs_l.get("package") == "SOD-123"

    # 2. MBR0520LT1G: On Semi, 500mA, 20V, SOD-123
    specs_m = decode_manufacturer_pn("MBR0520LT1G")
    assert specs_m.get("manufacturer") == "On Semi"
    assert specs_m.get("manufacturer_code") == "ON"
    assert specs_m.get("type_diode") == "Schottky"
    assert specs_m.get("current") == "500mA"
    assert specs_m.get("voltage") == "20V"
    assert specs_m.get("package") == "SOD-123"

    # 3. LMBR0540: 500mA, 40V
    specs_40 = decode_manufacturer_pn("LMBR0540")
    assert specs_40.get("voltage") == "40V"
    assert specs_40.get("current") == "500mA"


def test_unicode_dash_and_negative_context_guard_diode_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "0.5 AMPERES, 20 VOLTS")
    c.drawString(100, 730, "SURFACE MOUNT SCHOTTKY POWER RECTIFIER")
    # Notice Unicode minus (\u2212) simulating onsemi PDF text
    c.drawString(100, 710, "Plastic SOD\u2212123 Package")
    # Notice negative context mentioning MELF as an alternative
    c.drawString(100, 690, "This package provides an alternative to the leadless 34 MELF style package.")
    c.save()

    res = analyze_datasheet(buf.getvalue(), part_number="LMBR0520T1G")

    assert res["erp_code"] == "DIOSZ101Z500-L0"
    assert res["specs"]["package"] == "SOD-123"
    assert res["specs"]["manufacturer"] == "LRC"
    assert res["specs"]["manufacturer_code"] == "L0"
    assert res["specs"]["voltage"] == "20V"
    assert res["specs"]["current"] == "500mA"
    assert "MELF" not in res["erp_code"]
    for expected_token in ["SMD SOD-123", "Schottky", "20V", "500mA", "LMBR0520T1G", "LRC"]:
        assert expected_token in res["rich_description"]


def test_universal_is_critical_spec_missing():
    from bom_enricher.datasheet_analyzer import is_critical_spec_missing

    # Text empty always triggers
    assert is_critical_spec_missing("RES", {"resistance_str": "10k"}, is_text_empty=True) is True
    assert is_critical_spec_missing("BİLİNMİYOR", {}, is_text_empty=False) is True

    # DIO missing voltage/current
    assert is_critical_spec_missing("DIO", {"package": "SOD-123"}, is_text_empty=False) is True
    assert is_critical_spec_missing("DIO", {"voltage": "20V"}, is_text_empty=False) is False

    # RES missing resistance_str
    assert is_critical_spec_missing("RES", {"package": "0805"}, is_text_empty=False) is True
    assert is_critical_spec_missing("RES", {"resistance_str": "10k"}, is_text_empty=False) is False

    # CON missing pitch/pin_number/type_connector
    assert is_critical_spec_missing("CON", {"package": "SMD"}, is_text_empty=False) is True
    assert is_critical_spec_missing("CON", {"pitch": "2.54mm"}, is_text_empty=False) is False


def test_molex_connector_gender_and_picoflex_decoding(rulebook):
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, map_specs_to_erp, build_rich_description

    # 1. Test PN Decoder for Molex Picoflex 903250004
    specs = decode_manufacturer_pn("903250004")
    assert specs["manufacturer"] == "Molex"
    assert specs["pitch"] == "1.27mm"
    assert specs["pin_number"] == "4 pin"
    assert specs["type_connector"] == "Wire to Board Konnektör"
    assert specs["pin_type"] == "Erkek"
    assert specs["mounting"] == "TH"
    assert specs["angle"] == "180°(Dik)"

    # 2. Test ERP Mapping and Descriptions
    mappings, desc, code = map_specs_to_erp(rulebook, "CON", specs)
    assert code == "CONTADAO0420000XX-MO"
    assert desc == "TH Wire to Board Konnektör, Erkek, 1.27mm, 4 pin, 180°(Dik), Molex"

    rich_desc = build_rich_description("CON", specs, "")
    assert rich_desc == "TH Wire to Board Konnektör, Erkek, 1.27mm, 4 pin, 180°(Dik), Molex"

    # 3. Test with injected serial counter
    specs_injected = {**specs, "serial_counter": "0315"}
    _, _, code_injected = map_specs_to_erp(rulebook, "CON", specs_injected)
    assert code_injected == "CONTADAO0420315-MO"

    # 4. Test Molex DuraClik 5023520200
    dc_specs = decode_manufacturer_pn("5023520200")
    assert dc_specs["manufacturer"] == "Molex"
    assert dc_specs["pitch"] == "2mm"
    assert dc_specs["pin_number"] == "2 pin"
    assert dc_specs["pin_type"] == "Erkek"
    assert dc_specs["mounting"] == "SMD"
    assert dc_specs["angle"] == "90°(Yatık)"


def test_connector_pdf_mating_receptacle_does_not_override_male_header(rulebook):
    pdf_bytes = _create_sample_pdf(
        "Molex Picoflex PF-50 Header, Vertical",
        "90325-0004",
        [
            "Product Category: PCB Headers and Receptacles",
            "Component Type: PCB Header",
            "Pitch - Mating Interface 1.27mm",
            "Circuits: 4",
            "Gender Male",
            "Orientation: Vertical",
            "Termination Interface Style: Through Hole",
            "Mates with Part(s): Picoflex PF-50 Low Profile IDT Receptacles 90327",
        ],
    )
    result = analyze_single_datasheet("903250004.pdf", pdf_bytes, rulebook, part_number_override="903250004")
    assert result.material_type == "CON"
    codes_by_attr = {p.attribute_name: p.erp_code for p in result.properties}
    assert codes_by_attr["Pin Type of Connector"] == "A"
    assert codes_by_attr["Pitch Connector"] == "O"
    assert codes_by_attr["Pin Number"] == "04"
    assert codes_by_attr["Mounting"] == "T"
    assert result.suggested_erp_code == "CONTADAO0420000XX-MO"
    assert result.suggested_description == "TH Wire to Board Konnektör, Erkek, 1.27mm, 4 pin, 180°(Dik), Molex"
    assert result.rich_description == "TH Wire to Board Konnektör, Erkek, 1.27mm, 4 pin, 180°(Dik), Molex"


def test_molex_minifit_jr_locking_connector_and_2digit_pin(rulebook):
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, map_specs_to_erp, build_rich_description

    # 1. Test PN Decoder for 39281023 (5566-02A)
    specs = decode_manufacturer_pn("39281023")
    assert specs["manufacturer"] == "Molex"
    assert specs["pitch"] == "4.2mm"
    assert specs["pin_number"] == "2 pin"
    assert specs["type_connector"] == "Wire to Board Kilitli Konnektör"
    assert specs["pin_type"] == "Erkek"
    assert specs["mounting"] == "TH"
    assert specs["angle"] == "180°(Dik)"

    # 2. Test ERP Mapping: CON + T + AE + A + 3 + 02 + 2 + 0000XX + -MO -> CONTAEA30220000XX-MO
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "CON", specs)
    assert erp_code == "CONTAEA30220000XX-MO"
    assert desc == "TH Wire to Board Kilitli Konnektör, Erkek, 4.2mm, 2 pin, 180°(Dik), Molex"

    codes_by_attr = {m.attribute_name: m.erp_code for m in mappings}
    assert codes_by_attr["Type Connector"] == "AE"
    assert codes_by_attr["Pitch Connector"] == "3"
    assert codes_by_attr["Pin Number"] == "02"
    assert codes_by_attr["Angle"] == "2"

    # 3. Test with serial counter 0007 injected -> CONTAEA30220007-MO
    specs_injected = {**specs, "serial_counter": "0007"}
    _, _, erp_code_injected = map_specs_to_erp(rulebook, "CON", specs_injected)
    assert erp_code_injected == "CONTAEA30220007-MO"

    # 4. Test rich description
    rich_desc = build_rich_description("CON", specs, "")
    assert rich_desc == "TH Wire to Board Kilitli Konnektör, Erkek, 4.2mm, 2 pin, 180°(Dik), Molex"

    # 5. Test PDF extraction with locking text
    pdf_bytes = _create_sample_pdf(
        "Mini-Fit Jr. Header, Dual Row, Vertical, with Snap-in Plastic Peg PCB Lock, 2 Circuits",
        "39-28-1023",
        [
            "Overview: Mini-Fit Jr. Power Connectors",
            "Product Name: Mini-Fit Jr.",
            "Circuits: 2",
            "Color - Resin: Natural",
            "Durability (cycles max): 30",
            "Flammability: 94V-2",
            "Glow-Wire Capable: No",
            "Lock to Mating Part: Yes",
            "Material - Metal: Brass",
            "Material - Plating Mating: Tin",
            "Orientation: Vertical",
            "PCB Locator: Yes",
            "PCB Retention: Yes",
            "Pitch - Mating Interface: 4.20mm",
            "Termination Interface Style: Through Hole",
        ],
    )
    result = analyze_single_datasheet("39281023.pdf", pdf_bytes, rulebook, part_number_override="39281023")
    assert result.material_type == "CON"
    pdf_codes = {p.attribute_name: p.erp_code for p in result.properties}
    assert pdf_codes["Type Connector"] == "AE"
    assert result.suggested_erp_code == "CONTAEA30220000XX-MO"
    assert result.suggested_description == "TH Wire to Board Kilitli Konnektör, Erkek, 4.2mm, 2 pin, 180°(Dik), Molex"


def test_molex_sim_card_connector_detection_and_erp_mapping(rulebook):
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, map_specs_to_erp, build_rich_description

    # 1. Test PN Decoder for Molex 475532001 (Series 47553)
    specs = decode_manufacturer_pn("475532001")
    assert specs["manufacturer"] == "Molex"
    assert specs["type_connector"] == "SIM Card"
    assert specs["mounting"] == "SMD"
    assert specs["pin_number"] == "8 pin"
    assert specs["angle"] == "90°(Yatık)"
    assert specs["style"] == "Push-Push"

    # 2. Test ERP Mapping: CON + S + BD + 0 + 00 + 08 + 1 + 0000XX + -MO -> CONSBD0000810000XX-MO
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "CON", specs)
    assert erp_code == "CONSBD0000810000XX-MO"
    codes_by_attr = {m.attribute_name: m.erp_code for m in mappings}
    assert codes_by_attr["Mounting"] == "S"
    assert codes_by_attr["Type Connector"] == "BD"
    assert codes_by_attr["Pin Type of Connector"] == "0"
    assert codes_by_attr["Pitch Connector"] == "00"
    assert codes_by_attr["Pin Number"] == "08"
    assert codes_by_attr["Angle"] == "1"

    # 3. Test with injected serial counter 0562 -> CONSBD0000810562-MO
    specs_injected = {**specs, "serial_counter": "0562"}
    _, _, erp_code_injected = map_specs_to_erp(rulebook, "CON", specs_injected)
    assert erp_code_injected == "CONSBD0000810562-MO"

    # 4. Test rich description
    rich_desc = build_rich_description("CON", specs, "")
    assert rich_desc == "SMD SIM Card, Push-Push, 8 pin, 90°(Yatık), Molex"

    # 5. Test PDF extraction with real Molex 475532001 datasheet text
    pdf_bytes = _create_sample_pdf(
        "1.27mm Pitch SIM Card Holder with 2 Detect Pins, 6 Circuits, Push-Push Style, Housing Height 1.80mm, with Pegs, Lead-Free",
        "475532001",
        [
            "Product Category: Memory / SIM Card Connectors",
            "Series Number: 47553",
            "Status: Active",
            "Product Name: mini-SIM",
            "Style: Push-Push",
            "Card Detection Switch: Open",
            "Card Entry Location: Front",
            "Circuits (Loaded): 6",
            "Circuits (maximum): 6",
            "Current - Maximum per Contact: 0.5A",
            "Voltage - Maximum: 50V DC",
            "Pitch - Mating Interface: 2.54mm",
            "Pitch - Termination Interface: 1.27mm",
            "Termination Interface Style: Surface Mount",
            "Mates With: Standard SIM Card",
        ],
    )
    result = analyze_single_datasheet("475532001.pdf", pdf_bytes, rulebook, part_number_override="475532001")
    assert result.material_type == "CON"
    pdf_codes = {p.attribute_name: p.erp_code for p in result.properties}
    assert pdf_codes["Type Connector"] == "BD"
    assert pdf_codes["Mounting"] == "S"
    assert pdf_codes["Pin Type of Connector"] == "0"
    assert pdf_codes["Pitch Connector"] == "00"
    assert pdf_codes["Pin Number"] == "08"
    assert pdf_codes["Angle"] == "1"
    assert result.suggested_erp_code == "CONSBD0000810000XX-MO"
    assert result.rich_description == "SMD SIM Card, Push-Push, 8 pin, 90°(Yatık), Molex"


def test_keyword_context_guards_prevent_false_positives():
    from bom_enricher.datasheet_analyzer import classify_material_type, extract_component_specs

    # 1. Step-down SMPS converter mentions 'switches', 'housing', 'header', 'fuse protection' -> ICS, not SWT/CON/FUS
    smps_text = (
        "--- PAGE 1 ---\n"
        "MP2307 3A, 23V, 340kHz Synchronous Rectified Step-Down Converter\n"
        "DESCRIPTION\n"
        "The MP2307 integrates internal power switches to achieve high efficiency. "
        "The thermal housing provides low thermal resistance. "
        "Features internal fuse protection on input line and table header layout."
    )
    mat_type, conf, _ = classify_material_type(smps_text, part_number="MP2307DN")
    assert mat_type == "ICS"

    # 2. Resistor mentioning 'gate array' does not become 'Resistor Array'
    res_text = "Thick film chip resistor designed for connection to microcontroller gate array."
    res_specs = extract_component_specs("RES", res_text, "RC0805JR-0710KL")
    assert res_specs.get("composition") != "Resistor Array"
    assert res_specs.get("composition") == "Thick Film"

    # 3. SMD Buzzer mentioning 'Pin 1' and 'Pin 2' does not become TH
    buz_text = "SMD Magnetic Buzzer 3V. Surface mount reflow soldering. Pin 1 (+), Pin 2 (-)."
    buz_specs = extract_component_specs("BUZ", buz_text, "HNB09A03")
    assert buz_specs.get("mounting") == "SMD"

    # 4. Battery holder with 'clean contact surface' does not falsely become SMD
    bah_text = "Through hole coin cell battery holder for CR2032. Clean contact surface with dry cloth."
    bah_specs = extract_component_specs("BAH", bah_text, "BH-2032-1")
    assert bah_specs.get("mounting") == "TH"

    # 5. Transformer with 'Isolation voltage 3000V' does not falsely become Isolation Transformer
    trf_text = "Flyback SMPS Transformer for offline converters. 3000V isolation voltage test."
    trf_specs = extract_component_specs("TRF", trf_text, "TRF-FLY-01")
    assert trf_specs.get("type_transformer") == "SMPS Transformer"

    # 6. Switch with 'push-pull output' or 'pushing force' does not falsely become Push Button
    swt_text = "Slide switch with 2 positions. Pushing force is 250gf."
    swt_specs = extract_component_specs("SWT", swt_text, "SS-12D00")
    assert swt_specs.get("type_switch") == "Slide Switch"

    # 7. Connector with 'table header' does not falsely override Box Header
    con_text = "Box header connector 2.54mm pitch, 10 pin. See table header for pin assignments."
    con_specs = extract_component_specs("CON", con_text, "BH-10P")
    assert con_specs.get("type_connector") == "Box Header"


def test_adp_classification_and_decoding():
    from bom_enricher.datasheet_analyzer import (
        classify_material_type,
        decode_manufacturer_pn,
        extract_component_specs,
        map_specs_to_erp,
        build_rich_description,
        get_default_rulebook,
    )

    # 1. Part number decoding
    pn_specs = decode_manufacturer_pn("SK03T-1200200V")
    assert pn_specs["voltage_output"] == "12VDC"
    assert pn_specs["current_output"] == "2A"
    assert pn_specs["manufacturer"] == "SIMSUKIAN"
    assert pn_specs["manufacturer_code"] == "SMK"
    assert pn_specs["type_adaptor"] == "AC/DC Duvar Tipi Adaptör"

    # 2. Material type classification
    mat_type, conf, reason = classify_material_type("", "SK03T-1200200V")
    assert mat_type == "ADP"
    assert conf >= 0.90

    # 3. Title cover classification
    cover_text = "--- PAGE 1 ---\nSwitching Power Supply\nModel No: SK03T-1200200V\nShenzhen Simsukian Electronics"
    mat_type_cover, conf_cover, _ = classify_material_type(cover_text, "")
    assert mat_type_cover == "ADP"
    assert conf_cover == 0.99

    # 4. Rich description
    specs = {
        "type_adaptor": "AC/DC Duvar Tipi Adaptör",
        "voltage_adapter_input": "176 - 264VAC",
        "voltage_output": "12VDC",
        "current_output": "2A",
        "cable": "1.5mm Cable",
        "serial_counter": "011",
        "manufacturer": "SIMSUKIAN",
    }
    desc = build_rich_description("ADP", specs)
    assert desc == "AC/DC Duvar Tipi Adaptör, 176 - 264VAC, 12VDC, 2A, 1.5mm Cable, SIMSUKIAN"

    # 5. ERP Mapping
    rb = get_default_rulebook()
    mappings, s_desc, erp_code_default = map_specs_to_erp(rb, "ADP", specs)
    assert erp_code_default == "ADPAA00A301011-SMK"
    assert s_desc == "AC/DC Duvar Tipi Adaptör, 176 - 264VAC, 12VDC, 2A, 1.5mm Cable, SIMSUKIAN"

    specs_counter = dict(specs)
    specs_counter["serial_counter"] = "001"
    _, _, erp_code_counter = map_specs_to_erp(rb, "ADP", specs_counter)
    assert erp_code_counter == "ADPAA00A301001-SMK"


def test_real_simsukian_adapter_datasheet_e2e():
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    pdf_path = Path(r"C:\Users\sudem\.gemini\antigravity\brain\8be84be6-1095-43d5-b848-ce9fe57a8347\.user_uploaded\media_1792570576636_17351353.pdf")
    if not pdf_path.exists():
        pytest.skip("User uploaded PDF not found at path")

    # Analyze without providing part_number (simulating raw upload)
    res_default = analyze_datasheet(pdf_path, filename="media_1792570576636_17351353.pdf")
    assert res_default["material_type"] == "ADP"
    assert res_default["part_number"] == "SK03T-1200200V"
    assert res_default["erp_code"] == "ADPAA00A301011-SMK"
    assert res_default["rich_description"] == "AC/DC Duvar Tipi Adaptör, 176 - 264VAC, 12VDC, 2A, 1.5mm Cable, SIMSUKIAN"

    # Analyze with injected serial counter
    res_counter = analyze_datasheet(pdf_path, filename="media_1792570576636_17351353.pdf", serial_counter="011")
    assert res_counter["material_type"] == "ADP"
    assert res_counter["part_number"] == "SK03T-1200200V"
    assert res_counter["erp_code"] == "ADPAA00A301011-SMK"


def test_real_ezurio_maf95310_antenna_datasheet_e2e():
    """Verify real Laird / Ezurio MAF95310 Antenna datasheet produces exact expected ERP code and rich description."""
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from bom_enricher.comparison import load_rulebook

    pdf_path = Path(r"C:\Users\sudem\.gemini\antigravity\brain\8be84be6-1095-43d5-b848-ce9fe57a8347\.user_uploaded\media_1792578508694_12998fb7.pdf")
    if not pdf_path.exists():
        pytest.skip("User uploaded PDF not found at path")

    rule_file = Path("data/rules.xlsx")
    rulebook = load_rulebook(rule_file) if rule_file.exists() else None

    # 1. Analyze with raw filename (simulating user upload without part number)
    res_default = analyze_datasheet(pdf_path, filename="media_1792578508694_12998fb7.pdf", rulebook=rulebook)
    assert res_default["material_type"] == "ANT"
    assert res_default["part_number"] == "MAF95310"
    assert res_default["specs"]["manufacturer"] == "Ezurio"
    assert res_default["specs"]["cable_length"] == "185mm"
    assert res_default["specs"]["mating_connector"] == "UFL (I-PEX)"
    assert res_default["specs"]["structure"] == "Flex"
    assert res_default["rich_description"] == "RF Flex Antenna, WLAN, 185mm UFL Cable (I-PEX), Ezurio"
    assert res_default["erp_code"] == "ANTT30000XX-EZU"

    # 2. Analyze with serial_counter="21"
    res_counter = analyze_datasheet(pdf_path, filename="media_1792578508694_12998fb7.pdf", rulebook=rulebook, serial_counter="21")
    assert res_counter["material_type"] == "ANT"
    assert res_counter["part_number"] == "MAF95310"
    assert res_counter["erp_code"] == "ANTT3000021-EZU"
    assert res_counter["rich_description"] == "RF Flex Antenna, WLAN, 185mm UFL Cable (I-PEX), Ezurio"


def test_auto_material_type_emf2449a1_10ufl_antenna_e2e():
    """Verify EMF2449A1-10UFL is classified as ANT under AUTO product type without false capacitor match."""
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet, classify_material_type
    from bom_enricher.comparison import load_rulebook

    # Test PN classification with empty text does NOT match CAP
    mat_type, conf, reason = classify_material_type("", "EMF2449A1-10UFL")
    assert mat_type == "ANT"
    assert conf >= 0.85

    pdf_path = Path(r"C:\Users\sudem\.gemini\antigravity\brain\8be84be6-1095-43d5-b848-ce9fe57a8347\.user_uploaded\media_1792581483711_8d7964b4.pdf")
    if not pdf_path.exists():
        pdf_path = Path(r"C:\Users\sudem\.gemini\antigravity\brain\8be84be6-1095-43d5-b848-ce9fe57a8347\.user_uploaded\media_1792578508694_12998fb7.pdf")
    if not pdf_path.exists():
        pytest.skip("User uploaded PDF not found at path")

    rule_file = Path("data/rules.xlsx")
    rulebook = load_rulebook(rule_file) if rule_file.exists() else None

    # Analyze with AUTO / default material type
    res = analyze_datasheet(pdf_path, part_number="EMF2449A1-10UFL", rulebook=rulebook)
    assert res["material_type"] == "ANT"
    assert res["part_number"] == "EMF2449A1-10UFL"
    assert res["specs"]["manufacturer"] == "Ezurio"
    assert res["specs"]["cable_length"] == "100mm"
    assert res["specs"]["mating_connector"] == "UFL (I-PEX)"
    assert res["erp_code"] == "ANTT30000XX-EZU"
    assert res["rich_description"] == "RF Flex Antenna, WLAN, 100mm UFL Cable (I-PEX), Ezurio"


