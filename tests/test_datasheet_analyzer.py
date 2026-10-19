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
    assert code == "CONTADAB410000XX-TT"
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
    assert pin_m.erp_code == "4"


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

