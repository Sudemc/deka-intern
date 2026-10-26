from openpyxl import Workbook
import pytest

from bom_enricher.excel_io import ColumnMapping, read_customer_bom


def test_read_customer_sample():
    rows = read_customer_bom("examples/customer_bom_sample.xlsx")

    assert len(rows) == 5
    assert rows[0].part == "C1, C2"
    assert rows[0].qty == 2
    assert rows[0].value == "10 uF 16 V X7R 0805"
    assert rows[0].part_number == "SYN-CAP-001"
    assert rows[0].manufacturer == "Example Components Ltd. (fictional)"
    assert rows[2].value == ""
    assert rows[4].qty == 3


def test_missing_and_duplicate_headers(tmp_path):
    missing_path = tmp_path / "missing.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Missing Fields"
    sheet.append(["Part", "Qty", "Part Number"])
    sheet.append(["C1", 1, "SYN-001"])
    workbook.save(missing_path)

    missing_rows = read_customer_bom(missing_path)
    assert len(missing_rows) == 1
    assert missing_rows[0].part == "C1"
    assert missing_rows[0].qty == 1
    assert missing_rows[0].value == ""
    assert missing_rows[0].part_number == "SYN-001"

    duplicate_path = tmp_path / "duplicate.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Duplicate Fields"
    sheet.append(["Part", "Qty", "Value", "Value", "Part Number"])
    sheet.append(["C1", 1, "first", "second", "SYN-001"])
    workbook.save(duplicate_path)

    duplicate_rows = read_customer_bom(duplicate_path)
    assert len(duplicate_rows) == 1
    assert duplicate_rows[0].part == "C1"
    assert duplicate_rows[0].qty == 1
    assert duplicate_rows[0].value == "first - second"
    assert duplicate_rows[0].part_number == "SYN-001"


def test_custom_column_mapping(tmp_path):
    path = tmp_path / "custom.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Parts"
    sheet.append(["Reference", "Count", "Spec", "MPN"])
    sheet.append(["R1, R2", 2, "5.1k 1% 0603", "SYN-R-01"])
    sheet.append([None, 1, None, "SYN-R-02"])
    workbook.save(path)

    rows = read_customer_bom(
        path,
        sheet_name="Parts",
        mapping=ColumnMapping(
            part="Reference",
            qty="Count",
            value="Spec",
            part_number="MPN",
        ),
    )

    assert len(rows) == 2
    assert rows[0].part == "R1, R2"
    assert rows[0].qty == 2
    assert rows[0].value == "5.1k 1% 0603"
    assert rows[0].part_number == "SYN-R-01"
    assert rows[0].manufacturer == ""
    assert rows[1].part == ""
    assert rows[1].value == ""

    with pytest.raises(ValueError) as ambiguous_mapping_error:
        read_customer_bom(
            path,
            sheet_name="Parts",
            mapping=ColumnMapping(
                part="Reference",
                qty="Reference",
                value="Spec",
                part_number="MPN",
            ),
        )
    assert "custom.xlsx" in str(ambiguous_mapping_error.value)
    assert "Parts" in str(ambiguous_mapping_error.value)
    assert "same source header" in str(ambiguous_mapping_error.value)


def test_auto_detects_header_row_synonyms_and_unlabeled_product_code_column(tmp_path):
    path = tmp_path / "variable_layout.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Customer BOM"
    sheet.append(["Customer assembly export"])
    sheet.append([None])
    sheet.append(["Ref", "Quantity", "Description", "Internal Material"])
    sheet.append(["C1, C2", 2, "10 uF 16 V X7R 0805", "CAP0603X7R10UF"])
    sheet.append(["R1", 1, "1 kOhm 1% 0603", "RES0603-1K-1PCT"])
    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)

    assert len(rows) == 2
    assert rows[0].source_row == 4
    assert rows[0].part == "C1, C2"
    assert rows[0].qty == 2
    assert rows[0].value == "10 uF 16 V X7R 0805"
    assert rows[0].part_number == ""
    assert rows[0].product_code_candidates[0].value == "CAP0603X7R10UF"
    assert rows[0].product_code_candidates[0].source_column == "Internal Material"
    assert rows[1].product_code_candidates[0].value == "RES0603-1K-1PCT"


def test_auto_detection_keeps_multiple_code_candidates_for_manual_resolution(tmp_path):
    path = tmp_path / "ambiguous_code_columns.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Reference", "Description", "Vendor Ref", "Internal Code"])
    sheet.append(["U1", "logic device", "ABC12345", "XYZ98765"])
    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)

    assert len(rows) == 1
    assert len(rows[0].product_code_candidates) == 2
    assert {candidate.value for candidate in rows[0].product_code_candidates} == {
        "ABC12345",
        "XYZ98765",
    }


def test_selects_bom_sheet_in_workbook_with_revision_cover(tmp_path):
    path = tmp_path / "customer_with_cover.xlsx"
    workbook = Workbook()
    cover = workbook.active
    cover.title = "Revision Cover"
    cover.append([None, "Date", "Revision", "Revision Content"])
    cover.append(["REVISION", "2026-07-03", 5, "Document updated"])
    bom = workbook.create_sheet("Current")
    bom.append(["Product", "Remote Heater Control"])
    bom.append(["BOM"])
    bom.append(["No", "Part", "Qty", "Description", "TH/SMD", "Part Number"])
    bom.append([1, "C1,C4,C5", 3, "SMD 0805 Package, 100nF, 50V", "SMD", "CL21B104KBCNNNC"])
    bom.append(["Revision History"])
    bom.append(["No", "Rev. Number", None, "Description", "Date", "Published By"])
    bom.append([1, "Rev0.0", None, "BOM follow-up note", "2026-09-03", "Melih Dündar"])
    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)

    code_rows = [row for row in rows if row.part_number]
    assert len(code_rows) == 1
    assert code_rows[0].part_number == "CL21B104KBCNNNC"
    assert code_rows[0].part == "C1,C4,C5"
    assert code_rows[0].value == "SMD 0805 Package, 100nF, 50V"
    assert all(
        candidate.value != "Melih Dündar"
        for row in rows
        for candidate in row.product_code_candidates
    )


def test_turkish_header_normalization():
    from bom_enricher.excel_io import _header_field, _normalize_header
    assert _normalize_header("ÜRETİCİ KODU") == "uretici kodu"
    assert _header_field("ÜRETİCİ KODU") == "part_number"
    assert _header_field("REFERANS") == "part"
    assert _header_field("DEĞER") == "value"


def test_parca_and_referans_disambiguation(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom

    file_path = tmp_path / "customer_bom_with_parca_and_ref.xlsx"
    df = pd.DataFrame({
        "Parça": [37, 44],
        "REFERANS": ["D13,D15", "D29,D30"],
        "MIKTAR": [2, 8],
        "ÜRETİCİ KODU": ["MBR0540T1G", "SD12CT1G"],
        "DEĞER": ["MBR0540", "SD12CT1G"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 2
    assert rows[0].part == "D13,D15"
    assert rows[1].part == "D29,D30"


def test_generic_multi_code_candidate_detection(tmp_path):
    from bom_enricher.excel_io import read_customer_bom
    import pandas as pd
    file_path = tmp_path / "multi_code_customer_bom.xlsx"
    df = pd.DataFrame({
        "Parça": [37],
        "DTSIS KODU": ["H010101005006"],
        "ÜRETİCİ KODU": ["MBR0540T1G"],
        "REFERANS": ["D13,D15"],
        "MIKTAR": [2],
        "DEĞER": ["MBR0540"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 1
    candidate_values = {c.value for c in rows[0].product_code_candidates}
    assert "H010101005006" in candidate_values
    assert "MBR0540T1G" in candidate_values


def test_altium_bom_format(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom

    file_path = tmp_path / "altium_bom.xlsx"
    df = pd.DataFrame({
        "Comment": ["100nF", "10k"],
        "Description": ["CAP CER 100NF 50V X7R 0603", "RES 10K 1% 0603"],
        "Designator": ["C1, C2", "R1"],
        "Footprint": ["0603", "0603"],
        "LibRef": ["CAP_0603_100NF", "RES_0603_10K"],
        "Quantity": [2, 1],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 2
    assert rows[0].part == "C1, C2"
    assert rows[0].qty == 2
    assert rows[0].value == "100nF - CAP CER 100NF 50V X7R 0603"
    assert rows[0].part_number == "CAP_0603_100NF"
    assert rows[0].package == "0603"


def test_kicad_bom_format_with_inferred_qty(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom

    file_path = tmp_path / "kicad_bom.xlsx"
    # KiCad BOM often has no Quantity column, Qty is implicit in Reference count
    df = pd.DataFrame({
        "Reference": ["C1, C2, C3", "U1"],
        "Value": ["100nF", "STM32F401RET6"],
        "Footprint": ["Capacitor_SMD:C_0603", "Package_QFP:LQFP-64"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 2
    assert rows[0].part == "C1, C2, C3"
    assert rows[0].qty == 3
    assert rows[0].value == "100nF"
    assert rows[1].part == "U1"
    assert rows[1].qty == 1
    assert rows[1].part_number == "STM32F401RET6"


def test_refdes_range_expansion_and_qty_inference(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom, split_references

    assert split_references("R1-R5, R8, C10-C12") == [
        "R1", "R2", "R3", "R4", "R5", "R8", "C10", "C11", "C12"
    ]

    file_path = tmp_path / "range_bom.xlsx"
    df = pd.DataFrame({
        "Designator": ["R1-R5", "C10-C14"],
        "Val": ["10k", "100nF"],
        "P/N": ["RC0603FR-0710KL", "CC0603KRX7R9BB104"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 2
    assert rows[0].qty == 5
    assert rows[0].part_number == "RC0603FR-0710KL"
    assert rows[1].qty == 5
    assert rows[1].part_number == "CC0603KRX7R9BB104"


def test_dual_package_columns_combination(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom

    file_path = tmp_path / "dual_pkg_bom.xlsx"
    df = pd.DataFrame({
        "Part": ["U1"],
        "Qty": [1],
        "Part Number": ["MAX3232ESE+"],
        "Value": ["RS232 Transceiver"],
        "Footprint": ["SOIC-16"],
        "Package": ["SMD"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 1
    assert rows[0].package == "SOIC-16 SMD"


def test_pair_lr_references():
    from bom_enricher.excel_io import pair_lr_references

    # Prefix L vs R
    paired, pairs = pair_lr_references(["ST1L_B1", "ST1L_B2", "ST1R_B1", "ST1R_B2"])
    assert paired is True
    assert len(pairs) == 2

    # Infix L vs R
    paired, pairs = pair_lr_references(["ST_L_R1", "ST_R_R1"])
    assert paired is True
    assert len(pairs) == 1

    # Suffix L vs R
    paired, pairs = pair_lr_references(["R1_L", "R1_R", "R2_L", "R2_R"])
    assert paired is True
    assert len(pairs) == 2

    # Unbalanced or non-symmetrical
    paired, _ = pair_lr_references(["R1", "R2", "R3"])
    assert paired is False

    paired, _ = pair_lr_references(["ST1L_B1", "ST1R_B1", "ST1L_B2"])
    assert paired is False


def test_customer_bom_lr_symmetry_qty(tmp_path):
    import pandas as pd
    from bom_enricher.excel_io import read_customer_bom

    file_path = tmp_path / "lr_bom.xlsx"
    df = pd.DataFrame({
        "Designator": ["ST1L_B1, ST1L_B2, ST1R_B1, ST1R_B2", "R1_L, R1_R"],
        "Part Number": ["ZTL432AQFTA", "RC0603FR-0710KL"],
        "Quantity": [2, 1],  # Customer entered single-lamp / per-side qty
        "Description": ["Voltage Ref", "Resistor 10k"],
    })
    df.to_excel(file_path, index=False)
    rows = read_customer_bom(file_path)
    assert len(rows) == 2
    # Inferred to total vehicle set quantity
    assert rows[0].qty == 4
    assert rows[1].qty == 2


def test_multi_table_vertical_stacking(tmp_path):
    """Sheets containing stacked tables with different column structures (like Aydın format)."""
    path = tmp_path / "multi_table.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BOM"

    # Table 1: Standard layout (Part, Value, Qty, Part Number)
    sheet.append(["Part", "Value", "Qty", "Part Number"])
    sheet.append(["C1, C2", "100nF 50V", 2, "CAP-001"])
    sheet.append(["R1", "10k 1%", 1, "RES-001"])

    # Blank rows between tables
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])

    # Table 2: YMML layout (Item, Status, Qty, Stock No, MPN, Ref, Value, Footprint, Mounting)
    sheet.append(["Parça", "Durum", "Miktar", "Stok No", "Üretici Kodu", "Referans", "Değer", "Kılıf", "Montaj Tipi"])
    sheet.append([1, "FULL", 1, "STK-100", "MBR0540T1G", "D1", "Schottky Diode", "SOD-123", "SMD"])
    sheet.append([2, "FULL", 2, "STK-200", "SD12CT1G", "D2, D3", "TVS Diode", "SOD-323", "SMD"])

    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)
    assert len(rows) == 4
    # From Table 1
    assert rows[0].part == "C1, C2"
    assert rows[0].part_number == "CAP-001"
    assert rows[0].qty == 2
    # From Table 2 (dynamically re-resolved column indices)
    assert rows[2].part == "D1"
    assert rows[2].part_number == "MBR0540T1G"
    assert rows[2].qty == 1
    assert rows[2].mounting == "SMD"
    assert rows[3].part == "D2, D3"
    assert rows[3].part_number == "SD12CT1G"
    assert rows[3].qty == 2


def test_multi_section_repeated_headers(tmp_path):
    """Sheets with repeated BOM headers per assembly section (like FEKA format)."""
    path = tmp_path / "multi_section.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sections"

    # Section 1
    sheet.append(["--- ST1 MODULE ---"])
    sheet.append(["FEKA NO", "Part Number", "Description", "Designator", "QTY"])
    sheet.append(["FK-001", "LM2903BYDT", "Dual Comparator", "ST1L_U1, ST1R_U1", 2])

    # Section 2
    sheet.append(["--- ST2 MODULE ---"])
    sheet.append(["FEKA NO", "Part Number", "Description", "Designator", "QTY"])
    sheet.append(["FK-002", "LA G6SP.02", "Yellow LED", "ST2L_LED1, ST2R_LED1", 2])

    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)
    assert len(rows) == 2
    assert rows[0].part == "ST1L_U1, ST1R_U1"
    assert rows[0].part_number == "LM2903BYDT"
    assert rows[1].part == "ST2L_LED1, ST2R_LED1"
    assert rows[1].part_number == "LA G6SP.02"


def test_cad_export_with_dummy_placeholders(tmp_path):
    """Altium/CAD exports with single-letter placeholder codes ('c', 'C', '-') in manufacturer column."""
    path = tmp_path / "cad_export.xlsx"
    workbook = Workbook()
    sheet = workbook.active

    sheet.append(["Item", "Part", "Quantity", "Part Number", "Description", "Üretici Kodu"])
    # Row 1 has both customer internal stock code and genuine manufacturer PN
    sheet.append([1, "C308", 1, "CR-HM-001-0001", "Capacitor 100uF", "865080553014"])
    # Row 2 has customer stock code, but dummy 'c' in manufacturer column
    sheet.append([2, "C100, C101", 2, "CR-HM-001-0006", "Capacitor 100nF", "c"])
    # Row 3 has no customer code and dummy '-' in manufacturer column, but valid ref and value
    sheet.append([3, "R10, R11", 2, "-", "Resistor 1k 0603", "R"])

    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)
    assert len(rows) == 3
    # Row 1 keeps MPN
    assert rows[0].part_number in {"865080553014", "CR-HM-001-0001"}
    assert {c.value for c in rows[0].product_code_candidates} == {"865080553014", "CR-HM-001-0001"}
    # Row 2 filters out dummy 'c' and keeps genuine customer code
    assert rows[1].part_number == "CR-HM-001-0006"
    assert "c" not in [c.value for c in rows[1].product_code_candidates]
    # Row 3 filters out '-' and 'R' placeholders
    assert rows[2].part_number == ""
    assert "R" not in [c.value for c in rows[2].product_code_candidates]
    assert "-" not in [c.value for c in rows[2].product_code_candidates]


def test_terminal_non_bom_sections_pruned(tmp_path):
    """Terminal manufacturing process and revision history tables (DKF-RD-017 / EGLO format)."""
    path = tmp_path / "engineering_doc.xlsx"
    workbook = Workbook()
    sheet = workbook.active

    # BOM Header & Data
    sheet.append(["Item", "Part", "Description", "Quantity", "Part Number"])
    sheet.append([1, "C1", "Capacitor 100nF", 1, "AC0603KRX7R9BB104"])
    sheet.append([2, "R1", "Resistor 220R", 1, "0603(S/W)AJ0221T5E"])

    # Terminal Manufacturing Process Table
    sheet.append([None, None, None, None, None])
    sheet.append(["Üretim Prosesleri", None, None, None, None])
    sheet.append(["Proses Adı", "Açıklama", "Operatör", "Ekipman", None])
    sheet.append(["SMD Dizgi", "SMD Dizgi Yapılacak", "Ahmet Y.", "Yamaha YSM20", None])
    sheet.append(["Selektif Lehim", "Selektif Lehim Yapılacak", "Mehmet K.", "Ersa Versaflow", None])

    # Terminal Revision History Table
    sheet.append([None, None, None, None, None])
    sheet.append(["Revizyon Tarihçesi", None, None, None, None])
    sheet.append(["Tarih", "Revizyon", "Açıklama", "Revize Eden", "Onaylayan"])
    sheet.append(["2024-05-10", "01", "İlk yayın", "Ali ÇOLAK", "Hüseyin BİÇEN"])

    workbook.save(path)
    workbook.close()

    rows = read_customer_bom(path)
    # Exactly 2 component rows, 0 ghost rows from processes or revision history
    assert len(rows) == 2
    assert rows[0].part == "C1"
    assert rows[0].part_number == "AC0603KRX7R9BB104"
    assert rows[1].part == "R1"
    assert rows[1].part_number == "0603(S/W)AJ0221T5E"


def test_delta_v_and_altium_patterns():
    from bom_enricher.excel_io import clean_part_number, split_references, _header_field, _explicit_code_like

    # 1. Company prefix extraction vs bin notes
    assert clean_part_number("DELTA-V (IXTT140P10T)") == "IXTT140P10T"
    assert clean_part_number("LA G6SP.02 (8E-5F)(2-4 color)") == "LA G6SP.02"
    assert clean_part_number("LY G6SP.02 (6E-5F)(5-6 color)") == "LY G6SP.02"

    # 2. Designator prose stripping
    assert split_references("M3- müşteriden gelecek") == ["M3"]
    assert split_references("ACS1, ACS2") == ["ACS1", "ACS2"]
    assert split_references("C1, C7, C13, C14") == ["C1", "C7", "C13", "C14"]

    # 3. Header mapping for '#', 'No.'
    assert _header_field("#") == "item"
    assert _header_field("No") == "item"
    assert _header_field("No.") == "item"

    # 4. Part numbers with commas and unclosed parens
    assert _explicit_code_like("BT151S-500R,118") is True
    assert _explicit_code_like("TLP293(BL-TPL,E") is True

