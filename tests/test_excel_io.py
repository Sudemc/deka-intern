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
