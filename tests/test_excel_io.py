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

    with pytest.raises(ValueError) as duplicate_error:
        read_customer_bom(duplicate_path)
    assert "duplicate.xlsx" in str(duplicate_error.value)
    assert "Duplicate Fields" in str(duplicate_error.value)
    assert "Value" in str(duplicate_error.value)


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
