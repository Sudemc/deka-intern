from openpyxl import Workbook, load_workbook

from bom_enricher.comparison import compare_workbooks, load_rulebook


def _write_book(path, sheet_name, headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def _source_files(tmp_path):
    customer = tmp_path / "BOM_Malzeme_Listesi.xlsx"
    erp = tmp_path / "Kitap1.xlsx"
    rules = tmp_path / "Malzeme_Tipleri.xlsx"
    _write_book(customer, "BOM", ["Part Number", "Value"], [
        ["CAP-001", "10uF 50V 1206"],
        ["RES-002", None],
        ["CON-003", "2 pin connector TH"],
        ["ONLY-CUSTOMER", "SMD 0603"],
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part"], [
        [" cap-001 ", "SMD 0805 Package, 10uF, 25V", "S", "CAP", "C1"],
        ["RES-002", "SMD 0603 Package, 10k, 1%", "S", "RES", "R1"],
        ["CON-003", "TH 2 pin connector", "S", "CON", "J1"],
        ["ONLY-ERP", "SMD 0603 Package, 1k, 1%", "S", "RES", "R2"],
    ])
    _write_book(rules, "BİRLEŞTİRİLMİŞ", [
        "MALZEME TİPİ", "ÖZELLİK (ATTRIBUTE)", "ÖZELLİK AÇIKLAMASI", "VALUE",
        "VALUE AÇIKLAMASI", "LANGU", "SIRALAMA (SORTNO)", "PREFIX", "SUFFIX",
    ], [
        ["CAP", 1, "Mounting", "S", "SMD", "T", 0, "CAP", None],
        ["CAP", 1, "Mounting", "T", "TH", "T", 0, "CAP", None],
        ["CAP", 105, "Package", "S3", "0805 Package,", "T", 10, None, None],
        ["CAP", 16, "Value Capacitors", "A1", "10uF,", "T", 20, None, None],
        ["CAP", 10, "Voltage", "C7", "25V,", "T", 30, None, None],
        ["CAP", 10, "Voltage", "D0", "50V,", "T", 30, None, None],
        ["RES", 1, "Mounting", "S", "SMD", "T", 10, "RES", None],
        ["RES", 1, "Mounting", "T", "TH", "T", 10, "RES", None],
        ["RES", 2, "Package", "S3", "0603 Package,", "T", 20, None, None],
        ["RES", 15, "Value Resistors", "R0", "10k,", "T", 30, None, None],
        ["RES", 4, "Tolerance", "F", 0.01, "T", 40, None, None],
        ["CON", 1, "Mounting", "T", "TH", "T", 0, "CON", None],
        ["CON", 1, "Mounting", "S", "SMD", "T", 0, "CON", None],
        ["CON", 41, "Pin Number", 2, "2 pin,", "T", 20, None, None],
        ["RES", 7, "Q", "YA", "Yageo", "T", 80, "-", None],
    ])
    return customer, erp, rules


def test_loads_every_rule_dimension_from_combined_sheet(tmp_path):
    _, _, rules_path = _source_files(tmp_path)

    rules = load_rulebook(rules_path)

    assert "cap" in rules.by_type_attribute
    assert "package" in rules.by_type_attribute["cap"]
    assert any(entry.prefix == "CAP" and entry.sort_order == "0" for entry in rules.entries)
    assert any(entry.language == "T" for entry in rules.entries)
    assert any(entry.value == "YA" and entry.prefix == "-" for entry in rules.entries)
    assert all(entry.suffix == "" for entry in rules.entries)


def test_comparison_reports_missing_conflicting_blank_and_mounting_mismatch(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    output = tmp_path / "Hata_ve_Kiyaslama_Raporu.xlsx"

    result = compare_workbooks(customer, erp, rules, output)

    assert result.total_part_count == 5
    rows = {row["Part Number"].strip(): row for row in result.rows}
    assert "Açıklama Çelişkisi" in rows["CAP-001"]["Hata/Uyarı Durumu"]
    assert "Müşteri verisi boş, ERP verisi kullanılabilir" in rows["RES-002"]["Hata/Uyarı Durumu"]
    assert "TH/SMD ile Description montaj tipi çelişkisi" in rows["CON-003"]["Hata/Uyarı Durumu"]
    assert "Sistemde Yok" in rows["ONLY-CUSTOMER"]["Hata/Uyarı Durumu"]
    assert "Sistemde Yok" in rows["ONLY-ERP"]["Hata/Uyarı Durumu"]
    workbook = load_workbook(output, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        assert [cell.value for cell in sheet[1]] == [
            "Durum", "Part Number", "Müşteri Qty", "ERP Qty", "Müşteri Value", "ERP Description", "Müşteri TH/SMD", "ERP TH/SMD", "Hata/Uyarı Durumu"
        ]
        assert sheet.max_row == 6
    finally:
        workbook.close()


def test_duplicate_part_numbers_are_reported_without_cartesian_merge(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value"], [
        ["CAP-001", "10uF 50V"], ["CAP-001", "10uF 50V"]
    ])
    output = tmp_path / "duplicates.xlsx"

    result = compare_workbooks(customer, erp, rules, output)

    duplicate = next(row for row in result.rows if row["Part Number"] == "CAP-001")
    assert "yinelenen Part Number" in duplicate["Hata/Uyarı Durumu"]
    assert sum(row["Part Number"] == "CAP-001" for row in result.rows) == 1


def test_detects_header_after_title_rows_and_rejects_missing_required_sheet(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Müşteri BOM listesi"], [])
    workbook = load_workbook(customer)
    sheet = workbook.active
    sheet.append(["Ürün Kodu", "Value"])
    sheet.append(["C-1", "1uF 50V"])
    workbook.save(customer)
    workbook.close()

    result = compare_workbooks(customer, erp, rules, tmp_path / "header.xlsx")

    assert any(row["Part Number"] == "C-1" for row in result.rows)


def test_customer_code_column_with_unfamiliar_heading_is_inferred_from_values(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Müşteri özel alanı", "Value"], [["CUSTOM-10001", "1uF"]])

    result = compare_workbooks(customer, erp, rules, tmp_path / "inferred.xlsx")

    assert any(row["Part Number"] == "CUSTOM-10001" for row in result.rows)


def test_description_values_are_checked_against_inferred_component_type(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Part"], [
        ["cap-001", "SMD 0805 Package, 10k, 25V", "S", "C1"]
    ])

    result = compare_workbooks(customer, erp, rules, tmp_path / "property-mismatch.xlsx")

    row = next(row for row in result.rows if row["Part Number"].strip() == "CAP-001")
    assert "Description sözlük uyumsuzluğu" in row["Hata/Uyarı Durumu"]


def test_description_attribute_order_uses_sortno_from_dictionary(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part"], [
        ["CAP-001", "50V 10uF 0805 Package SMD", "S", "CAP", "C1"]
    ])

    result = compare_workbooks(customer, erp, rules, tmp_path / "sort-order.xlsx")

    row = next(row for row in result.rows if row["Part Number"] == "CAP-001")
    assert "SORTNO sırasıyla uyuşmuyor" in row["Hata/Uyarı Durumu"]


def test_unresolved_material_type_is_not_reported_as_fully_validated(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD"], [
        ["cap-001", "SMD 0805 Package", "S"]
    ])

    result = compare_workbooks(customer, erp, rules, tmp_path / "unknown-type.xlsx")

    row = next(row for row in result.rows if row["Part Number"].strip() == "CAP-001")
    assert "Malzeme tipi belirlenemedi" in row["Hata/Uyarı Durumu"]


def test_value_valid_for_selected_type_is_not_flagged_for_other_type_attributes(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    workbook = load_workbook(rules)
    sheet = workbook["BİRLEŞTİRİLMİŞ"]
    sheet.append(["CAP", 100, "Finish", "STD", "Standard", "T", 90, None, None])
    sheet.append(["DIO", 12, "Size Capacitor", "S", "Standard", "T", 30, None, None])
    workbook.save(rules)
    workbook.close()
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part"], [
        ["CAP-001", "SMD 0805 Package, Standard", "S", "CAP", "C1"]
    ])

    result = compare_workbooks(customer, erp, rules, tmp_path / "valid-overlap.xlsx")

    row = next(row for row in result.rows if row["Part Number"].strip() == "CAP-001")
    assert "Description sözlük uyumsuzluğu" not in row["Hata/Uyarı Durumu"]


def test_punctuation_insensitive_part_number_matching(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value"], [
        ["RC0603FR-0710KL", "10k 1% 0603"]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type"], [
        ["RC0603FR0710KL", "SMD 0603 Package, 10k, 1%", "S", "RES"]
    ])
    output = tmp_path / "punct.xlsx"
    result = compare_workbooks(customer, erp, rules, output)
    # They should match without generating "Sistemde Yok"
    assert not any("Sistemde Yok" in row.get("Hata/Uyarı Durumu", "") for row in result.rows)


def test_qty_mismatch_reported(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Qty"], [
        ["SYN-001", "10k", 10]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Qty"], [
        ["SYN-001", "10k", "S", 5]
    ])
    output = tmp_path / "qty.xlsx"
    result = compare_workbooks(customer, erp, rules, output)
    row = next(r for r in result.rows if "SYN-001" in r["Part Number"])
    assert "Miktar Uyuşmazlığı" in row["Hata/Uyarı Durumu"]
    assert "Müşteri: 10" in row["Hata/Uyarı Durumu"]
    assert "ERP: 5" in row["Hata/Uyarı Durumu"]


def test_customer_th_smd_mismatch_with_erp(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Montaj Tipi"], [
        ["SYN-002", "Connector", "TH"]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD"], [
        ["SYN-002", "Connector", "S"]
    ])
    output = tmp_path / "mounting.xlsx"
    result = compare_workbooks(customer, erp, rules, output)
    row = next(r for r in result.rows if "SYN-002" in r["Part Number"])
    assert "Montaj Tipi Çelişkisi" in row["Hata/Uyarı Durumu"]


def test_semantic_equivalence_100nf_equals_0_1uf(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value"], [
        ["CAP-100N", "100nF 50V 0805"]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type"], [
        ["CAP-100N", "SMD 0805 Package, 0.1uF, 50V", "S", "CAP"]
    ])
    output = tmp_path / "semantic_eq.xlsx"
    result = compare_workbooks(customer, erp, rules, output)
    # Equivalent 100nF and 0.1uF with same 50V and 0805 should NOT trigger Açıklama Çelişkisi
    row = next((r for r in result.rows if "CAP-100N" in r["Part Number"]), None)
    if row:
        assert "Açıklama Çelişkisi" not in row.get("Hata/Uyarı Durumu", "")


def test_semantic_voltage_contradiction(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value"], [
        ["CAP-DIFF", "100nF 50V 0805"]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type"], [
        ["CAP-DIFF", "SMD 0805 Package, 100nF, 25V", "S", "CAP"]
    ])
    output = tmp_path / "volt_diff.xlsx"
    result = compare_workbooks(customer, erp, rules, output)
    row = next(r for r in result.rows if "CAP-DIFF" in r["Part Number"])
    assert "Voltaj farklı" in row["Hata/Uyarı Durumu"]


def test_refdes_pairing_for_different_part_numbers(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Designator", "Qty"], [
        ["SAMSUNG-CAP-0603", "2.2nF 50V 0603", "C7", 1]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part", "Qty"], [
        ["YAGEO-CAP-0805", "SMD 0805 Package, 2.2nF, 200V", "S", "CAP", "C7", 1]
    ])
    output = tmp_path / "refdes_pair.xlsx"
    result = compare_workbooks(customer, erp, rules, output)

    # Should be paired into a single row instead of two "Sistemde Yok" rows
    row = next((r for r in result.rows if "SAMSUNG-CAP-0603" in r["Part Number"]), None)
    assert row is not None
    assert "YAGEO-CAP-0805" in row["Part Number"]
    assert "Part Number farklı" in row["Hata/Uyarı Durumu"]
    assert "Sistemde Yok" not in row["Hata/Uyarı Durumu"]


def test_erp_multi_code_sibling_filtering(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Designator", "Qty"], [
        ["CL21B104KBCNNNC", "100nF 50V 0805", "C10", 1]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part", "Qty"], [
        ["CL21B104KBCNNNC 885012207098 CC0805KRX7R9BB104", "SMD 0805 Package, 100nF, 50V", "S", "CAP", "C10", 1]
    ])
    output = tmp_path / "multi_code.xlsx"
    result = compare_workbooks(customer, erp, rules, output)

    # Sibling codes should NOT produce false "Sistemde Yok" rows
    sistemde_yok = [r for r in result.rows if "Sistemde Yok" in r.get("Hata/Uyarı Durumu", "")]
    assert len(sistemde_yok) == 0


def test_refdes_pairing_without_issues_produces_clean_row(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Designator", "Qty"], [
        ["SAMSUNG-CAP-0805", "10uF 50V 0805", "C2", 1]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Part", "Qty"], [
        ["YAGEO-CAP-0805", "SMD 0805 Package, 10uF, 50V", "S", "CAP", "C2", 1]
    ])
    output = tmp_path / "clean_alt.xlsx"
    result = compare_workbooks(customer, erp, rules, output)

    # Identical technical specs on same RefDes means successful alternative - marked UYGUN without errors!
    row = next(r for r in result.rows if "SAMSUNG-CAP-0805" in r["Part Number"] or "YAGEO-CAP-0805" in r["Part Number"])
    assert row["Durum"] == "UYGUN"
    assert row["Hata/Uyarı Durumu"] == ""


def test_description_matching_when_customer_part_number_missing(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    # Müşteri dosyasında Part Number sütunu boş, fakat Value alanı ERP Description ile birebir aynı:
    _write_book(customer, "BOM", ["Part Number", "Value", "Qty"], [
        [None, "TH 2 pin connector", 3]
    ])
    _write_book(erp, "Malzemeler", ["Part Number", "Description", "TH/SMD", "Material Type", "Qty"], [
        ["CON-SPECIAL-01", "TH 2 pin connector", "T", "CON", 3]
    ])
    output = tmp_path / "desc_match.xlsx"
    result = compare_workbooks(customer, erp, rules, output)

    row = next(r for r in result.rows if "CON-SPECIAL-01" in r["Part Number"])
    assert "Part Number eksik" in row["Hata/Uyarı Durumu"]
    assert "Sistemde Yok" not in row["Hata/Uyarı Durumu"]
    assert str(row["Müşteri Qty"]) == "3"
    assert str(row["ERP Qty"]) == "3"


def test_customer_item_without_part_number_missing_in_erp(tmp_path):
    customer, erp, rules = _source_files(tmp_path)
    _write_book(customer, "BOM", ["Part Number", "Value", "Designator", "Qty"], [
        [None, "Özel Üretim Braket", "BRK1", 2]
    ])
    output = tmp_path / "cust_no_pn.xlsx"
    result = compare_workbooks(customer, erp, rules, output)

    row = next(r for r in result.rows if "BRK1" in r["Part Number"])
    assert "Sistemde Yok: ERP dosyasında bulunamadı" in row["Hata/Uyarı Durumu"]
    assert "Part Number eksik" in row["Hata/Uyarı Durumu"]
    assert row["Durum"] == "HATA"


def test_compare_workbooks_without_rules(tmp_path):
    customer, erp, _ = _source_files(tmp_path)
    output = tmp_path / "no_rules.xlsx"
    result = compare_workbooks(customer, erp, output_path=output)

    assert result.total_part_count == 5
    assert output.exists()
    # Comparison without rules should work and not fail with any missing rules error
    cust_row = next(r for r in result.rows if r["Part Number"] == "ONLY-CUSTOMER")
    assert "Sistemde Yok: ERP dosyasında bulunamadı" in cust_row["Hata/Uyarı Durumu"]
    con_row = next(r for r in result.rows if r["Part Number"] == "CON-003")
    assert "Montaj Tipi Çelişkisi" in con_row["Hata/Uyarı Durumu"]


def test_multi_customer_files_aggregated_against_single_erp(tmp_path):
    tr1_file = tmp_path / "TR1.xlsx"
    st1_file = tmp_path / "ST1.xlsx"
    erp_file = tmp_path / "erp.xlsx"

    _write_book(tr1_file, "BOM", ["Part Number", "Value", "Part", "Qty"], [
        ["AC0603KRX7R9BB104", "100nF 50V", "C10, C14, C45, C47", 4],
        ["CAP-01", "100nF 50V", "C7", 1],
    ])
    _write_book(st1_file, "BOM", ["Part Number", "Value", "Part", "Qty"], [
        ["AC0603KRX7R9BB104", "100nF 50V", "C18, C22", 2],
        ["CAP-02", "10uF 25V", "C7", 1],
    ])
    _write_book(erp_file, "ERP", ["Part Number", "Description", "Part", "Qty", "TH/SMD"], [
        ["AC0603KRX7R9BB104", "100nF 50V SMD", "TR1_C10, TR1_C14, TR1_C45, TR1_C47, ST1_C18, ST1_C22", 6, "SMD"],
        ["CAP-01", "100nF 50V SMD", "TR1_C7", 1, "SMD"],
        ["CAP-02", "10uF 25V SMD", "ST1_C7", 1, "SMD"],
    ])

    output = tmp_path / "multi_customer_out.xlsx"
    result = compare_workbooks([tr1_file, st1_file], erp_file, output_path=output)

    assert result.total_part_count == 3
    ac_row = next(r for r in result.rows if r["Part Number"] == "AC0603KRX7R9BB104")
    assert str(ac_row["Müşteri Qty"]) == "6"
    assert str(ac_row["ERP Qty"]) == "6"
    assert "Miktar Uyuşmazlığı" not in ac_row["Hata/Uyarı Durumu"]
    assert "yinelenen Part Number" not in ac_row["Hata/Uyarı Durumu"]
    assert ac_row["Durum"] == "UYGUN"

    cap1 = next(r for r in result.rows if r["Part Number"] == "CAP-01")
    assert cap1["Durum"] == "UYGUN"
    assert str(cap1["Müşteri Qty"]) == "1"

    cap2 = next(r for r in result.rows if r["Part Number"] == "CAP-02")
    assert cap2["Durum"] == "UYGUN"
    assert str(cap2["Müşteri Qty"]) == "1"


def test_multi_sheet_single_customer_workbook_aggregated(tmp_path):
    wb_file = tmp_path / "panel_customer.xlsx"
    erp_file = tmp_path / "erp_panel.xlsx"

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "TR1"
    ws1.append(["Part Number", "Value", "Part", "Qty"])
    ws1.append(["AC0603KRX7R9BB104", "100nF 50V", "C1, C2, C3", 3])

    ws2 = wb.create_sheet("ST1")
    ws2.append(["Part Number", "Value", "Part", "Qty"])
    ws2.append(["AC0603KRX7R9BB104", "100nF 50V", "C4, C5, C6, C7, C8", 5])

    wb.save(wb_file)
    wb.close()

    _write_book(erp_file, "ERP", ["Part Number", "Description", "Part", "Qty", "TH/SMD"], [
        ["AC0603KRX7R9BB104", "100nF 50V SMD", "TR1_C1, TR1_C2, TR1_C3, ST1_C4, ST1_C5, ST1_C6, ST1_C7, ST1_C8", 8, "SMD"]
    ])

    output = tmp_path / "multi_sheet_out.xlsx"
    result = compare_workbooks(wb_file, erp_file, output_path=output)

    assert result.total_part_count == 1
    row = result.rows[0]
    assert row["Part Number"] == "AC0603KRX7R9BB104"
    assert str(row["Müşteri Qty"]) == "8"
    assert str(row["ERP Qty"]) == "8"
    assert row["Durum"] == "UYGUN"





