"""Golden corpus regression suite for historical production datasets.

Tests all 5 historical production datasets:
1. Upper BOM (STOP TAIL 1/2, TURN 1/2 PCB sheets vs BOMT01)
2. Lower BOM (Multi-workbook customer package vs BOMT01)
3. Delta-V Orta (Altium BOM vs ERP BOM)
4. Delta-V Telemetri V2 (Altium BOM with malformed designator '1' vs ERP BOM)
5. DCP4 (Batch production BOM vs ERP BOM)

Ensures zero regressions, exact baseline counts, and clean issue cells for all UYGUN rows.
"""

import os
from pathlib import Path
import pytest
from bom_enricher.comparison import compare_workbooks

UPLOAD_DIR = Path(
    os.environ.get(
        "GOLDEN_CORPUS_DIR",
        "C:/Users/sudem/.gemini/antigravity/brain/212f92ee-c0ed-4409-b01a-8aa6bf1d8d4b/.user_uploaded",
    )
)


def _status_counts(comparison_result) -> dict[str, int]:
    counts: dict[str, int] = {}
    for row in comparison_result.rows:
        status = row.get("Durum", "")
        counts[status] = counts.get(status, 0) + 1
    return counts


def _verify_clean_cells(comparison_result) -> None:
    for idx, row in enumerate(comparison_result.rows):
        if row.get("Durum") == "UYGUN":
            issue = str(row.get("Hata/Uyarı Durumu") or "").strip()
            assert issue == "", f"Row {idx} is UYGUN but has non-empty issue cell: {issue!r}"


@pytest.mark.skipif(
    not (UPLOAD_DIR / "media_1792750235310_fed29c89.xlsx").exists()
    or not (UPLOAD_DIR / "media_1792750248966_63236d2c.xlsx").exists(),
    reason="Upper BOM golden files not found in user uploaded directory",
)
def test_golden_upper_bom():
    cust_file = UPLOAD_DIR / "media_1792750235310_fed29c89.xlsx"
    erp_file = UPLOAD_DIR / "media_1792750248966_63236d2c.xlsx"

    res = compare_workbooks(cust_file, erp_file)
    counts = _status_counts(res)

    assert counts.get("UYGUN") == 34, f"Expected 34 UYGUN, got {counts.get('UYGUN')}"
    assert counts.get("UYARI") == 5, f"Expected 5 UYARI, got {counts.get('UYARI')}"
    assert counts.get("HATA") == 4, f"Expected 4 HATA, got {counts.get('HATA')}"
    _verify_clean_cells(res)

    # Verify no drawing notes or 'See PCB Spec' rows were ingested
    for row in res.rows:
        val = str(row.get("Müşteri Adı / Değer") or "")
        pn = str(row.get("Müşteri Part Number") or "")
        assert "see pcb spec" not in val.lower(), f"Unexpected drawing note ingested: {val}"
        assert "see pcb spec" not in pn.lower(), f"Unexpected drawing note in PN: {pn}"


@pytest.mark.skipif(
    not (UPLOAD_DIR / "media_1792745893184_41e69dec.xlsx").exists()
    or not (UPLOAD_DIR / "media_1792745916907_ad8c26b6.xlsx").exists(),
    reason="Lower BOM golden files not found in user uploaded directory",
)
def test_golden_lower_bom():
    cust_files = [
        UPLOAD_DIR / "media_1792745916907_ad8c26b6.xlsx",
        UPLOAD_DIR / "media_1792745916907_2935931c.xlsx",
        UPLOAD_DIR / "media_1792745916907_2cec9e93.xlsx",
    ]
    erp_file = UPLOAD_DIR / "media_1792745893184_41e69dec.xlsx"

    res = compare_workbooks(cust_files, erp_file)
    counts = _status_counts(res)

    assert counts.get("UYGUN") == 71, f"Expected 71 UYGUN, got {counts.get('UYGUN')}"
    assert counts.get("UYARI") == 7, f"Expected 7 UYARI, got {counts.get('UYARI')}"
    assert counts.get("HATA") == 1, f"Expected 1 HATA, got {counts.get('HATA')}"
    _verify_clean_cells(res)

    # Verify the single genuine HATA is TR2_P1 prefix mismatch
    hata_rows = [r for r in res.rows if r.get("Durum") == "HATA"]
    assert len(hata_rows) == 1
    hata_row = hata_rows[0]
    assert "TR2_P1" in str(hata_row.get("Müşteri Referans") or "")
    assert "TR1_P1" in str(hata_row.get("ERP Referans") or "")


@pytest.mark.skipif(
    not (UPLOAD_DIR / "media_1793104524070_fdaaf120.xlsx").exists()
    or not (UPLOAD_DIR / "media_1793104537002_fc05c18a.xlsx").exists(),
    reason="Delta-V Orta golden files not found in user uploaded directory",
)
def test_golden_delta_v_orta():
    cust_file = UPLOAD_DIR / "media_1793104524070_fdaaf120.xlsx"
    erp_file = UPLOAD_DIR / "media_1793104537002_fc05c18a.xlsx"

    res = compare_workbooks(cust_file, erp_file)
    counts = _status_counts(res)

    assert counts.get("UYGUN") == 20, f"Expected 20 UYGUN, got {counts.get('UYGUN')}"
    assert counts.get("UYARI") == 1, f"Expected 1 UYARI, got {counts.get('UYARI')}"
    assert counts.get("HATA") == 2, f"Expected 2 HATA, got {counts.get('HATA')}"
    _verify_clean_cells(res)


@pytest.mark.skipif(
    not (UPLOAD_DIR / "media_1793105469871_388f4028.xlsx").exists()
    or not (UPLOAD_DIR / "media_1793105469871_0a773617.xlsx").exists(),
    reason="Delta-V Telemetri V2 golden files not found in user uploaded directory",
)
def test_golden_delta_v_telemetri_v2():
    cust_file = UPLOAD_DIR / "media_1793105469871_388f4028.xlsx"
    erp_file = UPLOAD_DIR / "media_1793105469871_0a773617.xlsx"

    res = compare_workbooks(cust_file, erp_file)
    counts = _status_counts(res)

    # 41 UYGUN, 2 UYARI (including Task 1 malformed '1' designator warning), 6 HATA
    assert counts.get("UYGUN") == 41, f"Expected 41 UYGUN, got {counts.get('UYGUN')}"
    assert counts.get("UYARI") == 2, f"Expected 2 UYARI, got {counts.get('UYARI')}"
    assert counts.get("HATA") == 6, f"Expected 6 HATA, got {counts.get('HATA')}"
    _verify_clean_cells(res)

    # Verify that the malformed '1' designator was caught as a warning
    malformed_warning_found = any(
        "Malformed reference designator '1'" in str(r.get("Hata/Uyarı Durumu") or "")
        or "missing component class letter" in str(r.get("Hata/Uyarı Durumu") or "")
        for r in res.rows
    )
    assert malformed_warning_found, "Expected malformed designator warning for '1' not found"


@pytest.mark.skipif(
    not (UPLOAD_DIR / "media_1793082034672_fa79893f.xlsx").exists()
    or not (UPLOAD_DIR / "media_1793082020014_0b7dd885.xlsx").exists(),
    reason="DCP4 golden files not found in user uploaded directory",
)
def test_golden_dcp4():
    cust_file = UPLOAD_DIR / "media_1793082034672_fa79893f.xlsx"
    erp_file = UPLOAD_DIR / "media_1793082020014_0b7dd885.xlsx"

    res = compare_workbooks(cust_file, erp_file)
    counts = _status_counts(res)

    assert counts.get("UYGUN") == 143, f"Expected 143 UYGUN, got {counts.get('UYGUN')}"
    assert counts.get("UYARI") == 78, f"Expected 78 UYARI, got {counts.get('UYARI')}"
    assert counts.get("HATA") == 4, f"Expected 4 HATA, got {counts.get('HATA')}"
    _verify_clean_cells(res)
