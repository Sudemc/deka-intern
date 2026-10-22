# Generic Multi-Code and Item/Ref Disambiguation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Disambiguate `Parça` (Item #) from `REFERANS` (RefDes) to eliminate ambiguous column errors and enable generic multi-code matching (Customer Internal Code vs. MPN) against ERP without hardcoding any company names.

**Architecture:** Robust Turkish diacritic normalization and cell-content inspection in `excel_io.py` distinguish row numbers from circuit references. Generic suffix/pattern matching identifies all code columns (`* kodu`, `* code`, `* no`) per row, and `comparison.py` cross-matches ERP Part Numbers against any candidate code belonging to that component.

**Tech Stack:** Python 3.10+, Pandas, OpenPyXL, unicodedata, pytest.

**Spec:** `docs/superpowers/specs/2026-10-22-generic-multi-code-and-item-ref-matching-design.md`

## Global Constraints

- Never hardcode any specific company name (e.g. `dtsis`, `deka`, `acme`) in logic or column mappings.
- Preserve backward compatibility with existing single-code customer BOMs and ERP formats.
- All existing tests in `tests/test_comparison.py` and `tests/test_excel_io.py` must pass.
- When reporting matches, provide full transparency by displaying both the matched code and the alternate code.

## Review Focus

1. Customer BOM has both `Parça` (integers) and `REFERANS` (`D13, D15`): must resolve `REFERANS` as `part` and `Parça` as item sequence without raising ambiguous column error.
2. ERP list uses Customer Internal Code (`H010101005006`): must match customer row containing `DTSIS KODU: H010101005006` and `ÜRETİCİ KODU: MBR0540T1G` with status `UYGUN`.
3. ERP list uses Manufacturer Code (`MBR0540T1G`): must also match the exact same customer row with status `UYGUN`.
4. Arbitrary company name in BOM (e.g., `ACME KODU: X1`, `ÜRETİCİ KODU: Y1`): must work generically without custom rules.
5. Turkish uppercase `İ` in header (`ÜRETİCİ KODU`): must normalize cleanly to `uretici kodu` without Unicode combining dot artifacts (`\u0307`).

---

### Task 1: Robust Turkish Header Normalization and Parça vs. Referans Disambiguation

**Files:**
- Modify: `src/bom_enricher/excel_io.py:27-108` and `161-205`
- Test: `tests/test_excel_io.py`

**Interfaces:**
- Consumes: `_normalize_header(value: Any) -> str`, `_HEADER_ALIASES`
- Produces: Correctly classified headers where `referans` takes precedence over `parça` for circuit references, and Turkish diacritics are normalized deterministically.

- [ ] **Step 1: Write the failing tests**

```python
def test_turkish_header_normalization():
    from bom_enricher.excel_io import _header_field, _normalize_header
    assert _normalize_header("ÜRETİCİ KODU") == "uretici kodu"
    assert _header_field("ÜRETİCİ KODU") == "part_number"
    assert _header_field("REFERANS") == "part"
    assert _header_field("DEĞER") == "value"

def test_parca_and_referans_disambiguation(tmp_path):
    from bom_enricher.excel_io import read_customer_bom
    import pandas as pd
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_excel_io.py::test_turkish_header_normalization tests/test_excel_io.py::test_parca_and_referans_disambiguation -v`  
Expected: FAIL with `AssertionError` or `ValueError: has ambiguous column(s): Parça`.

- [ ] **Step 3: Implement Turkish normalization & RefDes disambiguation in `src/bom_enricher/excel_io.py`**

1. Replace `_normalize_header` with diacritic-stripping NFKD normalization.
2. In `_HEADER_ALIASES`, define `item` alias (`parca`, `parça`, `item`, `kalem no`, `sira no`, `line no`). Remove `parca`, `parça` from `part`.
3. In `_auto_layout`, if multiple columns map to `part` or if both `item` and `part` are present, inspect sample rows using `_REFERENCE_PATTERN` to ensure the true schematic designator column is chosen.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_excel_io.py::test_turkish_header_normalization tests/test_excel_io.py::test_parca_and_referans_disambiguation -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/excel_io.py tests/test_excel_io.py
git commit -m "feat(excel_io): add robust Turkish normalization and disambiguate Parça from RefDes"
```

---

### Task 2: Generic Multi-Code Column Detection in Customer BOMs

**Files:**
- Modify: `src/bom_enricher/excel_io.py:100-108`, `170-230`, `320-396`
- Test: `tests/test_excel_io.py`

**Interfaces:**
- Consumes: Raw sheet headers and sample row values.
- Produces: `CustomerBomRow.product_code_candidates` containing all codes found across any detected code columns (e.g. `* KODU`, `* CODE`, `* NO`, `* NUMBER`, `* PN`).

- [ ] **Step 1: Write the failing test**

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_excel_io.py::test_generic_multi_code_candidate_detection -v`  
Expected: FAIL because `DTSIS KODU` is not yet recognized as a code column.

- [ ] **Step 3: Implement generic code column detection in `src/bom_enricher/excel_io.py`**

1. In `_header_field`: recognize any header ending in `kodu`, `kod`, `code`, `number`, `no`, `pn`, `p n`, `mpn` as `part_number` (unless it is a pure row/line item header like `kalem no`, `sira no`, `line no`).
2. In `_auto_layout`: collect all columns mapped to `part_number` into `code_columns`.
3. In `read_customer_bom`: extract non-empty candidate values from all `code_columns` into `product_code_candidates`, tagging each candidate with its source header title.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_excel_io.py::test_generic_multi_code_candidate_detection -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/excel_io.py tests/test_excel_io.py
git commit -m "feat(excel_io): generically detect all product code columns without company-specific names"
```

---

### Task 3: Multi-Code PartRecord and Cross-Matching in `comparison.py`

**Files:**
- Modify: `src/bom_enricher/comparison.py:58-75`, `470-575`, `730-790`
- Test: `tests/test_comparison.py`

**Interfaces:**
- Consumes: `CustomerBomRow.product_code_candidates`
- Produces: `PartRecord.all_codes`, `PartRecord.code_labels`, and Phase 1 cross-matching against any candidate code.

- [ ] **Step 1: Write the failing tests**

```python
def test_cross_matching_with_customer_internal_code_in_erp(tmp_path):
    from bom_enricher.comparison import compare_workbooks
    import pandas as pd

    cust_path = tmp_path / "customer.xlsx"
    pd.DataFrame({
        "Parça": [37],
        "DTSIS KODU": ["H010101005006"],
        "ÜRETİCİ KODU": ["MBR0540T1G"],
        "REFERANS": ["D13,D15"],
        "MIKTAR": [2],
        "DEĞER": ["MBR0540"],
    }).to_excel(cust_path, index=False)

    erp_path = tmp_path / "erp.xlsx"
    pd.DataFrame({
        "Part Number": ["H010101005006"],
        "Description": ["MBR0540 SMD DIYOT"],
        "Reference": ["D13,D15"],
        "Qty": [2],
    }).to_excel(erp_path, index=False)

    res = compare_workbooks(cust_path, erp_path)
    assert len(res.rows) == 1
    row = res.rows[0]
    assert row["Durum"] == "UYGUN"
    assert "H010101005006" in row["Müşteri Part Number"]
    assert "MBR0540T1G" in row["Müşteri Part Number"]
    assert row["ERP Part Number"] == "H010101005006"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_comparison.py::test_cross_matching_with_customer_internal_code_in_erp -v`  
Expected: FAIL.

- [ ] **Step 3: Implement multi-code matching and reporting in `src/bom_enricher/comparison.py`**

1. Extend `PartRecord` with `all_codes: tuple[str, ...] = ()` and `code_labels: dict[str, str] = field(default_factory=dict)`.
2. In `load_records` (`kind="customer"`): populate `all_codes` with all candidates from `row.product_code_candidates`. Do not raise an error when multiple candidates exist if they originate from different code columns for the same component.
3. In `compare_workbooks`: in Phase 1 Part Number matching, index customer records by ALL of their candidate codes (normalized). If ERP part number matches ANY of these codes, match the pair as Phase 1 match!
4. In report generation: if multiple codes exist and ERP matched one, format `Müşteri Part Number` cleanly as `<matched_code> [<other_label>: <other_code>]`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_comparison.py::test_cross_matching_with_customer_internal_code_in_erp -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/comparison.py tests/test_comparison.py
git commit -m "feat(comparison): cross-match ERP part numbers against internal and manufacturer candidate codes"
```

---

### Task 4: End-to-End Verification on Real Data Sample

**Files:**
- Test: `tests/test_comparison.py`

- [ ] **Step 1: Write an end-to-end test using the user's exact provided BOM snippet**

```python
def test_real_user_dtsis_bom_comparison_snippet(tmp_path):
    from bom_enricher.comparison import compare_workbooks
    import pandas as pd
    import io

    csv_data = """Parça,FULL,MIKTAR,DTSIS KODU,ÜRETİCİ KODU,REFERANS,DEĞER,KILIF,MontajTipi,Revised: 03.09.2026
37,+,2,H010101005006,MBR0540T1G,"D13,D15",MBR0540,smdmbr0540,SMD,
44,+,8,H010101005052,SD12CT1G,"D29,D30,D34,D35,D41,D43,D44,D46",SD12CT1G,SOD-323_SD12CT,SMD,
"""
    df_cust = pd.read_csv(io.StringIO(csv_data))
    cust_path = tmp_path / "user_cust.xlsx"
    df_cust.to_excel(cust_path, index=False)

    df_erp = pd.DataFrame({
        "Part Number": ["H010101005006", "H010101005052"],
        "Description": ["MBR0540", "SD12CT1G"],
        "Reference": ["D13,D15", "D29,D30,D34,D35,D41,D43,D44,D46"],
        "Qty": [2, 8],
    })
    erp_path = tmp_path / "user_erp.xlsx"
    df_erp.to_excel(erp_path, index=False)

    res = compare_workbooks(cust_path, erp_path)
    assert len(res.rows) == 2
    assert all(r["Durum"] == "UYGUN" for r in res.rows)
```

- [ ] **Step 2: Run all comparison and Excel IO tests to verify complete suite health**

Run: `.venv\Scripts\pytest tests/test_comparison.py tests/test_excel_io.py -v`  
Expected: ALL PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/test_comparison.py
git commit -m "test: add real customer BOM snippet end-to-end verification"
```
