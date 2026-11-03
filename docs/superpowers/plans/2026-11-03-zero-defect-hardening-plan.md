# Zero-Defect BOM Comparison & Enrichment Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Eliminate all remaining parsing, reference-range expansion, multi-row reconciliation, semantic comparison, datasheet analysis, and web/UI vulnerabilities across the entire `bom_enricher` codebase so that any customer or ERP BOM format is processed with zero defects.

**Architecture:** We harden the pipeline in six decoupled layers: (1) Reference range expansion & multi-underscore designator normalization across `excel_io.py` and `comparison.py`; (2) Universal quantity parsing, DNP/NM immunity even when MPNs are present, and multi-sheet ingestion guards; (3) Customer multi-code splitting, expanded designator pairing (`LED`, `SW`, `FB`, `IC`, etc.), and bidirectional multi-row reconciliation; (4) Context-aware parametric & category/color/polarity comparison in `semantic_parser.py`; (5) Datasheet analyzer spec-key alignment, resource leak fixes, and `validation.py` synchronization; and (6) Web backend exception/temp-file safety and frontend table/filter completeness.

**Tech Stack:** Python 3.10+, `openpyxl`, `pypdf`/`pypdfium2`, `FastAPI`, vanilla JS/HTML5 (`static/index.html`), `pytest`.

**Spec:** Derived from the comprehensive full-project audit conducted on 2026-11-03.

## Global Constraints

- **Zero Regression on Golden Corpus:** All existing 246 tests (including Golden Corpus Upper, Lower, Delta-V Orta, Delta-V Telemetri V2, DCP4, and FEKA Merve) must continue to pass with identical status counts.
- **Strict Clean Cell Rule:** Every row with `Durum == "UYGUN"` MUST have `"Hata/Uyarı Durumu": ""` (empty string).
- **Backward-Compatible Signatures:** Public functions (`compare_workbooks`, `read_customer_bom`, `read_bom_workbook`, `parse_specs`, `compare_descriptions`, `analyze_single_datasheet`) must keep backward-compatible call signatures.

## Review Focus

- **Range designators with board prefixes or spaces (`R1 - R5`, `C10..C15`, `LED1~LED4` on sheet `TR1`):** Must expand into individual prefixed tokens (`TR1_R1 .. TR1_R5`) without turning `R1-R5` into `R1_R5` or dropping `..`/`~` ranges.
- **Active MPN shared with a DNP row on the same board:** A row marked `DNP` / `NM` / `Qty=0` in Description, Mounting, Status, or RefDes (`R15 (DNP)`) that carries a valid MPN must not be merged into active rows of that MPN or flagged as `HATA: ERP'de Eksik Kalem`.
- **JEDEC part numbers (`1N4148`, `2N7002`) and capacitor R-codes (`4u7`, `4n7`) vs Megaohm resistors (`10M`, `10MOhm`):** Must never parse diodes/capacitors as inductances, and must parse uppercase `M`/`MOhm` as $10^6\,\Omega$ while keeping lowercase `mOhm`/`mR` as $10^{-3}\,\Omega$.
- **Cross-category, color, polarity, and `mpn_matched=True` value mismatches (`100nF` vs `100R`, `Red LED` vs `Yellow LED`, `NPN` vs `PNP`, `4.7K` vs `10K`):** Must be deterministically flagged by `compare_descriptions`.
- **Corrupted/encrypted `.xlsx` and `>20MB` PDF uploads in `web.py`:** Must never leak temporary directories on disk, crash with unhandled 500 tracebacks, or silently truncate PDF streams.

---

### Task 1: Reference Range & Designator Pipeline Hardening

**Files:**
- Modify: `src/bom_enricher/excel_io.py:336-455`
- Modify: `src/bom_enricher/comparison.py:277-315, 384-415, 610-650, 1290-1330, 1965-1990, 2350-2385`
- Test: `tests/test_excel_io.py`, `tests/test_comparison.py`

**Interfaces:**
- Produces:
  - `expand_ref_range(raw: str) -> list[str]` (supports `TR1_C10..C15`, `TR1_C10-15`, `R1-R5`, `LED1~LED4`)
  - `split_references(text: str) -> list[str]` (normalizes spaces around `-`, `..`, `~` before splitting)
  - `_normalize_ref_token(token: str) -> str` (does not convert range `R1-R5` into `R1_R5`)
  - `_ref_base_component(ref: str) -> tuple[str, str, str]` (uses last underscore when suffix is a valid designator so `FOG_L_D1` -> `("FOG_L", "D1", "D")`)

- [ ] **Step 1: Write the failing tests in `tests/test_comparison.py`**

```python
def test_range_designators_with_board_prefix_and_spaces(tmp_path):
    cust_file = tmp_path / "TR1_BOM.xlsx"
    erp_file = tmp_path / "ERP.xlsx"
    _write_book(cust_file, "TR1", ["Part Number", "Value", "RefDes", "Qty"], [
        ["MPN-RES1", "10k", "R1 - R5", 5],
        ["MPN-CAP1", "100nF", "C10..C12", 3],
        ["MPN-LED1", "Red LED", "LED1~LED3", 3],
    ])
    _write_book(erp_file, "ERP", ["Part Number", "Description", "Reference", "Qty", "Mounting"], [
        ["MPN-RES1", "10k SMD", "TR1_R1, TR1_R2, TR1_R3, TR1_R4, TR1_R5", 5, "SMD"],
        ["MPN-CAP1", "100nF SMD", "TR1_C10, TR1_C11, TR1_C12", 3, "SMD"],
        ["MPN-LED1", "Red LED SMD", "TR1_LED1, TR1_LED2, TR1_LED3", 3, "SMD"],
    ])
    res = compare_workbooks(cust_file, erp_file)
    for row in res.rows:
        assert row["Durum"] == "UYGUN", f"Unexpected issue on {row['Part Number']}: {row['Hata/Uyarı Durumu']}"
        assert row["Hata/Uyarı Durumu"] == ""


def test_multi_underscore_prefix_base_extraction_and_material_inference():
    from bom_enricher.comparison import _ref_base_component, _infer_material_type, PartRecord
    assert _ref_base_component("FOG_L_D1") == ("FOG_L", "D1", "D")
    assert _ref_base_component("ST_LH_R12") == ("ST_LH", "R12", "R")
    rec_cap = PartRecord(part_number="X", description="100nF", mounting="SMD", reference="REV1_C10", material_type="", qty="1")
    rec_res = PartRecord(part_number="Y", description="10K", mounting="SMD", reference="DRL1_R5", material_type="", qty="1")
    assert _infer_material_type(rec_cap, None) == "cap"
    assert _infer_material_type(rec_res, None) == "res"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -k "test_range_designators_with_board_prefix_and_spaces or test_multi_underscore_prefix_base_extraction" -v`
Expected: FAIL (`R1 - R5` fails Qty/RefDes count; `C10..C12` becomes `TR1_C10..C12` and drops all refs; `_ref_base_component("FOG_L_D1")` returns `("FOG", "L_D1", "L")`; `_infer_material_type` returns `"res"` for `REV1_C10`).

- [ ] **Step 3: Implement range expansion & designator fixes in `src/bom_enricher/excel_io.py` and `src/bom_enricher/comparison.py`**

1. In `excel_io.py:expand_ref_range`:
   - Match `r"^([A-Za-z0-9_]*?[A-Za-z]+)_?(\d+)\s*(?:\.\.|~|-|–|—)\s*([A-Za-z0-9_]*?[A-Za-z]+)?_?(\d+)$"`.
   - If `p2` is present without board prefix while `p1` has a board prefix (`p1 = "TR1_C"`, `p2 = "C"`), check `p1.split("_")[-1].upper() == p2.upper()` and preserve `p1`'s full prefix.
2. In `excel_io.py:split_references` and `validate_reference_tokens`:
   - Normalize whitespace around range operators (`..`, `~`, `-`, `–`, `—`) between designator tokens before splitting on whitespace.
   - Expand ranges via `expand_ref_range` in `validate_reference_tokens` before testing `DESIGNATOR_REGEX`.
3. In `comparison.py:_normalize_ref_token`:
   - If token matches a range pattern (`expand_ref_range(tok)` returns `> 1` items), do not rewrite `-` or `.` to `_`.
4. In `comparison.py:_qualify_customer_references`:
   - Expand any range tokens via `split_references` first before prefixing each designator token, and check prefix qualification per token rather than skipping the entire row when `any(...)` is already prefixed.
5. In `comparison.py:_ref_base_component`, `_extract_erp_prefixes`, and `_infer_material_type`:
   - Split on the last `_` (`norm.rsplit("_", 1)`) when the right part matches `^[A-Z]+\d+$`.
   - Use `_split_refs` for Qty-vs-RefDes count validation at lines 1294, 1313, and 1968.
   - In `_infer_material_type`, extract `_, base_ref, comp_type = _ref_base_component(segment)` and classify based on `comp_type.casefold()` instead of the board prefix.

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py tests/test_excel_io.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/excel_io.py src/bom_enricher/comparison.py tests/test_comparison.py
git commit -m "fix(refs): harden range designator expansion and multi-underscore prefix parsing"
```

---

### Task 2: Quantity Normalization, DNP/NM Immunity & Multi-Sheet Ingestion Guard

**Files:**
- Modify: `src/bom_enricher/excel_io.py:134-199, 425-455, 850-1230`
- Modify: `src/bom_enricher/comparison.py:560-575, 665-715, 1280-1330, 1510-1540, 1934-2025, 2185-2205`
- Test: `tests/test_excel_io.py`, `tests/test_comparison.py`

**Interfaces:**
- Produces:
  - `parse_numeric_qty(raw: object) -> float | None` in `excel_io.py` (parses `"12"`, `"12.0"`, `"4,0"`, `"0,05"`, `"10 ADET"`, `"5 PCS"`, `"2 EA"` into float)
  - `format_qty_value(val: float) -> str` (returns `"12"` for `12.0`, `"0.05"` for `0.05`)
  - `is_dnp_row(part_number: str, description: str, mounting: str, status: str, reference: str, qty: str) -> tuple[bool, str]`

- [ ] **Step 1: Write the failing tests in `tests/test_comparison.py` and `tests/test_excel_io.py`**

```python
def test_comma_decimal_and_unit_quantities_and_fractional_preservation(tmp_path):
    cust_file = tmp_path / "cust_qty.xlsx"
    erp_file = tmp_path / "erp_qty.xlsx"
    _write_book(cust_file, "BOM", ["Part Number", "Value", "RefDes", "Qty"], [
        ["MPN-1", "10k", "R1, R2, R3, R4", "4,0 ADET"],
        ["MPN-TAPE", "Kapton Tape", "TAPE1", "0,25"],
    ])
    _write_book(erp_file, "ERP", ["Part Number", "Description", "Reference", "Qty", "Mounting"], [
        ["MPN-1", "10k SMD", "R1, R2, R3, R4", "4 PCS", "SMD"],
        ["MPN-TAPE", "Kapton Tape", "TAPE1", "0.25", "SMD"],
    ])
    res = compare_workbooks(cust_file, erp_file)
    for r in res.rows:
        assert r["Durum"] == "UYGUN", f"Row {r['Part Number']} failed: {r['Hata/Uyarı Durumu']}"


def test_dnp_row_with_valid_mpn_not_merged_into_active_or_flagged_as_error(tmp_path):
    cust_file = tmp_path / "cust_dnp.xlsx"
    erp_file = tmp_path / "erp_dnp.xlsx"
    _write_book(cust_file, "BOM", ["Part Number", "Value", "RefDes", "Qty"], [
        ["RC0402FR-0710KL", "10K 0402", "R1, R2", 2],
        ["RC0402FR-0710KL", "10K 0402 (DNP)", "R3 (DNP)", 1],
        ["CL05B104KO5NNNC", "100nF (NOT MOUNTED)", "C99", 1],
    ])
    _write_book(erp_file, "ERP", ["Part Number", "Description", "Reference", "Qty", "Mounting"], [
        ["RC0402FR-0710KL", "10K 0402 SMD", "R1, R2", 2, "SMD"],
    ])
    res = compare_workbooks(cust_file, erp_file)
    active_row = next(r for r in res.rows if r["ERP Part Number"] == "RC0402FR-0710KL")
    assert active_row["Durum"] == "UYGUN"
    assert active_row["Hata/Uyarı Durumu"] == ""
    dnp_rows = [r for r in res.rows if r["ERP Part Number"] == ""]
    assert all(r["Durum"] == "UYARI" and "Montaj Dışı" in r["Hata/Uyarı Durumu"] for r in dnp_rows)


def test_multi_sheet_ignores_cover_page_with_project_code_and_keeps_cabin_pcb(tmp_path):
    from openpyxl import Workbook
    from bom_enricher.excel_io import find_customer_bom_sheets
    wb_path = tmp_path / "multi.xlsx"
    wb = Workbook()
    ws_cover = wb.active
    ws_cover.title = "Genel Bilgi"
    ws_cover.append(["Proje Kodu", "PRJ-2026-01"])
    ws_cover.append(["Revizyon", "1.0"])
    ws_bom = wb.create_sheet("CABIN_PCB")
    ws_bom.append(["Part Number", "Description", "Designator", "Quantity"])
    ws_bom.append(["RC0402FR-0710KL", "10K", "R1", 1])
    wb.save(wb_path)
    wb.close()
    assert find_customer_bom_sheets(wb_path) == ["CABIN_PCB"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -k "test_comma_decimal_and_unit_quantities or test_dnp_row_with_valid_mpn or test_multi_sheet_ignores_cover_page" -v`
Expected: FAIL

- [ ] **Step 3: Implement quantity normalization, DNP separation, and sheet selection guards**

1. Add `parse_numeric_qty` and `format_qty_value` in `excel_io.py` and use them in `comparison.py` (`_unique_records`, `compare_workbooks`, `_qualify_customer_references`) and `excel_io.py` (`read_customer_bom`). Preserve fractional values (`float` vs `int`) and keep numeric `total_qty` even if a duplicate row has a non-numeric note.
2. Detect `(DNP)`, `(NM)`, `(NC)`, `NOT MOUNTED`, `DO NOT POPULATE`, `DİZİLMEZ`, `NO FIT` in `Description`/`Value`, `Mounting`, `Status`, `Part Number`, or `RefDes` (before stripping parentheticals in `split_references`), as well as explicit `Qty == 0`. Set `status="DNP"` on those rows so `_unique_records` (`comparison.py:1943`) assigns them a separate `__DNP_` key instead of merging them into active rows, and allow `_check_certain_non_assembly` (`comparison.py:697`) to classify them as `"Montaj Dışı"` (`UYARI`) even when `has_code` is `True`.
3. In `excel_io.py:find_customer_bom_sheets` and `_select_customer_sheet`:
   - Match `_IGNORED_SHEET_KEYWORDS` on word/token boundaries (`re.split(r"[^a-z0-9]+", title_lower)`) so `"cabin_pcb"`, `"combined_bom"`, `"data_pcb"`, `"special_bom"` are never ignored.
   - Require `find_customer_bom_sheets` to verify `_is_bom_header_row` (at least 2 core BOM fields on the same row) before scoring a sheet as a BOM sheet.
   - Apply the blank/noise row filter in `read_customer_bom` Path C (`_auto_layout`) and `comparison.py:read_bom_workbook` (require `has_ref` for no-code rows, not just any digit in `qty`).
   - Normalize Turkish `'ı'` -> `'i'` in `comparison.py:_header_key` and `_normalize_label`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py tests/test_excel_io.py tests/test_golden_corpus.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/excel_io.py src/bom_enricher/comparison.py tests/test_comparison.py
git commit -m "fix(ingestion): normalize unit/comma quantities, isolate DNP rows with MPNs, and harden sheet detection"
```

---

### Task 3: Multi-Code Customer MPNs, Inverse Multi-Row Reconciliation & Pairing Completeness

**Files:**
- Modify: `src/bom_enricher/excel_io.py:255-275`
- Modify: `src/bom_enricher/comparison.py:315-382, 900-1170, 1455-1505`
- Test: `tests/test_comparison.py`

**Interfaces:**
- Modifies:
  - `_should_pair_refs(crefs: set[str], erefs: set[str]) -> bool` (supports all multi-letter designator types: `LED`, `SW`, `FB`, `IC`, `CN`, `CON`, `P`, `X`, `Y`, `T`, `TR`, `K`, `RL`, `B`, `VR`, `ZD`, etc.)
  - `compare_workbooks` Pass 1 & Pass 2 multi-row reconciliation (supports both `c_q < e_q` and `e_q < c_q` splits, customer multi-code cells, and immediate ERP sibling removal)

- [ ] **Step 1: Write the failing tests in `tests/test_comparison.py`**

```python
def test_led_and_sw_refdes_pairing_and_inverse_erp_row_split(tmp_path):
    cust_file = tmp_path / "cust_split.xlsx"
    erp_file = tmp_path / "erp_split.xlsx"
    _write_book(cust_file, "BOM", ["Part Number", "Value", "RefDes", "Qty"], [
        ["LA-G6SP / LS-G6SP", "Amber LED", "LED1, LED2, LED3, LED4", 4],
        ["CAP-100N-A", "100nF 50V 0402", "C1, C2, C3, C4", 4],
    ])
    _write_book(erp_file, "ERP", ["Part Number", "Description", "Reference", "Qty", "Mounting"], [
        ["LS-G6SP", "Amber LED SMD", "LED1, LED2, LED3, LED4", 4, "SMD"],
        ["CAP-100N-A", "100nF 50V 0402 SMD", "C1, C2", 2, "SMD"],
        ["CAP-100N-B", "100nF 50V 0402 SMD", "C3, C4", 2, "SMD"],
    ])
    res = compare_workbooks(cust_file, erp_file)
    led_row = next(r for r in res.rows if "LS-G6SP" in r["ERP Part Number"])
    assert led_row["Durum"] == "UYGUN"
    assert led_row["Hata/Uyarı Durumu"] == ""
    cap_row = next(r for r in res.rows if "CAP-100N" in r["Part Number"])
    assert cap_row["Durum"] in ("UYGUN", "UYARI")
    assert "Miktar Uyuşmazlığı" not in cap_row["Hata/Uyarı Durumu"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -k "test_led_and_sw_refdes_pairing_and_inverse_erp_row_split" -v`
Expected: FAIL (`LA-G6SP / LS-G6SP` is not split on customer side, `"LED" in "CRLDQUJSF"` is False, and ERP splitting `C1..C4` across two rows causes `Miktar Uyuşmazlığı (4 vs 2)`).

- [ ] **Step 3: Implement multi-code customer splitting, `_should_pair_refs` type set, and bidirectional multi-row aggregation**

1. In `excel_io.py:clean_part_number`:
   - Only extract inner parenthetical in `m_cust` (`line 261`) if the inner token is at least 5 characters and looks like a genuine MPN (`_code_like`), so `"FUSE (5A)"` or `"HEADER (2X5)"` is not reduced to `"5A"` or `"2X5"`.
   - Before stripping trailing `(...)` on line 268, ensure `read_customer_bom` extracts any valid code inside `(...)` into `product_code_candidates`.
2. In `comparison.py:read_bom_workbook`:
   - Split customer `part_number` and `product_code_candidates` using `_split_multi_code` so slash/comma/space-separated alternative customer MPNs populate `all_codes`.
3. In `comparison.py:_should_pair_refs`:
   - Replace `list(c_types)[0] in "CRLDQUJSF"` with `list(c_types)[0] in _PAIRABLE_COMP_TYPES` containing all standard single- and multi-letter designator prefixes (`C`, `R`, `L`, `D`, `Q`, `U`, `J`, `S`, `F`, `LED`, `LD`, `SW`, `FB`, `IC`, `CN`, `CON`, `P`, `X`, `Y`, `T`, `TR`, `K`, `RL`, `B`, `BT`, `VR`, `RV`, `ZD`, `DZ`, `TP`, `JP`, `TH`, `MOV`, `LS`, `BZ`).
4. In `comparison.py:compare_workbooks`:
   - When an ERP sibling code matches in Pass 1 or Pass 2, immediately mark its `sibling_group` as matched and remove sibling clones from `unmatched_erp`.
   - Support inverse multi-row aggregation (`e_q < c_q`) when remaining customer references (`crefs - erefs`) are covered by another unmatched ERP row with compatible specs, as well as Customer multi-row splits (`c_q < e_q`) when `other_c` has compatible specs and its `o_refs` are a subset of `missing_refs`.
   - In `_find_ref_near_misses`, require single-ref components (`len(c_refs) == 1`) to have numeric index difference $\le 2$ (e.g. `LED1` vs `LED2`) and matching `cp == ep`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py tests/test_golden_corpus.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/excel_io.py src/bom_enricher/comparison.py tests/test_comparison.py
git commit -m "feat(matching): support customer multi-codes, multi-letter ref pairing, and bidirectional row splits"
```

---

### Task 4: Semantic Parser Overhaul (`semantic_parser.py`)

**Files:**
- Modify: `src/bom_enricher/semantic_parser.py:1-516`
- Test: `tests/test_semantic_parser.py`

**Interfaces:**
- Modifies:
  - `MultiElectricalSpecs` & `ElectricalSpecs`: add optional fields with defaults (`frequencies: frozenset[Decimal] = frozenset()`, `currents: frozenset[Decimal] = frozenset()`, `powers: frozenset[Decimal] = frozenset()`, `dielectrics: frozenset[str] = frozenset()`, `categories: frozenset[str] = frozenset()`, `colors: frozenset[str] = frozenset()`, `polarities: frozenset[str] = frozenset()`)
  - `compare_descriptions(customer_text: str, erp_text: str, mpn_matched: bool = False) -> list[str]`

- [ ] **Step 1: Write the failing tests in `tests/test_semantic_parser.py`**

```python
def test_jedec_diodes_and_cap_rcodes_do_not_parse_as_inductance():
    from bom_enricher.semantic_parser import parse_specs_set, compare_descriptions
    s_diode = parse_specs_set("1N4148")
    assert not s_diode.capacitances
    assert not s_diode.inductances
    s_cap = parse_specs_set("4u7 16V 0603")
    assert s_cap.capacitances and not s_cap.inductances
    assert compare_descriptions("1N4148", "1N4007 Rectifier Diode", mpn_matched=False) == []


def test_megaohm_vs_milliohm_and_resistor_arrays():
    from decimal import Decimal
    from bom_enricher.semantic_parser import parse_specs_set
    assert Decimal("10000000") in parse_specs_set("10M").resistances
    assert Decimal("10000000") in parse_specs_set("10MOhm").resistances
    assert Decimal("0.01") in parse_specs_set("10mOhm").resistances
    assert Decimal("10000") in parse_specs_set("4x10K 0603").resistances


def test_cross_category_color_polarity_frequency_and_mpn_matched_value_conflicts():
    from bom_enricher.semantic_parser import compare_descriptions
    assert compare_descriptions("100nF", "RES SMD 100R 1% 0402", mpn_matched=False)
    assert compare_descriptions("Red LED 0603", "Yellow LED 0603", mpn_matched=False)
    assert compare_descriptions("NPN SOT-23", "PNP SOT-23", mpn_matched=False)
    assert compare_descriptions("8MHz", "16MHz Crystal", mpn_matched=False)
    assert compare_descriptions("4.7K 0603", "SMD 0603, 10k, %1, Yageo", mpn_matched=True)
    assert compare_descriptions("100nF 0402", "SMD 0805 Package, 100nF", mpn_matched=True)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.\.venv\Scripts\pytest.exe tests/test_semantic_parser.py -k "test_jedec_diodes or test_megaohm_vs_milliohm or test_cross_category_color_polarity" -v`
Expected: FAIL

- [ ] **Step 3: Implement semantic parser fixes in `src/bom_enricher/semantic_parser.py`**

1. **Remove optional `h?` in inductance R-code regex (`L240, L254`)**: Require explicit `h` (`r"(?<![a-z0-9])(\d+)(u|µ|\u03bc|m|n)h(\d+)(?![a-z0-9])"`) so `4u7`, `4n7`, `2m2` are never parsed as inductances.
2. **Exclude JEDEC codes (`1N\d{3,4}`, `2N\d{3,4}`, `3N\d{3,4}`) before capacitance/inductance R-code matching**: Strip or guard against `r"\b[123]n\d{3,4}[a-z]?\b"` so `1N4148` is not parsed as `1.4148nF`.
3. **Preserve `M` vs `m` before `.casefold()`**: Convert uppercase `(\d+(?:\.\d+)?)\s*(?:MOhm|MΩ|Meg|M)(?![a-zA-Z0-9])` to `\1megohm` before `.casefold()`, and include `megohm|meg` in `_RES_MULTIPLIERS` (`1e6`) and the resistance regex. Support array prefix `(?:\b\d+\s*[xX*]\s*)?` before resistance/capacitance values.
4. **Extract and compare additional attributes in `compare_descriptions`**:
   - **Cross-category conflict**: If one side has capacitance (and no resistance/inductance) while the other side has resistance or inductance (and no capacitance) — and neither is an RC network or ferrite bead with both — report `"Açıklama Çelişkisi (Malzeme Tipi farklı...)"`.
   - **Frequency** (`Hz`, `kHz`, `MHz`, `GHz`), **Color** (`Red/Kırmızı`, `Green/Yeşil`, `Yellow/Sarı`, `Blue/Mavi`, `White/Beyaz`, `Amber`, `Orange/Turuncu`), and **Polarity/Channel** (`NPN` vs `PNP`, `N-Channel` vs `P-Channel`).
   - **When `mpn_matched=True`**: Flag resistance/capacitance/inductance mismatches even when `ratio < 10`, and flag explicit package conflicts (`0402` vs `0805`).

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_semantic_parser.py tests/test_comparison.py tests/test_golden_corpus.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/semantic_parser.py tests/test_semantic_parser.py
git commit -m "feat(semantic): fix JEDEC/R-code/Megaohm bugs and add category, frequency, color, and polarity checks"
```

---

### Task 5: Datasheet Analyzer, PDF Renderer, Vision Client & Legacy Validation Sync

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:320-335, 4905-4925, 5255-5265, 5478-5492`
- Modify: `src/bom_enricher/pdf_renderer.py:135-160`
- Modify: `src/bom_enricher/vision_client.py:40-55, 250-260, 340-390`
- Modify: `src/bom_enricher/validation.py:65-75, 295-315`
- Modify: `src/bom_enricher/main.py:38-45`
- Test: `tests/test_datasheet_analyzer.py`, `tests/test_vision_client.py`, `tests/test_validation.py`

**Interfaces:**
- Modifies:
  - `CRITICAL_SPECS_BY_MATERIAL_TYPE` in `datasheet_analyzer.py` (keys `"REL"`, `"SWT"`, `"TRA"`, `"FUS"` match extracted spec dictionary keys)
  - `extract_pdf_data` in `datasheet_analyzer.py` (closes file handle when opened from `Path`)

- [ ] **Step 1: Write the failing tests in `tests/test_validation.py` and `tests/test_vision_client.py`**

```python
def test_critical_specs_by_material_type_matches_extracted_keys():
    from bom_enricher.datasheet_analyzer import is_critical_spec_missing
    assert not is_critical_spec_missing("REL", {"relay_coil_voltage": "12VDC", "relay_contact_current": "10A", "package": "TH"})
    assert not is_critical_spec_missing("SWT", {"switch_characteristic": "ON-OFF", "package": "SMD"})
    assert not is_critical_spec_missing("TRA", {"tra_type": "NPN", "value_tra": "45V / 100MA", "package": "SOT-23"})
    assert not is_critical_spec_missing("FUS", {"type_fuse": "FAST", "fuse_current": "2A", "package": "1206"})


def test_validation_quantity_expands_reference_ranges():
    from bom_enricher.models import CustomerBomRow
    from bom_enricher.validation import _validate_quantity
    row = CustomerBomRow(row_number=2, part="C1-C5", part_number="MPN1", value="100nF", qty="5")
    qty, status, msg, issue = _validate_quantity(row)
    assert status == "confirmed"
    assert issue is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.\.venv\Scripts\pytest.exe tests/test_validation.py -k "test_critical_specs_by_material_type or test_validation_quantity_expands" -v`
Expected: FAIL

- [ ] **Step 3: Implement fixes in `datasheet_analyzer.py`, `pdf_renderer.py`, `vision_client.py`, `validation.py`, and `main.py`**

1. In `datasheet_analyzer.py`:
   - Close `Path` stream in `extract_pdf_data` using a `try ... finally` block.
   - Update `CRITICAL_SPECS_BY_MATERIAL_TYPE` (`L4907-4924`): `"REL": ("relay_coil_voltage", "relay_contact_current", "package")`, `"SWT": ("switch_characteristic", "package")`, `"TRA": ("tra_type", "value_tra", "package")`, `"FUS": ("type_fuse", "fuse_current", "package")`.
   - Use word-boundary matching (`re.search(rf"\b{re.escape(s_term)}\b", mfr_lower)`) in LLM manufacturer resolution (`L5480-5489`).
2. In `pdf_renderer.py` & `datasheet_analyzer.py`:
   - Do not pass blank white images to Vision LLM if PDF rendering failed.
3. In `vision_client.py`:
   - Fix `"SWI"` -> `"SWT"` (`L254`) and do not overwrite `"reverse_voltage"` with `"forward_voltage"` (`L345`).
4. In `validation.py` & `main.py`:
   - Use `split_references(row.part)` in `_validate_quantity` (`L68`), add `"SMT"`, `"THT"`, `"DIP"` to `_canonical_mounting` (`L297`), and catch `(BadZipFile, InvalidFileException)` in `main.py` (`L40`).

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_validation.py tests/test_vision_client.py tests/test_cli.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py src/bom_enricher/pdf_renderer.py src/bom_enricher/vision_client.py src/bom_enricher/validation.py src/bom_enricher/main.py tests/test_validation.py
git commit -m "fix(datasheet,validation): align critical spec keys, close PDF streams, and expand ranges in validation"
```

---

### Task 6: Web Backend & Frontend Reliability (`web.py` & `static/index.html`)

**Files:**
- Modify: `src/bom_enricher/web.py:80-335`
- Modify: `src/bom_enricher/static/index.html:1330-1360, 1600-1820, 1975-2335`
- Test: `tests/test_web.py`

**Interfaces:**
- Modifies:
  - `/api/compare` in `web.py`: guarantees `job_dir` cleanup on any exception, supports `.csv` uploads (auto-converted to workbook) and clear Turkish error messages for encrypted/corrupted files, fixes `info_count` double-counting, and preserves subfolder prefixes on duplicate filenames.
  - `/api/datasheet-analyze` in `web.py`: enforces `_MAX_UPLOAD_BYTES` limit (HTTP 413) and uses `get_default_rulebook()`.
  - `static/index.html`: renders `"Kaynak Kart / BOM"`, `"Müşteri Referans"`, and `"ERP Referans"` columns, adds interactive status filter on metric cards, and eliminates inline `onclick` string-interpolation bugs.

- [ ] **Step 1: Write the failing tests in `tests/test_web.py`**

```python
def test_web_handles_csv_upload_and_corrupted_zip_without_500():
    client = TestClient(app)
    csv_cust = b"Part Number,Description,Reference,Qty\nRC0402FR-0710KL,10K,R1,1\n"
    erp_bytes = _workbook_bytes("ERP", ["Part Number", "Description", "Reference", "Qty", "Mounting"], [
        ["RC0402FR-0710KL", "10K SMD", "R1", 1, "SMD"],
    ])
    resp = client.post(
        "/api/compare",
        files=[
            ("customer_files", ("cust.csv", csv_cust, "text/csv")),
            ("erp_file", ("erp.xlsx", erp_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
        ],
    )
    assert resp.status_code == 200
    data = resp.json()
    assert data["clean_count"] == 1


def test_web_datasheet_analyze_rejects_oversized_pdf(monkeypatch):
    import bom_enricher.web as web_mod
    monkeypatch.setattr(web_mod, "_MAX_UPLOAD_BYTES", 100)
    client = TestClient(app)
    resp = client.post(
        "/api/datasheet-analyze",
        files=[("pdf_files", ("big.pdf", b"%PDF-1.4 " + b"0" * 200, "application/pdf"))],
    )
    assert resp.status_code == 413
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.\.venv\Scripts\pytest.exe tests/test_web.py -k "test_web_handles_csv_upload or test_web_datasheet_analyze_rejects_oversized" -v`
Expected: FAIL

- [ ] **Step 3: Implement backend and frontend hardening in `src/bom_enricher/web.py` and `src/bom_enricher/static/index.html`**

1. In `web.py`:
   - Add `_safe_cleanup(job_dir)` that ignores `PermissionError` on Windows.
   - Catch `Exception` in `compare_uploads` after specific handlers so `job_dir` is always cleaned up and malformed XML/missing ZIP entries return HTTP 422 JSON instead of HTTP 500.
   - Allow `.csv` uploads in `compare_uploads` by converting CSV rows into a temporary `.xlsx` workbook via `csv.reader` + `openpyxl.Workbook`. Detect OLE2 header (`b"\xd0\xcf\x11\xe0"`) and return a clear Turkish message explaining that the file is password-protected or in legacy `.xls` format.
   - Check `len(content) > _MAX_UPLOAD_BYTES` and `len(pdf_bytes) > _MAX_UPLOAD_BYTES` in `analyze_datasheets` (`L206, L246`), raising HTTP 413.
   - Use `get_default_rulebook()` in `analyze_datasheets` (`L220`).
   - Fix `info_count` (`L143`) to require `row.get("Durum") == "UYARI"`.
2. In `static/index.html`:
   - Include `"Kaynak Kart / BOM"`, `"Müşteri Referans"`, and `"ERP Referans"` in `renderTable`.
   - Make metric cards clickable to filter `result.rows` by status (`ALL`, `HATA`, `UYARI`, `INFO`, `UYGUN`).
   - Replace inline `onclick="copyText(...)"` and `onclick="removePdfFile(...)"` with `data-*` attributes + delegated click listeners so quotes in descriptions or filenames never cause JS syntax errors.

- [ ] **Step 4: Run tests to verify they pass**

Run: `.\.venv\Scripts\pytest.exe tests/test_web.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/web.py src/bom_enricher/static/index.html tests/test_web.py
git commit -m "fix(web): harden upload error handling, add CSV support, enforce PDF size limits, and show refs in UI"
```
