# Multi-BOM Comparison Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enhance multi-BOM comparison by adding source file/sheet details with per-board quantity breakdown, customer and ERP reference columns to reports, expanding ERP prefix delimiter support (`_`, `.`, `-`), and issuing warnings for unmapped customer files.

**Architecture:** Extend `comparison.py` to preserve source breakdown per record during deduplication, format designator reference columns for both customer and ERP, generalize prefix extraction to handle multiple punctuation delimiters, and update report headers and Excel styling.

**Tech Stack:** Python 3.10+, pandas, openpyxl, pytest.

**Spec:** User conversation and requirements discussed on 2026-10-20.

## Global Constraints
- Pure Python and offline-only: no external API or cloud calls.
- Maintain backward compatibility of `compare_workbooks` signature.
- Keep `Müşteri Qty` strictly numeric/parsable so quantity mismatch comparisons remain mathematically exact.

## Review Focus
1. Numeric Qty comparison: `customer_qty` must not break `float()` parsing (breakdown belongs in `Kaynak Kart / BOM`).
2. Single-BOM backward compatibility: Single BOM inputs must still output clean rows without cluttered breakdown text.
3. Multiple files using same part number: Deduplication must correctly aggregate references and list individual file quantities.
4. Prefix delimiters: ERP references formatted as `TR1.C1` or `TR1-C1` must be recognized identically to `TR1_C1`.
5. Missing prefix warning: Files that cannot be matched to any ERP prefix must generate a clear advisory warning.

---

### Task 1: Report Headers and Column Extension

**Files:**
- Modify: `src/bom_enricher/comparison.py:20-31`, `src/bom_enricher/comparison.py:630-650`, `src/bom_enricher/comparison.py:830-875`
- Test: `tests/test_comparison.py:80-92`

**Interfaces:**
- Consumes: `OUTPUT_HEADERS`
- Produces: Updated 13-column `OUTPUT_HEADERS` and Excel formatting

- [ ] **Step 1: Write failing test for 13-column output headers**

Add test in `tests/test_comparison.py` verifying that `OUTPUT_HEADERS` includes `"Kaynak Kart / BOM"`, `"Müşteri Referans"`, and `"ERP Referans"`, and that the written Excel report has 13 header columns.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_output_headers`
Expected: FAIL

- [ ] **Step 3: Update `OUTPUT_HEADERS` and `write_comparison_report` in `src/bom_enricher/comparison.py`**

Define `OUTPUT_HEADERS`:
```python
OUTPUT_HEADERS = (
    "Durum",
    "Müşteri Part Number",
    "ERP Part Number",
    "Müşteri Qty",
    "ERP Qty",
    "Müşteri Value",
    "ERP Description",
    "Müşteri TH/SMD",
    "ERP TH/SMD",
    "Kaynak Kart / BOM",
    "Müşteri Referans",
    "ERP Referans",
    "Hata/Uyarı Durumu",
)
```
Adjust `widths` tuple in `write_comparison_report` to match 13 columns.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_output_headers`
Expected: PASS

---

### Task 2: Source Breakdown and Reference Display in Records

**Files:**
- Modify: `src/bom_enricher/comparison.py:55-70` (`PartRecord`), `src/bom_enricher/comparison.py:874-940` (`_unique_records`), `src/bom_enricher/comparison.py:625-648` (report mapping)
- Test: `tests/test_comparison.py`

**Interfaces:**
- Consumes: `PartRecord.source_file`, `PartRecord.source_sheet`, `PartRecord.qty`
- Produces: Formatted `source_breakdown` field on `PartRecord` and mapped report columns

- [ ] **Step 1: Write failing test for quantity breakdown and reference columns**

Create test `test_multi_bom_source_breakdown_and_references` in `tests/test_comparison.py` with 2 customer BOMs sharing a common resistor. Verify `Kaynak Kart / BOM` displays the file/sheet breakdown with quantities, and `Müşteri Referans` lists all designators.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_multi_bom_source_breakdown`
Expected: FAIL

- [ ] **Step 3: Implement source breakdown tracking in `_unique_records` and report mapping**

Add `source_breakdown: str = ""` to `PartRecord`.
In `_unique_records`, build per-source quantities:
```python
source_counts = defaultdict(int)
for item in group:
    src_label = f"{item.source_file} ({item.source_sheet})" if item.source_sheet else item.source_file
    try:
        source_counts[src_label] += int(float(item.qty))
    except (ValueError, TypeError):
        source_counts[src_label] += 1
```
Format:
If multiple sources, `source_breakdown = ", ".join(f"{src}: {cnt} adet" for src, cnt in source_counts.items())`.
If single source, `source_breakdown = list(source_counts.keys())[0]`.
Map `"Kaynak Kart / BOM"`, `"Müşteri Referans"`, and `"ERP Referans"` in `compare_workbooks`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_multi_bom_source_breakdown`
Expected: PASS

---

### Task 3: Delimiter Flexibility and Unmapped Prefix Warning

**Files:**
- Modify: `src/bom_enricher/comparison.py:198-235` (`_extract_erp_prefixes`), `src/bom_enricher/comparison.py:297-364` (`_qualify_customer_references`)
- Test: `tests/test_comparison.py`

**Interfaces:**
- Consumes: ERP references with delimiters (`_`, `.`, `-`)
- Produces: Normalized prefixes and warning issues when files fail prefix resolution

- [ ] **Step 1: Write failing test for dot/dash delimiters and unmapped prefix warning**

Add test `test_prefix_delimiter_flexibility_and_unmapped_warning` in `tests/test_comparison.py` testing ERP references with dot `TR1.C1` and an unmapped customer file `Unknown_Board.xlsx`.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_prefix_delimiter`
Expected: FAIL

- [ ] **Step 3: Implement regex delimiter split and unmapped file warning**

In `_extract_erp_prefixes`, split with `re.split(r"[_\.\-]", ref, maxsplit=1)`.
In `_qualify_customer_references`, check if any `(source_file, source_sheet)` could not be resolved when `erp_prefixes` is non-empty. For unresolved records, append `"Önek Eşleşemedi: ERP kart öneki tespit edilemedi"` to description or issue list.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\pytest tests/test_comparison.py -k test_prefix_delimiter`
Expected: PASS

---

### Task 4: Regression and Full Test Suite Verification

**Files:**
- Test: `tests/test_comparison.py`

- [ ] **Step 1: Run full comparison test suite**

Run: `.venv\Scripts\pytest tests/test_comparison.py`
Expected: ALL PASS

- [ ] **Step 2: Update existing tests if header length expectations changed**

Update `sheet[1]` header assertion in `test_comparison_reports_missing_conflicting_blank_and_mounting_mismatch` to check the new 13 columns.
Run: `.venv\Scripts\pytest tests/test_comparison.py`
Expected: 100% PASS
