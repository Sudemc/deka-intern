# 2D PDF Table Extraction with pdfplumber Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Integrate offline, open-source 2D table extraction using `pdfplumber` into `datasheet_analyzer.py` to accurately parse multi-variant catalog tables, pin configurations, and electrical specifications without any external API or cloud calls.

**Architecture:** Add a native 2D table parser layer on top of `pdfplumber`. For any PDF, extract tabular grids per page; locate rows containing target Part Numbers (or suffixes); resolve columns to technical attributes (dimensions, connectors, pins, voltages, values) with coordinate precision, falling back seamlessly to `pypdf` text extraction when tables are absent or `pdfplumber` is unavailable.

**Tech Stack:** Python 3.10+, `pdfplumber`, `pypdf`, `openpyxl`, `FastAPI`, `pytest`.

**Spec:** Local/On-Premise Company Portal Requirement: 100% offline, zero external API keys, zero cloud latency/costs, zero data leakage, deterministic table matrix parsing.

## Global Constraints
- Must remain 100% offline (no external network/API calls).
- Must have graceful fallback to `pypdf` if `pdfplumber` encounters a corrupted page or image-only PDF.
- Preserve all existing 88 passing tests and existing ERP mapping rules.

## Review Focus
1. Multi-page catalog where target part number appears on page 5 in table 2 (e.g. TE antenna `2108857-1`).
2. Catalog table where table header has "Length" and "Connector" columns and target row has corresponding cells.
3. PDF with no tables (plain text or non-bordered text), verifying fallback does not crash.
4. Part numbers with dashes and suffixes matching table rows cleanly.
5. In-memory `bytes` support so web uploads without disk paths parse directly via `io.BytesIO`.

---

### Task 1: Update Project Dependencies with `pdfplumber`
**Files:**
- Modify: `pyproject.toml:11-20`
- Test: Verify import in Python

- [x] Add `pdfplumber>=0.11` to `dependencies` in `pyproject.toml`.
- [x] Test importing `pdfplumber` in the virtual environment.

---

### Task 2: Implement Native 2D Table Extractor in `datasheet_analyzer.py`
**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Test: `tests/test_datasheet_analyzer.py`

- [x] Add `extract_pdf_tables(file_bytes: bytes) -> list[dict[str, Any]]` helper using `io.BytesIO` and `pdfplumber`.
- [x] Implement `find_part_row_in_tables(tables: list[dict], part_number: str) -> dict[str, Any]` to match rows with target MPNs and extract exact column cells.
- [x] Integrate 2D table matching into `extract_component_specs` before falling back to line-by-line regex.

---

### Task 3: Comprehensive Unit Tests & Real Datasheet Verification
**Files:**
- Modify: `tests/test_datasheet_analyzer.py`

- [x] Write unit test `test_pdfplumber_table_extraction_antenna` testing TE Connectivity Page 5 table extraction.
- [x] Write unit test `test_pdfplumber_table_extraction_connector` testing tabular connector pin and pitch resolution.
- [x] Verify fallback behavior when a PDF has no detectable tables.

---

### Task 4: Full Test Suite & Server Verification
**Files:**
- Test: All 88+ pytest tests
- Web: `http://127.0.0.1:8000`

- [x] Run `pytest` across all test suites to confirm 100% pass rate.
- [x] Restart web server daemon and verify API endpoint `/api/datasheet-analyze`.
- [x] Commit and push changes to `origin` and `personal` branches.
