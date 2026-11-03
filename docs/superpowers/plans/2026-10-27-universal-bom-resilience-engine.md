# Universal BOM Resilience Engine Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transform BOM table parsing, MPN extraction, reference designator validation, and semantic description comparison into a robust, non-destructive, multi-value engine that eliminates fragile regex patching and prevents false errors on unseen customer/ERP BOMs.

**Architecture:** 
1. Replace single-value regex matching with multi-value set extraction (`MultiSpec`) and an MPN-verified semantic shield.
2. Replace whitelist-based character filtering with non-destructive MPN extraction and cross-BOM catalog candidate resolution.
3. Replace regex-fallback designator splitting with an IEEE/IPC-compliant designator grammar and embedded note/MPN extractor.
4. Unify ERP and Customer table loading into a shared data-profiling reader that detects columns by content density rather than fragile header string matching.
5. Add per-row pre-aggregation validation and golden corpus regression verification covering all historical production BOMs.

**Tech Stack:** Python 3.10+, Pandas, OpenPyXL, FastAPI, Pytest.

**Spec:** User approval on 2026-10-27 for the 6-phase permanent architectural plan.

## Global Constraints
- Python 3.10+ compatibility with zero new external runtime dependencies.
- Preserve 100% backward compatibility on existing test suite (`tests/test_comparison.py`, `tests/test_excel_io.py`).
- Existing passing baselines must remain intact:
  - Upper BOM: 34 UYGUN, 5 UYARI, 4 HATA
  - Lower BOM: 71 UYGUN, 7 UYARI, 1 HATA
  - Delta-V Orta: 20 UYGUN, 1 UYARI, 2 HATA
  - Delta-V Telemetri V2: 42 UYGUN, 1 UYARI, 6 HATA
- Clean cell rule: all `UYGUN` rows MUST have an empty string `""` in the `Hata/Uyarı Durumu` column.

## Review Focus
1. Descriptions containing multiple values of the same unit (e.g. Ferrite bead 220R impedance + 45mOhm DCR, or Tantalum 10uF cap + 800mOhm ESR) must never trigger false value mismatch errors.
2. MPNs containing unusual punctuation (`#`, `(`, `)`, `,`, `/`, `:`, `~`, `=`, `[`, `]`, `%`, `@`, `;`, `+`, `-`, `_`, `.`, `*`) must never be wiped to empty strings.
3. Reference designator columns containing human notes or alternative MPNs (e.g. `U11 alternatif eklenecek (ADR4525ARZ)`) must not treat MPNs or notes as circuit designators.
4. ERP BOMs with diverse column naming (e.g. `Parça` for line numbers, `ÜRETİCİ KODU` alongside `DTSIS KODU`) must correctly identify reference columns and capture all part numbers.
5. Malformed designators lacking letter prefixes (e.g. `'1'` instead of `'U1'`) must produce an explicit warning rather than passing silently as valid designators.

---

### Task 1: Evrensel PCB Referans (Designator) Grameri ve Not/Gürültü Ayıklayıcı

**Files:**
- Modify: `src/bom_enricher/excel_io.py:180-191, 287-315`
- Test: `tests/test_excel_io.py`

**Interfaces:**
- Consumes: `raw_ref_string: str`
- Produces: `split_references(s: str) -> list[str]`, `extract_embedded_part_numbers(s: str) -> list[str]`, `validate_reference_tokens(refs: list[str]) -> list[str]` (malformed token warnings)

- [x] **Step 1: Write failing tests for designator grammar and embedded note stripping**
  In `tests/test_excel_io.py`, test:
  - `split_references("U11 alternatif eklenecek (ADR4525ARZ)") == ["U11"]`
  - `split_references("R1, R2 - dik monte edilecek (0805)") == ["R1", "R2"]`
  - `split_references("L2, L3, L4, L5, L6, L8, L9") == ["L2", "L3", "L4", "L5", "L6", "L8", "L9"]`
  - `extract_embedded_part_numbers("U11 alternatif eklenecek (ADR4525ARZ)") == ["ADR4525ARZ"]`
  - Malformed designator detection: `'1'` tagged as malformed (missing component class letter).

- [x] **Step 2: Run test to verify failure**
  Run: `.\.venv\Scripts\pytest.exe tests/test_excel_io.py -k "test_split_references" -v`

- [x] **Step 3: Implement IPC/IEEE-compliant designator grammar and embedded code extraction in `excel_io.py`**
  - Delete line 311 fallback (`elif any(c.isdigit() for c in r_clean) and len(r_clean) <= 20: refs.append(r_clean)`).
  - Define `DESIGNATOR_REGEX = re.compile(r"^(?:[A-Za-z0-9]{2,10}[_\-])?[A-Za-z]{1,5}\d{1,4}[A-Za-z]?$")`.
  - Add filtering against component values (`100NF`, `10K`, `25V`, `0805`), natural language words, and extract embedded MPNs.
  - Implement `extract_embedded_part_numbers(s: str) -> list[str]`.

- [x] **Step 4: Run test to verify it passes**
  Run: `.\.venv\Scripts\pytest.exe tests/test_excel_io.py -v`

- [x] **Step 5: Commit**
  ```bash
  git add src/bom_enricher/excel_io.py tests/test_excel_io.py
  git commit -m "feat(parser): add universal designator grammar and embedded note extraction"
  ```

---

### Task 2: Karartmasız (Non-Destructive) MPN Çözümleyici & Çapraz Katalog MPN Yakalama

**Files:**
- Modify: `src/bom_enricher/excel_io.py:212-266`
- Modify: `src/bom_enricher/comparison.py:151-200`
- Test: `tests/test_excel_io.py`, `tests/test_comparison.py`

**Interfaces:**
- Consumes: Raw MPN string from cell, optional context of catalog codes
- Produces: `clean_part_number(s: str) -> str`, `is_valid_part_number(s: str) -> bool`, `split_multi_code(raw: str) -> list[str]`

- [x] **Step 1: Write failing tests for non-destructive MPN validation**
  In `tests/test_excel_io.py`, test:
  - MPNs with `#`: `LT3092ITS8#TRPBF`, `LT6015IS5#TRMPBF` preserved.
  - MPNs with punctuation: `TLP293(BL-TPL,E`, `ACS780LLRTR-100U-T`, `MUSBK15230`, `ABM8AAIG-25.000MHZ-V2R-T3`, `PESD5V0S5UD,115` preserved without truncation.
  - Reject only explicit placeholders (`YOK`, `DNP`, `NONE`, `TBD`, `-`), pure dates, sequential integers, and multi-word prose.

- [x] **Step 2: Run test to verify failure**
  Run: `.\.venv\Scripts\pytest.exe tests/test_excel_io.py -k "test_part_number" -v`

- [x] **Step 3: Implement non-destructive MPN cleaning and validation**
  - In `clean_part_number()`, preserve all electronic packaging punctuation.
  - In `_explicit_code_like()`, replace narrow whitelist with blacklist of prose and placeholders.
  - In `_split_multi_code()`, extract all potential MPN tokens including parenthetical codes and package suffixes.

- [x] **Step 4: Run test to verify it passes**
  Run: `.\.venv\Scripts\pytest.exe tests/test_excel_io.py tests/test_comparison.py -v`

- [x] **Step 5: Commit**
  ```bash
  git add src/bom_enricher/excel_io.py src/bom_enricher/comparison.py tests/test_excel_io.py
  git commit -m "feat(mpn): non-destructive MPN validation preserving all component codes"
  ```

---

### Task 3: Çoklu Değer Kümesi (Multi-Set) ve MPN-Korumalı Semantik Kıyaslama Motoru

**Files:**
- Modify: `src/bom_enricher/semantic_parser.py`
- Modify: `src/bom_enricher/comparison.py:1097-1099`
- Test: `tests/test_semantic_parser.py`

**Interfaces:**
- Consumes: `customer_text: str`, `erp_text: str`, `mpn_matched: bool = False`
- Produces: `compare_descriptions(customer_text: str, erp_text: str, mpn_matched: bool = False) -> list[str]`, `parse_specs_set(text: str) -> MultiElectricalSpecs`

- [x] **Step 1: Write failing tests for multi-spec extraction and MPN-verified comparison**
  Create `tests/test_semantic_parser.py` (or update existing) testing:
  - Ferrite bead: `"220 Ohms @ 100 MHz 1 Power Line Ferrite Bead 0805 (2012 Metric) 2A 45mOhm"` vs `"SMD 0805 Package, F.Bead: 220R, 2A, 45mOhm, Murata"` returns `[]` (no contradiction).
  - Tantalum cap: `"10 µF Molded Tantalum Capacitors 16 V 1210 (3225 Metric) 800mOhm @ 100kHz"` matches `"10uF 16V 1210"`.
  - When `mpn_matched=True`, description wording variations never flag false critical errors.
  - Genuine mismatch: `100R` vs `10k` resistor still flags contradiction.

- [x] **Step 2: Run test to verify failure**
  Run: `.\.venv\Scripts\pytest.exe tests/test_semantic_parser.py -v`

- [x] **Step 3: Implement `MultiElectricalSpecs` and set-based `compare_descriptions` in `semantic_parser.py`**
  - Implement regex finditer for all resistances, capacitances, voltages, tolerances, and packages returning `frozenset[Decimal]` and `frozenset[str]`.
  - Evaluate compatibility via set intersection (`cust_set & erp_set != empty`).
  - Add `mpn_matched` guard: when MPNs are verified identical, skip secondary-parameter contradiction checks.

- [x] **Step 4: Run test to verify it passes**
  Run: `.\.venv\Scripts\pytest.exe tests/test_semantic_parser.py -v`

- [x] **Step 5: Commit**
  ```bash
  git add src/bom_enricher/semantic_parser.py src/bom_enricher/comparison.py tests/test_semantic_parser.py
  git commit -m "feat(semantic): multi-spec set extraction and MPN-verified comparison shield"
  ```

---

### Task 4: Birleştirilmiş Veri-Güdümlü Tablo Okuyucu (Unified Table Reader)

**Files:**
- Modify: `src/bom_enricher/excel_io.py`
- Modify: `src/bom_enricher/comparison.py:1299-1335, 1721-1736`
- Test: `tests/test_comparison.py`, `tests/test_excel_io.py`

**Interfaces:**
- Consumes: Customer or ERP workbook path
- Produces: `read_bom_workbook(path: Path, kind: str) -> list[PartRecord]` with data profiling for column detection

- [x] **Step 1: Write failing tests for ERP data-profiling column detection**
  In `tests/test_comparison.py`, test:
  - ERP file with `Parça` (row numbers 1..N) and `REFERANS` (`D13, D15`) correctly maps `REFERANS` to reference designators, not `Parça`.
  - ERP file with both `DTSİS KODU` and `ÜRETİCİ KODU` captures both in `all_codes`.
  - Customer row with bare number `'1'` as designator generates a malformed designator warning.

- [x] **Step 2: Run test to verify failure**
  Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -k "test_erp_table_profiling" -v`

- [x] **Step 3: Replace `_read_table_with_detected_header` with unified data-profiling table reader**
  - Use `excel_io.py` table detection logic for both Customer and ERP.
  - Disambiguate columns by checking contents (% reference tokens, % MPN codes, sequential integer sequences).
  - Aggregate all code columns into `all_codes`.
  - Flag malformed reference designators (`STM32F103C8T6` with `'1'`).

- [x] **Step 4: Run test to verify it passes**
  Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py tests/test_excel_io.py -v`

- [x] **Step 5: Commit**
  ```bash
  git add src/bom_enricher/excel_io.py src/bom_enricher/comparison.py tests/test_comparison.py
  git commit -m "feat(reader): unified data-profiling table reader for both customer and ERP workbooks"
  ```

---

### Task 5: Satır İzlenebilirliği ve Ön-Doğrulama

**Files:**
- Modify: `src/bom_enricher/comparison.py:1010-1060, 1611-1718`
- Test: `tests/test_comparison.py`

**Interfaces:**
- Consumes: `records: list[PartRecord]`
- Produces: `_unique_records(records: list[PartRecord]) -> dict[str, PartRecord]` with per-row validation and clear source annotations

- [x] **Step 1: Write failing test for multi-row MPN aggregation annotations**
  In `tests/test_comparison.py`, test:
  - If a BOM has 2 rows for `BLM21PG221SN1D` (Row 2: 12 pcs + Row 18: 2 pcs), `Kaynak Kart / BOM` records `Satır 2 (12 adet) + Satır 18 (2 adet)`.
  - If an individual row has Qty mismatch before merge, the error pinpoints the exact line number.

- [x] **Step 2: Run test to verify failure**
  Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -k "test_multi_row_annotation" -v`

- [x] **Step 3: Implement per-row pre-validation and aggregation tracking**
  - Check `Qty == len(refs)` per source row before grouping.
  - Track source line numbers and individual quantities.

- [x] **Step 4: Run test to verify it passes**
  Run: `.\.venv\Scripts\pytest.exe tests/test_comparison.py -v`

- [x] **Step 5: Commit**
  ```bash
  git add src/bom_enricher/comparison.py tests/test_comparison.py
  git commit -m "feat(comparison): row-level pre-validation and multi-row trace annotations"
  ```

---

### Task 6: Altın Standart (Golden Corpus) Regresyon Testleri ve Canlı Sunucu Doğrulaması

**Files:**
- Create: `tests/test_golden_corpus.py`
- Modify: `src/bom_enricher/web.py`
- Test: Full test suite

**Interfaces:**
- Consumes: All 5 historical production datasets (`Upper`, `Lower`, `DCP4`, `Delta-V Orta`, `Delta-V Telemetri`)
- Produces: Automated verification that every dataset matches its exact expected baseline with zero regressions.

- [x] **Step 1: Implement `tests/test_golden_corpus.py` covering all 5 datasets**
  - Verify Upper BOM: 34 UYGUN, 5 UYARI, 4 HATA.
  - Verify Lower BOM: 71 UYGUN, 7 UYARI, 1 HATA.
  - Verify Delta-V Orta: 20 UYGUN, 1 UYARI, 2 HATA.
  - Verify Delta-V Telemetri: 41 UYGUN, 2 UYARI, 6 HATA.
  - Verify all UYGUN rows have clean, empty issue cells.

- [x] **Step 2: Run the full test suite**
  Run: `.\.venv\Scripts\pytest.exe tests/test_golden_corpus.py tests/test_comparison.py tests/test_excel_io.py -v`
  Expected: All tests PASS.

- [x] **Step 3: Verify and restart web application**
  Verify `web.py` endpoints and ensure the service is cleanly running.

- [x] **Step 4: Final Git commit & push**
  ```bash
  git add tests/test_golden_corpus.py src/bom_enricher/web.py
  git commit -m "test(golden): add comprehensive golden corpus regression suite for all production BOMs"
  git push personal dev-sudem
  ```
