# Systemic Datasheet Intelligence & Universal Rules Architecture Design

**Date:** 2026-10-20  
**Author:** Antigravity Team & Human Partner  
**Status:** DRAFT (Under Builder Check)  
**Target Branch:** `dev-sudem`

---

## 1. Executive Summary & Intent

### Problem Statement
Previous fixes for component datasheets (such as Degson connectors, Raltron crystals, Foryard displays, and Panjit diodes) resolved individual component edge cases, but revealed three underlying architectural bottlenecks that caused repeat failures on newly encountered datasheets:
1. **The Fast-Path "Text Dropping" Trap:** When part number decoding succeeded for any property, `analyze_datasheet` set `text = ""`, completely disabling PDF text extraction. As a result, critical document-level specifications (package/case mold, voltage, current, and manufacturer) were never extracted from the PDF.
2. **Hardcoded Manufacturer Whitelist Bottleneck:** `rules.xlsx` defines **644 manufacturers** under the `Q` attribute with their official DEKA ERP codes and names. However, the codebase maintained a hardcoded Python tuple list (`KNOWN_MANUFACTURERS`) of only ~35 entries. Any manufacturer present in DEKA's rulebook but missing from the Python list was silently missed.
3. **Excel Null/Numeric Zero Asymmetry:** In `rules.xlsx`, standard commercial or default attributes (such as `Special Definition - ALL`) have an empty cell (`None`) for description and numeric `0` for value. The semantic rule matcher looked for text matches and skipped these rows, causing default ERP codes (like `00`) to be dropped or marked `BULUNAMADI`.

### Goal
Implement a truly systemic, data-driven architecture that:
- Automatically loads and indexes all 644 manufacturers directly from `rules.xlsx`.
- Guarantees PDF textual extraction is never suppressed by part number decoding, merging fast-path deductions with document-level text specs.
- Handles empty/standard default rules across all material types, ensuring complete and deterministic DEKA ERP codes.

---

## 2. System Architecture & The Three Pillars

```
                               ┌────────────────────────────────┐
                               │     DEKA rules.xlsx (Source)   │
                               │  - 67 Material Categories      │
                               │  - 644 Manufacturers (Q)       │
                               │  - Official ERP Codes          │
                               └───────────────┬────────────────┘
                                               │
                                               ▼
                              ┌──────────────────────────────────┐
                              │     comparison.load_rulebook     │
                              │  - Parse all RuleEntry records   │
                              │  - Dynamic Manufacturer Registry │
                              │    (rulebook.manufacturers)      │
                              └────────────────┬─────────────────┘
                                               │
               ┌───────────────────────────────┴───────────────────────────────┐
               ▼                                                               ▼
   ┌───────────────────────┐                                       ┌───────────────────────┐
   │ Fast-Path PN Decoder  │                                       │ PDF Text Extractor    │
   │ (Instant Spec Clues)  │                                       │ (Never suppressed!)   │
   └───────────┬───────────┘                                       └───────────┬───────────┘
               │                                                               │
               └───────────────────────────────┬───────────────────────────────┘
                                               │
                                               ▼
                               ┌────────────────────────────────┐
                               │    extract_component_specs     │
                               │  - Dynamic Mfr Matching (644)  │
                               │  - Package, Ratings, Polarity  │
                               └───────────────┬────────────────┘
                                               │
                                               ▼
                               ┌────────────────────────────────┐
                               │   Universal map_specs_to_erp   │
                               │  - Multilingual Synonyms       │
                               │  - Multi-Dimensional Guarding  │
                               │  - Null/Default "00" Fallback  │
                               └───────────────┬────────────────┘
                                               │
                                               ▼
                               ┌────────────────────────────────┐
                               │    Deterministic ERP Code      │
                               │    + Standard Rich Description │
                               └────────────────────────────────┘
```

---

## 3. Detailed Component Specifications

### Pillar 1: Dynamic Universal Manufacturer Registry
1. **RuleBook Extension (`src/bom_enricher/comparison.py`):**
   - In `RuleBook`, add `manufacturers: tuple[tuple[str, str, str], ...] = ()` containing `(search_term, display_name, erp_code)`.
   - When loading `rules.xlsx`, inspect all rows where attribute name is `Q` or `Manufacturer`.
   - For each entry with non-empty `value` and `value_description`:
     - Clean `value_description` (strip trailing commas, clean whitespace).
     - Extract normalized search tokens (e.g. `Panjit International Inc.` -> `panjit`, `panjit international`).
     - Remove common company suffixes (`Inc`, `Co., Ltd.`, `Corp`, `Electronics`, `Semiconductor`, `ALT 1`) to generate high-precision regex match patterns.
     - **Disambiguation & Collision Blacklist (Builder Call):** Generic words that collide with component categories (specifically manufacturer code `DI` described as `"Diodes"`, or words like `"Resistors"`, `"Capacitors"`, `"Connectors"`) are excluded from naked search; they must be qualified with corporate terms (e.g. `"diodes inc"`, `"diodes incorporated"`).
     - **Authoritative Precedence:** `KNOWN_MANUFACTURERS` acts as an authoritative curated alias layer, checked before or merging with dynamically generated patterns.
     - Sort manufacturer patterns by length descending so specific multi-word names match before general substrings.
2. **Datasheet Analyzer Integration (`src/bom_enricher/datasheet_analyzer.py`):**
   - `extract_component_specs` receives `rulebook: RuleBook | None = None`.
   - When extracting manufacturer, it matches against the curated alias layer and dynamic registry.
   - Any manufacturer appearing in `rules.xlsx` is recognized automatically across all 67 material categories without writing Python code.

### Pillar 2: Non-Destructive Fast-Path & Complete Text Awareness
1. **Remove Destructive `text = ""`:**
   - In `analyze_datasheet`, eliminate the assignment `text = ""` under `if pn_has_core:`.
   - PDF text extraction via `extract_pdf_data` is ALWAYS performed if `file_bytes` is present.
   - The fast-path optimization is scoped strictly where it belongs: bypassing heavy 2D vector CAD table parsing (`file_bytes=None` passed to table extraction), keeping execution time <150ms while retaining full text searchability.
2. **Merge Precedence (Builder Call):**
   - Ground-truth precedence: Part number decoding is authoritative for parameters it explicitly decodes (e.g. `0603` package and `10K` resistance on chip resistor PNs; `20V` and `500mA` on `SS0520` diode).
   - Document text fills all remaining missing fields (e.g. `SOD-123` package and `Panjit` manufacturer for `SS0520`).
   - Merge rule: `pn_specs` values overwrite text extracted values only for fields present in `pn_specs`, ensuring family datasheets listing multiple packages do not contaminate decoded passive packages.

### Pillar 3: Null/Default Standard Rule Fallback
1. **Scope & Guardrails (Builder Call):**
   - Strictly restricted to attributes matching `specialdefinition...` (e.g. `Special Definition - ALL`, `Special Definition - Transistors`, `Special Definition - Capacitors`, `Special Definition - Resistors`).
   - Core functional attributes (`Package`, `Mounting`, `Wattage`, `Pin Number`, `Pitch`, `Value`) NEVER silently default to zero. If missing, they remain `BULUNAMADI`.
2. **Zero Formatting & Description Suppression (Builder Call):**
   - When `specs` has no explicit modifier (e.g. no `Automotive AEC-Q...` text found), and a rule entry exists with `value in (0, "0")` or `value_description in (None, "", "Standart", "Standard")`:
     - Dynamically inspect peer values for this attribute: if max sibling code length is 2 (e.g. `AU`, `AC`, `AE` in `DIO`), format numeric `0` as `"00"`. If max sibling code length is 1 (e.g. `A` in `TRA`), format as `"0"`.
     - Standard default qualifiers (`"Standart"`, `"None"`, `"General"`, `""`) are suppressed from description tokens (`suggested_desc` / `rich_description`), preserving clean engineering descriptions without clutter.
     - Mark status as `TAM EŞLEŞTİ` with matched value `"Standart"`.

---

## 4. Constraints & Non-Goals

### Constraints
- 100% offline and local execution; zero network or external API dependencies.
- ERP codes must remain 100% deterministic (strictly derived from `rules.xlsx`).
- Zero regressions across the entire test suite (151+ pytest tests).
- Fast execution (<150ms per datasheet in text mode).

### Non-Goals
- We are not modifying the physical `rules.xlsx` file. All normalization of Excel data types (e.g. numeric 0 vs string "00") is handled cleanly in memory.
- We are not replacing regex or Vision LLM with cloud LLMs.

---

## 5. Verification Strategy

1. **Universal Manufacturer Tests:**
   - Verify that manufacturers present in `rules.xlsx` but never hardcoded in Python (e.g., `Qualitek` -> `QA`, `Kepo` -> `KE`, `Piher` -> `PI`) are recognized directly from PDF text.
   - Verify that manufacturer code `DI` does NOT false-match arbitrary diode texts.
2. **Complete Text Extraction Tests:**
   - Verify that parts with decoded part numbers (e.g. `SS0520`, `RC0603FR-0710KL`) still extract case/package and manufacturer from PDF text without regressing package accuracy.
3. **Null/Default Fallback Tests:**
   - Verify that standard commercial diodes map `Special Definition` to `00`, while transistors map to `0`.
4. **Full Regression Suite:**
   - All 151 existing tests must pass with zero errors.

