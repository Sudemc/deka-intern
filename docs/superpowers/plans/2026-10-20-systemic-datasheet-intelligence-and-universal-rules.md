# Systemic Datasheet Intelligence & Universal Rules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement dynamic universal manufacturer registration from `rules.xlsx` (644 manufacturers), remove fast-path text dropping so datasheets are always read, and build robust default standard rule fallback for qualifiers across all component types.

**Architecture:** 
1. Dynamic Universal Manufacturer Registry: Parse all 644 manufacturers from `rules.xlsx` (`Q` attribute) into `RuleBook.manufacturers` with corporate suffix stripping, disambiguation blacklist (e.g. `DI` -> `Diodes Inc`), and integration with curated aliases.
2. Non-Destructive Fast-Path: Remove `text = ""` assignment in `analyze_datasheet`; ensure PDF text is always extracted while scoping fast-path speed bypass strictly to 2D vector CAD table parsing.
3. Generalized Standard Rule Fallback: Enable automatic `00`/`0` mapping for `specialdefinition...` attributes when standard commercial grade is detected, keeping core attributes strictly guarded.

**Tech Stack:** Python 3.10, openpyxl, pypdf, pytest, regex.

**Spec:** `docs/superpowers/specs/2026-10-20-systemic-datasheet-intelligence-and-universal-rules-design.md`

## Global Constraints

- 100% offline and local execution; zero external cloud API dependencies.
- ERP codes must remain 100% deterministic (generated strictly by rules.xlsx and map_specs_to_erp, never hallucinated by LLM).
- Preserves all 151 existing pytest tests without regression.
- Fast execution (<150ms per datasheet in text mode).

## Review Focus

1. Category Name Collision: Manufacturer code `DI` ("Diodes") must NEVER match generic diode text like "Schottky Barrier Diodes".
2. Passive Package Immunity: Decoded passive packages (e.g. `0603` chip resistors) must not be corrupted by document text listing family package tables (`0402, 0603, 0805`).
3. Core Attribute Guard: Core attributes (`Package`, `Mounting`, `Wattage`, `Pin Number`) must NEVER default to `0` or "Standart" when missing; they must remain `BULUNAMADI`.
4. Peer Length Zero Padding: Attributes with 2-character sibling codes (like `DIO` Special Definition) format default `0` as `"00"`, while 1-character siblings (like `TRA`) format as `"0"`.
5. Description Cleanliness: Default `"Standart"` qualifiers must not appear in generated description tokens unless a non-standard modifier (e.g. `Automotive AEC-Q101`) is present.

---

### Task 1: Dynamic Universal Manufacturer Registry in `RuleBook`

**Files:**
- Modify: `src/bom_enricher/comparison.py:80-120`, `src/bom_enricher/comparison.py:860-910`
- Modify: `src/bom_enricher/datasheet_analyzer.py:210-270`, `src/bom_enricher/datasheet_analyzer.py:1320-1340`
- Test: `tests/test_comparison.py`, `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `rules.xlsx` rows with attribute `Q` or `Manufacturer`.
- Produces: `RuleBook.manufacturers: tuple[tuple[str, str, str], ...]` mapping `(search_term, display_name, erp_code)`; `extract_component_specs` and `analyze_datasheet` matching against dynamic 644-manufacturer registry.

- [x] **Step 1: Write failing unit tests for dynamic manufacturer loading and collision prevention**

```python
# In tests/test_datasheet_analyzer.py
def test_dynamic_manufacturer_loading_from_rules():
    from bom_enricher.comparison import load_rulebook
    from bom_enricher.datasheet_analyzer import extract_component_specs
    rulebook = load_rulebook("data/rules.xlsx")
    assert len(rulebook.manufacturers) > 100
    
    # 1. Unlisted manufacturer present in rules.xlsx (e.g. Qualitek -> QA, Kepo -> KE, Piher -> PI)
    text_qualitek = "Qualitek Delta Solder Wire Specifications"
    specs_q = extract_component_specs("RES", text_qualitek, "WIRE-01", rulebook=rulebook)
    assert specs_q.get("manufacturer") == "Qualitek"
    assert specs_q.get("manufacturer_code") == "QA"

    # 2. Collision test: code DI ('Diodes') must NOT match generic diode description
    text_generic_diode = "SURFACE MOUNT SCHOTTKY BARRIER DIODES VOLTAGE 40V"
    specs_dio = extract_component_specs("DIO", text_generic_diode, "GENERIC-01", rulebook=rulebook)
    assert specs_dio.get("manufacturer") != "Diodes"
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_dynamic_manufacturer_loading_from_rules -v`
Expected: FAIL (`RuleBook has no attribute 'manufacturers'` or `specs_q.get('manufacturer') == None`)

- [x] **Step 3: Implement dynamic manufacturer extraction in `comparison.py` and `datasheet_analyzer.py`**

1. In `RuleBook` dataclass in `comparison.py`: add field `manufacturers: tuple[tuple[str, str, str], ...] = ()`.
2. In `load_rulebook` in `comparison.py`:
   - Identify entries where `attribute_name in ("Q", "Manufacturer")` and `value` and `value_description`.
   - Strip company suffixes (`Inc`, `Co., Ltd.`, `Corp`, `Electronics`, `Semiconductor`, `ALT 1`, etc.).
   - Disambiguation: if search term matches common component category words (`"diodes"`, `"resistors"`, `"capacitors"`, `"connectors"`), require corporate suffix (`"diodes inc"`, `"diodes incorporated"`).
   - Merge with curated `KNOWN_MANUFACTURERS` alias layer.
   - Sort by search term length descending.
   - Store as `manufacturers` on `RuleBook`.
3. In `datasheet_analyzer.py`:
   - Pass `rulebook` into `extract_component_specs(..., rulebook: RuleBook | None = None)`.
   - In manufacturer detection block (around line 1328), search `rulebook.manufacturers` if available, falling back to `KNOWN_MANUFACTURERS`.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_dynamic_manufacturer_loading_from_rules -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/comparison.py src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: load dynamic 644-manufacturer registry from rules.xlsx with collision guards"
```

---

### Task 2: Non-Destructive Fast-Path & Complete Text Awareness

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:3340-3450`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `file_bytes`, `pn`, `pn_specs` in `analyze_datasheet`.
- Produces: `extract_pdf_data` always called when `file_bytes` is present; table extraction bypassed on fast-path; `pn_specs` merged cleanly without overwriting document-level text fields.

- [ ] **Step 1: Write failing test verifying text extraction is never suppressed on fast-path parts**

```python
# In tests/test_datasheet_analyzer.py
def test_fast_path_never_suppresses_pdf_text():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # PDF containing package and manufacturer in text, with PN having electrical specs
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "PANJIT SEMICONDUCTOR")
    c.drawString(100, 730, "SS0520 SURFACE MOUNT SCHOTTKY BARRIER")
    c.drawString(100, 710, "Case: SOD-123, Plastic")
    c.save()
    pdf_bytes = buf.getvalue()

    result = analyze_datasheet(pdf_bytes, part_number="SS0520")
    # Even though SS0520 decoded electrical specs on fast-path, text MUST be read
    assert result["specs"].get("package") == "SOD123"
    assert result["specs"].get("manufacturer") == "Panjit"
```

- [ ] **Step 2: Run test to verify it passes or fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_fast_path_never_suppresses_pdf_text -v`

- [ ] **Step 3: Refactor `analyze_datasheet` to cleanly separate text extraction from table extraction**

1. Remove `text = ""` assignment in `analyze_datasheet`.
2. Always execute `text, page_count, analyzed_pages = extract_pdf_data(file_bytes, part_number=pn)` whenever `file_bytes` is provided.
3. Keep fast-path speed bypass strictly where it belongs: `effective_file_bytes = None if pn_has_core else file_bytes` passed to `extract_component_specs` (skipping heavy 2D vector CAD table analysis).
4. Merge precedence: `for k, v in pn_specs.items(): specs[k] = v` ensuring decoded electrical parameters remain authoritative while text supplies missing package and manufacturer.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_fast_path_never_suppresses_pdf_text -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "refactor: eliminate text-dropping in fast-path and enforce non-destructive text extraction"
```

---

### Task 3: Generalized Null/Default Standard Rule Fallback for Qualifiers

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:2540-2570`, `src/bom_enricher/datasheet_analyzer.py:2930-2970`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: rulebook entries for `specialdefinition...` attributes in `map_specs_to_erp`.
- Produces: deterministic ERP codes with standard `00`/`0` padding; clean rich descriptions without `"Standart"` token pollution.

- [ ] **Step 1: Write failing tests for generic default qualifier fallback and zero padding**

```python
# In tests/test_datasheet_analyzer.py
def test_universal_default_special_definition_fallback():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp, build_rich_description
    rulebook = load_rulebook("data/rules.xlsx")

    # 1. Diode (Special Definition - ALL has 2-char siblings AU, AC, AE -> formats as '00')
    specs_dio = {
        "mounting": "SMD",
        "package": "SOD123",
        "type_diode": "Schottky",
        "value_diode": "SS0520",
        "manufacturer": "Panjit",
        "manufacturer_code": "PA",
    }
    mappings_d, desc_d, code_d = map_specs_to_erp(rulebook, "DIO", specs_dio)
    assert code_d == "DIOSZ41S1000-PA"
    assert "Standart" not in desc_d

    # 2. Core attributes (like Package) must NEVER default to zero if unextracted
    specs_dio_no_pkg = {"mounting": "SMD", "type_diode": "Schottky", "value_diode": "SS0520"}
    mappings_np, _, code_np = map_specs_to_erp(rulebook, "DIO", specs_dio_no_pkg)
    pkg_m = next((m for m in mappings_np if m.attribute_name == "Package"), None)
    assert pkg_m is not None
    assert pkg_m.erp_code == "—"
    assert pkg_m.status == "BULUNAMADI"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_universal_default_special_definition_fallback -v`

- [ ] **Step 3: Implement generalized default qualifier fallback in `map_specs_to_erp`**

1. In `map_specs_to_erp` under `specialdefinition`:
   - If `specs.get("special_def")` matches a specific entry, assign `matched_entry`.
   - Otherwise, if an entry exists with `str(e.value) in ("0", "00")` or `e.value_description in (None, "", "Standart", "Standard")`:
     - Inspect sibling entries for this attribute. If `max(len(str(x.value)) for x in entries) >= 2`, format code as `"00"`, else `"0"`.
     - Assign `matched_entry = RuleEntry(..., value="00" or "0", value_description="Standart")`.
     - Mark status as `"TAM EŞLEŞTİ"`.
2. In `description_tokens` building:
   - Suppress `"Standart"`, `"Standard"`, `"None"`, `"General"` from `description_tokens` when derived from default empty rules.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_universal_default_special_definition_fallback -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: implement generic default qualifier fallback with peer-length padding and description suppression"
```

---

### Task 4: End-to-End Regression & Diverse Manufacturer Validation

**Files:**
- Modify: `tests/test_datasheet_analyzer.py`
- Test: `tests/`

**Interfaces:**
- Consumes: full pipeline across diverse components (`RES`, `CAP`, `DIO`, `CON`, `CRY`, `DSP`).
- Produces: 100% green test suite across all 151+ tests.

- [ ] **Step 1: Write end-to-end integration tests for newly supported dynamic manufacturers**

```python
# In tests/test_datasheet_analyzer.py
def test_diverse_dynamic_manufacturers_from_rules_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Test Qualitek (QA)
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "QUALITEK INTERNATIONAL")
    c.drawString(100, 730, "10K Resistor SMD 0603")
    c.save()
    res = analyze_datasheet(buf.getvalue(), part_number="RC0603FR-0710KL")
    assert res["specs"].get("manufacturer") == "Qualitek"
    assert res["erp_code"].endswith("-QA")
```

- [ ] **Step 2: Run integration tests**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_diverse_dynamic_manufacturers_from_rules_e2e -v`
Expected: PASS

- [ ] **Step 3: Run full regression test suite**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -m "not slow" -v`
Expected: PASS (all 151+ tests passing with zero failures)

- [ ] **Step 4: Commit**

```bash
git add tests/test_datasheet_analyzer.py
git commit -m "test: add dynamic manufacturer e2e tests and verify full regression suite"
```
