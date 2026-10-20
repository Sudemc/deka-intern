# Universal Rule Matcher & Display (DSP) Analyzer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement a Universal Semantic Rule Matcher in `map_specs_to_erp` to dynamically map unhandled ERP attributes from `rules.xlsx`, and build first-class Display (`DSP`) support producing `DSPTE4BH0028-FO` and rich description `TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard`.

**Architecture:** 
1. Root-Token Title Authority & High-Confidence PN Decoder: Adding `DSP` to `COVER_TITLE_PATTERNS` and Foryard `FY[QSTD]-` display decoder to `decode_manufacturer_pn`.
2. Universal Semantic Matcher in `map_specs_to_erp`: Dynamic candidate token collation, multilingual synonym mapping (`White` $\leftrightarrow$ `Beyaz`, `Common Cathode` $\leftrightarrow$ `COM-`, etc.), and word-boundary matching over candidate `RuleEntry` records for all unhandled SORTNO attributes.
3. Display Spec Extraction & Rich Description: Handling `TH` mounting, `7 Segment Display`, LED colors, digit counts, and polarity in `extract_component_specs` and `build_rich_description`.

**Tech Stack:** Python 3.10, regex, pytest, openpyxl.

**Spec:** `docs/superpowers/specs/2026-10-20-universal-rule-matcher-and-display-analyzer-design.md`

## Global Constraints

- 100% offline and local execution; zero external cloud API dependencies.
- ERP codes must remain 100% deterministic (generated strictly by rules.xlsx and map_specs_to_erp, never hallucinated by LLM).
- Preserves all 136 existing pytest tests without regression.
- Fast execution (<0.1s for fast-path / deterministic text).

## Review Focus

1. Standalone / Common Anode Display: A part number ending with `D` instead of `C` (e.g. `FYQ-5643DUW-21`) must map to `COM+` (`BG`) rather than `COM-` (`BH`).
2. Alternate LED Colors: Red (`UR`/`R`), Green (`UG`/`G`), Blue (`UB`/`B`), Yellow (`UY`/`Y`) must map to their respective DEKA ERP color codes (`1`, `3`, `2`, `5`).
3. Single/Dual/Triple Digit Display variants: `FYS-` (1 Digit), `FYD-` (2 Digit), `FYT-` (3 Digit) must map to their respective Special Definition codes (`BA`-`BF`).
4. Universal matcher token safety: Multilingual synonyms must use strict word boundary matching so substrings like `"red"` inside `"unregistered"` do not trigger false positive matches.
5. Backward compatibility on existing categories: Passive and IC categories (`RES`, `CAP`, `DIO`, `CON`, `CRY`) must not be diverted or altered by the universal fallback.

---

### Task 1: Display (`DSP`) Root-Token Title Authority, PN Decoder & Manufacturer Registration

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:250-270`, `src/bom_enricher/datasheet_analyzer.py:430-610`, `src/bom_enricher/datasheet_analyzer.py:1040-1080`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `part_number`, `text` in `classify_material_type` and `decode_manufacturer_pn`.
- Produces: `classify_material_type` returning `"DSP"`, 0.99 for display cover titles and `FYQ-` prefixes; `decode_manufacturer_pn` returning `mounting="TH"`, `type_display="7 Segment Display"`, `color="Beyaz"`, `digits="4 Digit"`, `polarity="COM-"`, `manufacturer="Foryard"`.

- [x] **Step 1: Write failing unit tests for Display classification and PN decoding**

```python
# In tests/test_datasheet_analyzer.py
def test_dsp_classification_from_title_and_pn():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "NINGBO FORYARD OPTOELECTRONICS CO.,LTD\n"
        "PRODUCT SPECIFICATION\n"
        "Model No.: FYQ-5643C/DX-XX\n"
        "Descriptions & Features:\n"
        "0.56Inch (14.20mm) digit height Four Digits Display\n"
        "Case mold type.\n"
        "Easy mounting on P.C. board or socket.\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "FYQ-5643CUW-21", "")
    assert mat_type == "DSP"
    assert conf >= 0.95


def test_foryard_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn
    # Common Cathode White 4-digit
    specs = decode_manufacturer_pn("FYQ-5643CUW-21")
    assert specs.get("manufacturer") == "Foryard"
    assert specs.get("mounting") == "TH"
    assert specs.get("type_display") == "7 Segment Display"
    assert specs.get("color") == "Beyaz"
    assert specs.get("polarity") == "COM-"
    assert specs.get("digits") == "4 Digit"

    # Common Anode Red 1-digit
    specs_anode = decode_manufacturer_pn("FYS-5611DUR-11")
    assert specs_anode.get("polarity") == "COM+"
    assert specs_anode.get("color") == "Kırmızı"
    assert specs_anode.get("digits") == "1 Digit"
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_dsp_classification_from_title_and_pn tests/test_datasheet_analyzer.py::test_foryard_pn_decoding -v`
Expected: FAIL (`DSP` unclassified, `decode_manufacturer_pn` returns `{}`)

- [x] **Step 3: Implement Display classification and Foryard decoder in `datasheet_analyzer.py`**

1. In `KNOWN_MANUFACTURERS`, add: `("foryard", "Foryard", "FO")`.
2. In `COVER_TITLE_PATTERNS`, add:
   ```python
   ("DSP", r"\b(?:(?:seven|7)[\s-]segment\s+displays?|digits?\s+displays?|character\s+lcd|graphic\s+lcd|tft\s+displays?|oled\s+displays?|dotmatrix\s+displays?)\b", "Display (DSP)")
   ```
3. In `classify_material_type`:
   - Add high-confidence PN prefix: `re.match(r"^FY[QSTD]-\d{4}", pn)` $\rightarrow$ `"DSP"`, 0.98.
   - Add `"DSP": 0.0` to `scores` dict.
   - Add display scoring: `_score("DSP", r"\b(?:7[\s-]segment|seven\s+segment|digits?\s+display|tft|lcd|oled)\b", 15.0)`.
4. In `decode_manufacturer_pn`:
   - Match `r"^FY([QSTD])-(\d{2})\d{2}([CD])([A-Z]+)"`:
     - Digit count: `Q` $\rightarrow$ `"4 Digit"`, `D` $\rightarrow$ `"2 Digit"`, `T` $\rightarrow$ `"3 Digit"`, `S` $\rightarrow$ `"1 Digit"`.
     - Polarity: `C` $\rightarrow$ `"COM-"`, `D` $\rightarrow$ `"COM+"`.
     - Color map: `{"UW": "Beyaz", "W": "Beyaz", "UR": "Kırmızı", "R": "Kırmızı", "UG": "Yeşil", "G": "Yeşil", "UB": "Mavi", "B": "Mavi", "UY": "Sarı", "Y": "Sarı"}`.
     - Set `type_display="7 Segment Display"`, `mounting="TH"`, `manufacturer="Foryard"`.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_dsp_classification_from_title_and_pn tests/test_datasheet_analyzer.py::test_foryard_pn_decoding -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: add Display (DSP) root-token title authority, Foryard decoder, and manufacturer registration"
```

---

### Task 2: Universal Semantic Matcher in `map_specs_to_erp` & Display Attributes

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:1650-1750`, `src/bom_enricher/datasheet_analyzer.py:1860-2250`, `src/bom_enricher/datasheet_analyzer.py:2740-2780`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: rulebook, `specs`, and `material_type` in `map_specs_to_erp`.
- Produces: `extract_component_specs` extracting display parameters; `build_rich_description` generating `TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard`; `map_specs_to_erp` dynamically matching unhandled attributes via Universal Semantic Matcher.

- [x] **Step 1: Write failing unit tests for Display spec extraction, rich description, and universal ERP mapping**

```python
# In tests/test_datasheet_analyzer.py
def test_universal_rule_matcher_dsp_attributes():
    from bom_enricher.datasheet_analyzer import load_rulebook, map_specs_to_erp
    rulebook = load_rulebook("data/rules.xlsx")
    specs = {
        "mounting": "TH",
        "type_display": "7 Segment Display",
        "color": "Beyaz",
        "special_def_display": "4 Digit, COM-",
        "serial_counter": "0028",
        "manufacturer": "Foryard",
    }
    mappings, desc, erp_code = map_specs_to_erp(rulebook, "DSP", specs)
    assert erp_code == "DSPTE4BH0028-FO"
    attr_map = {m.attribute_name: (m.matched_erp_value, m.erp_code) for m in mappings}
    assert attr_map["Mounting"][1] == "T"
    assert attr_map["Type Display"][1] == "E"
    assert attr_map["Color Display LED"][1] == "4"
    assert attr_map["Special Definition - Display "][1] == "BH"
    assert attr_map["Serial Display Number"][1] == "0028"
    assert attr_map["Q"][1] == "FO"


def test_dsp_rich_description():
    from bom_enricher.datasheet_analyzer import build_rich_description
    specs = {
        "mounting": "TH",
        "type_display": "7 Segment Display",
        "color": "Beyaz",
        "special_def_display": "4 Digit, COM-",
        "manufacturer": "Foryard",
    }
    desc = build_rich_description("DSP", specs, "")
    assert desc == "TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard"
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_universal_rule_matcher_dsp_attributes tests/test_datasheet_analyzer.py::test_dsp_rich_description -v`
Expected: FAIL (`erp_code != "DSPTE4BH0028-FO"`, `build_rich_description` does not format DSP)

- [x] **Step 3: Implement extraction, rich description, and universal semantic matcher in `datasheet_analyzer.py`**

1. In `extract_component_specs` under `elif material_type == "DSP":`:
   - Mounting: if `smd` in text, `"SMD"`; else `"TH"`.
   - Type Display: if `7 segment` / `seven segment` in text, `"7 Segment Display"`.
   - Color: extract from text/specs using color dictionary (`Beyaz`, `Kırmızı`, `Yeşil`, `Mavi`, `Sarı`).
   - Special Definition: combine digit count and polarity (e.g. `f"{digits}, {polarity}"` $\rightarrow$ `"4 Digit, COM-"`).
   - Set `specs["special_def_display"] = "4 Digit, COM-"`.
2. In `build_rich_description` for `DSP`:
   - Join: `f"{mounting} {type_display}, {color}, {special_def_display},  {mfr}"`.
3. In `CORE_ATTRIBUTES_BY_TYPE`, add:
   `"DSP": {"mounting", "typedisplay", "colordisplayled", "specialdefinitiondisplay", "q"}`
4. In `map_specs_to_erp`:
   - Define `MULTILINGUAL_SYNONYMS` table for Turkish/English electrical terms.
   - For any SORTNO attribute not caught by the dedicated passives (`mounting`, `package`, `value`, `voltage`, etc.), gather candidates from `specs` matching the attribute name or normalized value.
   - For each candidate entry, check:
     - Exact match with `e.value_description`.
     - Synonym expansion match (e.g. `specs["color"] == "Beyaz"` matches `e.value_description == "Beyaz,"`).
     - Substring/token overlap with word boundaries.
   - If matched, assign `matched_entry = e`, `status = "TAM EŞLEŞTİ"`.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_universal_rule_matcher_dsp_attributes tests/test_datasheet_analyzer.py::test_dsp_rich_description -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: implement universal semantic rule matcher and Display spec extraction"
```

---

### Task 3: End-to-End Real Foryard Datasheet Verification & Regression Suite

**Files:**
- Modify: `tests/test_datasheet_analyzer.py`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `analyze_datasheet` end-to-end with synthetic and PDF canvas inputs.
- Produces: Verified end-to-end result producing `DSPTE4BH0028-FO` and `DSPTE4BH0000XX-FO`.

- [x] **Step 1: Write end-to-end integration test for Foryard Display datasheet**

```python
# In tests/test_datasheet_analyzer.py
def test_real_foryard_display_datasheet_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Synthetic Foryard 4-Digit Display datasheet
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "NINGBO FORYARD OPTOELECTRONICS CO.,LTD")
    c.drawString(100, 730, "PRODUCT SPECIFICATION - Model No.: FYQ-5643C/DX-XX")
    c.drawString(100, 710, "0.56Inch (14.20mm) digit height Four Digits Display")
    c.drawString(100, 690, "Descriptions & Features: Case mold type. Easy mounting on P.C. board or socket.")
    c.drawString(100, 670, "Mechanical Dimensions: All pins are Φ0.51mm, 2.54*6=15.24")
    c.drawString(100, 650, "Internal Circuit Diagrams: FYQ-5643Cx Common Cathode, FYQ-5643Dx Common Anode")
    c.drawString(100, 630, "Electrical-Optical: UW Ultra White InGaN CCT:9500K Forward Voltage 2.80V")
    c.save()
    pdf_bytes = buf.getvalue()

    pn = "FYQ-5643CUW-21"
    
    # 1. Default counter test
    result = analyze_datasheet(pdf_bytes, part_number=pn)
    assert result["material_type"] == "DSP"
    assert result["specs"].get("mounting") == "TH"
    assert result["specs"].get("type_display") == "7 Segment Display"
    assert result["specs"].get("color") == "Beyaz"
    assert result["specs"].get("polarity") == "COM-"
    assert result["specs"].get("digits") == "4 Digit"
    assert result["specs"].get("manufacturer") == "Foryard"
    assert result["rich_description"] == "TH 7 Segment Display, Beyaz, 4 Digit, COM-,  Foryard"
    assert result["erp_code"] == "DSPTE4BH0000XX-FO"

    # 2. Injected internal counter test (0028)
    result_counter = analyze_datasheet(pdf_bytes, part_number=pn, serial_counter="0028")
    assert result_counter["erp_code"] == "DSPTE4BH0028-FO"
```

- [x] **Step 2: Run end-to-end integration test**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_real_foryard_display_datasheet_e2e -v`
Expected: PASS

- [x] **Step 3: Run full regression test suite**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -m "not slow" -v`
Expected: PASS (all 136+ tests passing with zero failures)

- [x] **Step 4: Commit**

```bash
git add tests/test_datasheet_analyzer.py
git commit -m "test: add end-to-end verification for Foryard display and full regression validation"
```
