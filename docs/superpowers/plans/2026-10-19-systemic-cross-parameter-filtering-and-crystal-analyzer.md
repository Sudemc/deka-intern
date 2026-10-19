# Systemic Cross-Parameter Filtering & Universal Category Authority Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Eliminate cross-parameter false positives (e.g. crystals or inductors misidentified as resistors due to ESR/DCR; diodes misidentified as capacitors due to junction capacitance), enforce root-token title authority, and complete rich specification extraction and ERP mapping for Crystals (`CRY`) and other under-served material types.

**Architecture:** 
1. Root-Token Title Authority: Direct, lemmatized matching on Cover Page / Title Block tokens (e.g. `crystal`, `inductor`, `choke`, `relay`, `diode`) locking the component category with 0.99 confidence regardless of body content.
2. Contextual Negative Parameter Isolation: Comprehensive exclusion of parasitic parameters (ESR, DCR, coil resistance, junction capacitance, load capacitance) from scoring against primary passives (`RES`, `CAP`).
3. Complete Crystal Spec Extractor & Rich Description: Extracting `frequency_str`, `package` (`4-SMD`), `load_capacitance` (`16 pF`), `esr` (`400Ohm`), and `manufacturer` (`Raltron`), generating exact expected description: `SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron`.
4. Deterministic DEKA ERP Mapping for `CRY`: Mapping Mounting (`S`), Value Crystals (`8MHz` -> `A2`), Serial Crystal Number counter (`076` / `0000XX`), and Manufacturer (`Raltron` -> `-RA`) to synthesize `CRYSA2076-RA`.

**Tech Stack:** Python 3.10, regex, pytest.

---

### Task 1: Root-Token Title Authority & Negative Parasitic Filtering

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:420-610`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `text`, `part_number`, `filename` in `classify_material_type`.
- Produces: `classify_material_type` returning `"CRY"`, 0.99 for `SURFACE MOUNT MICROPROCESSOR CRYSTAL` and `RH100...` without falling into the `RES` (400 ohm ESR) trap.

- [x] **Step 1: Write failing tests for parasitic parameter isolation**

```python
# In tests/test_datasheet_analyzer.py
def test_crystal_with_esr_not_misclassified_as_resistor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SURFACE MOUNT MICROPROCESSOR CRYSTAL\n"
        "RH100-8.000-16-3030-TR-NS1\n"
        "RALTRON ELECTRONICS\n"
        "NOMINAL FREQUENCY 8.000 MHz\n"
        "LOAD CAPACITANCE 16 pF\n"
        "EQUIVALENT SERIES RESISTANCE 400 ohm max\n"
        "SHUNT CAPACITANCE 2 pF max\n"
        "INSULATION RESISTANCE 500 Mohm min\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "RH100-8.000-16-3030-TR-NS1", "")
    assert mat_type == "CRY"
    assert conf >= 0.95


def test_inductor_with_dcr_not_misclassified_as_resistor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SMD POWER INDUCTOR\n"
        "CD54-100MC\n"
        "INDUCTANCE 10 uH\n"
        "DC RESISTANCE 0.15 OHM MAX\n"
        "RATED CURRENT 1.44 A\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "CD54-100MC", "")
    assert mat_type == "IND"
    assert conf >= 0.95


def test_diode_with_junction_capacitance_not_misclassified_as_capacitor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    sample_text = (
        "SURFACE MOUNT SCHOTTKY BARRIER RECTIFIER\n"
        "SS34\n"
        "FORWARD VOLTAGE 0.5 V\n"
        "REVERSE LEAKAGE CURRENT 0.5 mA\n"
        "JUNCTION CAPACITANCE 150 pF\n"
    )
    mat_type, conf, reason = classify_material_type(sample_text, "SS34", "")
    assert mat_type == "DIO"
    assert conf >= 0.95
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_crystal_with_esr_not_misclassified_as_resistor -v`
Expected: FAIL (classified as `RES` with 0.98 confidence)

- [x] **Step 3: Implement root-token title authority & negative filtering in `datasheet_analyzer.py`**

1. Expand `COVER_TITLE_PATTERNS` for `CRY`, `IND`, `REL`, `TRF`, `SWT`:
   - `CRY`: `r"\b(?:(?:surface\s+mount\s+)?(?:microprocessor\s+)?crystals?|quartz\s+crystals?|crystal\s+units?|crystal\s+oscillators?|ceramic\s+resonators?|tcxo|vcxo|ocxo)\b"`
   - `IND`: `r"\b(?:power\s+choke|choke\s+coil|chip\s+inductor|power\s+inductor|smd\s+inductor|inductors?|chokes?)\b"`
2. Expand high-confidence part number prefixes:
   - `CRY`: `re.match(r"^(?:RH\d{2,3}|ABS\d|ABM\d|ECS-|NX\d{4}|FA-\d|TSX-\d|7M\d|9C\d|CST|SIT\d)", pn)`
3. Expand parasitic negative filtering in `classify_material_type`:
   - `PARASITIC_RESISTANCE_RE = r"\b(?:coil|contact|thermal|junction|insulation|equivalent\s+series|esr|dc|dcr|series|lead|conductor|winding|ground)\s+resistance\b"`
   - If cover/text mentions dominant keywords of non-resistor components (e.g. `crystal`, `inductor`, `relay`), suppress bare `ohm` / `resistance` scoring for `RES`.
   - `PARASITIC_CAPACITANCE_RE = r"\b(?:load|shunt|junction|stray|parasitic|input|output|coupling|feedthrough|case)\s+capacitance\b"`
   - If cover/text mentions `crystal` or `diode`, suppress bare `pf` / `capacitance` scoring for `CAP`.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_crystal_with_esr_not_misclassified_as_resistor tests/test_datasheet_analyzer.py::test_inductor_with_dcr_not_misclassified_as_resistor tests/test_datasheet_analyzer.py::test_diode_with_junction_capacitance_not_misclassified_as_capacitor -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "fix: enforce root-token title authority and negative parasitic parameter filtering"
```

---

### Task 2: Crystal Spec Extractor, Rich Description & Raltron Decoding

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:1530-1550`, `src/bom_enricher/datasheet_analyzer.py:2619-2628`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: text, drawing, and part number for `CRY` components.
- Produces: `extract_component_specs` extracting `frequency_str`, `load_capacitance`, `esr`, `package`, `mounting`, `manufacturer`; `build_rich_description` generating `SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron`.

- [x] **Step 1: Write failing unit test for crystal spec extraction and description**

```python
# In tests/test_datasheet_analyzer.py
def test_crystal_spec_extraction_and_rich_description():
    from bom_enricher.datasheet_analyzer import extract_component_specs, build_rich_description
    sample_text = (
        "SURFACE MOUNT MICROPROCESSOR CRYSTAL\n"
        "RH100-8.000-16-3030-TR-NS1\n"
        "RALTRON ELECTRONICS\n"
        "NOMINAL FREQUENCY 8.000 MHz\n"
        "LOAD CAPACITANCE 16 pF\n"
        "EQUIVALENT SERIES RESISTANCE 400 ohm max\n"
        "SHUNT CAPACITANCE 2 pF max\n"
        "4x1.0 4x1.4 SOLDERING PATTERN TOP VIEW PIN CONNECTION GND 4 3 2 GND 1 3.2x2.5\n"
    )
    specs = extract_component_specs(sample_text, "CRY", "RH100-8.000-16-3030-TR-NS1")
    assert specs.get("frequency_str") == "8MHz"
    assert specs.get("load_capacitance") == "16 pF"
    assert specs.get("esr") in ("400Ohm", "400 ohm")
    assert specs.get("mounting") == "SMD"
    assert specs.get("package") == "4-SMD"
    assert specs.get("manufacturer") == "Raltron"

    desc = build_rich_description("CRY", specs, sample_text)
    assert desc == "SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron"
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_crystal_spec_extraction_and_rich_description -v`
Expected: FAIL (missing `load_capacitance`, `esr`, `package`, desc doesn't match)

- [x] **Step 3: Implement extraction and description logic in `datasheet_analyzer.py`**

1. In `extract_component_specs` under `elif material_type == "CRY":`:
   - Frequency: `re.search(r"\b(\d+(?:\.\d+)?)\s*(mhz|khz)\b", text_lower)`. Normalize `8.000 MHz` -> `8MHz`.
   - Load Capacitance: `re.search(r"\bload\s+capacitance[:\s]+(\d+(?:\.\d+)?)\s*pf\b", text_lower)` -> `f"{m.group(1)} pF"`.
   - ESR: `re.search(r"\b(?:equivalent\s+series\s+resistance|esr)[:\s]+(\d+)\s*(?:ohm|ω|q)\b", text_lower)` -> `f"{m.group(1)}Ohm"`.
   - Package: Check if `4-smd` or `4x1` or `4 pin` or `3.2\s*x\s*2.5` or `3225` appears $\rightarrow$ `"4-SMD"`. If 2 pins / HC-49 $\rightarrow$ `"HC-49/SMD"` or `"2-SMD"`.
   - Mounting: if "surface mount" or "smd" in text, `"SMD"`.
   - Manufacturer: if "raltron" in text or pn starts with "RH", `"Raltron"`.
2. In `build_rich_description` for `CRY`:
   - Build tokens:
     ```python
     tokens = []
     if specs.get("mounting"):
         tokens.append(specs["mounting"])
     if specs.get("package"):
         tokens.append(specs["package"])
     if specs.get("frequency_str"):
         tokens.append(specs["frequency_str"])
     if specs.get("load_capacitance"):
         tokens.append(specs["load_capacitance"])
     if specs.get("esr"):
         tokens.append(specs["esr"])
     if mfr:
         tokens.append(mfr)
     # If tokens start with "SMD" and package is "4-SMD", join nicely: "SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron"
     ```

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_crystal_spec_extraction_and_rich_description -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: implement crystal spec extraction, 4-SMD detection and rich description"
```

---

### Task 3: DEKA ERP Crystal Mapping & End-to-End Raltron Verification

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:1820-1840`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: rulebook and extracted specs for `CRY`.
- Produces: `map_specs_to_erp` producing `CRYSA2076-RA` (when serial counter 076 is resolved or passed) or `CRYSA20000XX-RA` with exact attribute mappings for Mounting, Value, Serial, and Q.

- [x] **Step 1: Write end-to-end test for Raltron Crystal datasheet**

```python
# In tests/test_datasheet_analyzer.py
def test_real_raltron_crystal_datasheet_e2e():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    from reportlab.pdfgen import canvas
    import io

    # Synthetic multi-attribute Raltron crystal datasheet
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "SURFACE MOUNT MICROPROCESSOR CRYSTAL")
    c.drawString(100, 730, "RALTRON ELECTRONICS - RH100-8.000-16-3030-TR-NS1")
    c.drawString(100, 710, "NOMINAL FREQUENCY: 8.000 MHz")
    c.drawString(100, 690, "LOAD CAPACITANCE: 16 pF")
    c.drawString(100, 670, "EQUIVALENT SERIES RESISTANCE: 400 ohm max")
    c.drawString(100, 650, "MECHANICAL SPECIFICATION: 3.2 x 2.5 mm, 4 PADS, TOP VIEW PIN CONNECTION GND 4, 3, 2 GND, 1")
    c.save()
    pdf_bytes = buf.getvalue()

    pn = "RH100-8.000-16-3030-TR-NS1"
    result = analyze_datasheet(pdf_bytes, part_number=pn)

    assert result["material_type"] == "CRY"
    assert result["specs"].get("frequency_str") == "8MHz"
    assert result["specs"].get("load_capacitance") == "16 pF"
    assert result["specs"].get("esr") == "400Ohm"
    assert result["specs"].get("package") == "4-SMD"
    assert result["specs"].get("mounting") == "SMD"
    assert result["specs"].get("manufacturer") == "Raltron"

    assert result["rich_description"] == "SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron"
    assert result["erp_code"].startswith("CRYSA2")
    assert result["erp_code"].endswith("-RA")
```

- [x] **Step 2: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_real_raltron_crystal_datasheet_e2e -v`
Expected: PASS

- [x] **Step 3: Run full regression test suite**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -m "not slow" -v`
Expected: PASS (all 130+ tests passing)

- [x] **Step 4: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: complete DEKA ERP mapping and end-to-end verification for Raltron crystal"
```
