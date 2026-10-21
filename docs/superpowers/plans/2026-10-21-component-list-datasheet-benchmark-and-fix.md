# Component List Datasheet Benchmarking & Systematic Accuracy Verification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Systematically search, test, and align all electronic components in the user's master BOM list against `analyze_datasheet` to guarantee 100% deterministic ERP code generation and accurate Turkish rich descriptions per `rules.xlsx`, cleanly skipping items without available datasheets as requested.

**Architecture:** 
1. Group the 202 components by category (`ADP`, `ANT`, `BAH`, `CAP`, `BUZ`, `DIO`, `DSP`, `CRY`, `CON`, and `CBL`/`CHZ`/`DMB`).
2. Search web/online datasheets for each part number; skip non-datasheet items (such as office furniture `DMB` and custom internal cable harnesses `CBL`).
3. For found datasheets, test against `analyze_datasheet` and compare generated `erp_code` and `rich_description` with user expectations.
4. Enhance `decode_manufacturer_pn`, `extract_component_specs`, and `map_specs_to_erp` systemically whenever gaps are identified.
5. Guard all changes with unit/integration regression tests, maintaining zero regressions across the 173+ test suite.

**Tech Stack:** Python 3.10+, `bom_enricher.datasheet_analyzer`, `bom_enricher.comparison`, `pytest`, `reportlab`, `rules.xlsx`.

**Spec:** User instruction provided in conversation (202-item ERP list benchmarking request).

## Global Constraints
- 100% offline and deterministic ERP code generation: strictly derived from `rules.xlsx` and `map_specs_to_erp`.
- Zero regressions across existing test suite (all 173+ tests must remain passing).
- Skip components where datasheets cannot be found online as explicitly instructed.
- Preserve all existing comments and docstrings.
- Web server daemon running on port 8000 must be restarted upon completing batches.

## Review Focus
1. **Manufacturer Code Alignment:** Ensuring manufacturer names extracted from datasheets/PNs map to exact DEKA 2-3 letter manufacturer codes in `rules.xlsx`.
2. **Mounting & Package Disambiguation:** Distinguishing SMD vs TH, and specific packages (e.g. SOD-123 vs SOD-123W vs Minimelf vs SOT-23).
3. **Sequential Counter Preservation:** For components using sequential counters (`0000XX` or injected values like `0021`), preserving the exact prefix tokens and manufacturer suffix.
4. **Description Formatting:** Ensuring `rich_description` contains clean Turkish specifications without token pollution (e.g., no synthetic `"Standart"`, `"None"`).
5. **Datasheet Omission Protocol:** Cleanly documenting and skipping items without electronic datasheets without breaking batch processing.

---

### Task 1: Discrete Passives & Power: `ADP` (3), `ANT` (4), `BAH` (5), `CAP` (5)

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Modify: `src/bom_enricher/comparison.py`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `analyze_datasheet(file_bytes=..., part_number=..., expected_type=...)`
- Produces: `erp_code`, `specs`, `rich_description`

- [ ] **Step 1: Write test cases for Task 1 components**
  Test `ADP` (`20163` Kaizen, `CYS 1206-M` Yıldırım, `SK03T1200200V` Simsukian), `ANT` (`2108857-4` TE, `MAF95310` Ezurio, `EMF2449A1-10UFL` Ezurio, `2108857-1` TE), `BAH` (`L-KLS5-CR2032-09`, `L-KLS5-CR2032-11`, `DS1092-11-N8S`, `L-KLS5-SP-073-R`, `55` Keystone), and `CAP` (`GRT155R61C105KE01D`, `0402B102J160CT`, `CC0402JRX7R7BB102`, `CL05B104KO5VPND`, `GRM155R71C103KA01D`).
- [ ] **Step 2: Run test to inspect current results and identify discrepancies**
- [ ] **Step 3: Implement fixes in `datasheet_analyzer.py` for any gaps found in Task 1**
- [ ] **Step 4: Re-run tests to verify all Task 1 components pass**
- [ ] **Step 5: Commit Task 1 changes**

---

### Task 2: Acoustic & Semiconductor Passives: `BUZ` (15) & `DIO` (14)

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Modify: `src/bom_enricher/comparison.py`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `analyze_datasheet`
- Produces: Accurate buzzer sound/frequency/size mapping and diode voltage/current/package/special-definition mapping.

- [ ] **Step 1: Write test cases for Task 2 components (`BUZ` and `DIO`)**
  Test 15 buzzers (`FUET-5020`, `CR-HM-013-0006`, `YX7525P`, `BPT1703H09-02`, `KPT-G1713-K11236`, `KPT-G2312-K7462`, `EPT-22A2003P`, `PT2770AHP-06`, `YX-YT9503`, `YX-YT12065-YT05D`, `HPT-1780P-05TB`, `HNB09A05`, `YX-WY1470A`, etc.) and 14 diodes (`BZV55-C20`, `PMEG2010ER`, `BAS19LT1G`, `P6SMBJ120`, `SMBJ20A`, `NSR0520V2`, `LMBR0520T1G`, `NZH20C`, `MBR0520-HT`, `SS0520`, `PSR05-LF-T7`, `SMAJ30CA`).
- [ ] **Step 2: Run test to identify any buzzer/diode discrepancies**
- [ ] **Step 3: Implement fixes in `datasheet_analyzer.py` / `comparison.py`**
- [ ] **Step 4: Re-run tests to verify all Task 2 components pass**
- [ ] **Step 5: Commit Task 2 changes**

---

### Task 3: Visual Output Devices: `DSP` (24 Displays)

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Modify: `src/bom_enricher/comparison.py`
- Test: `tests/test_datasheet_analyzer.py`

- [ ] **Step 1: Write test cases for 24 displays**
  Cover character LCDs (`MTB0802A2`, `WH0802A1`, `EA DOGM162E-A`), 7-segment displays (Liteon, Strong Base, Foryard, Bright Led, Avago, Wuxi Compul, Kinley, Assemcorp, Semtech), backlight (`EA LED55X31-G`), HMI (`NX4827P043-011C-Y`).
- [ ] **Step 2: Run test to evaluate display decoder and semantic matcher output**
- [ ] **Step 3: Implement decoder/matcher enhancements for displays**
- [ ] **Step 4: Re-run tests to verify all Task 3 components pass**
- [ ] **Step 5: Commit Task 3 changes**

---

### Task 4: Frequency Control: `CRY` (66 Crystals, Oscillators, Resonators)

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Modify: `src/bom_enricher/comparison.py`
- Test: `tests/test_datasheet_analyzer.py`

- [ ] **Step 1: Write test cases for 66 crystal components**
  Cover Epson, ECS, HCI, Abracon, Seiko, Raltron, Murata (`CSTNE`, `CSTLS`, `XRCGB`), Kyocera AVX, Jauch, NDK, TXC, Citizen, CY, etc.
- [ ] **Step 2: Run test to identify any frequency/load capacitance/ESR/package discrepancies**
- [ ] **Step 3: Implement fixes in `datasheet_analyzer.py` / `comparison.py`**
- [ ] **Step 4: Re-run tests to verify all Task 4 components pass**
- [ ] **Step 5: Commit Task 4 changes**

---

### Task 5: Interconnects: `CON` (39 Connectors, Headers, Terminals)

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Modify: `src/bom_enricher/comparison.py`
- Test: `tests/test_datasheet_analyzer.py`

- [ ] **Step 1: Write test cases for 39 connector components**
  Cover wire-to-wire (Degson, Tianli), pin headers (KLS, Connfly `DS1023`), IDC (Stocko), wire-to-board (Wago, JST), board-to-board (Amphenol), FFC/FPC (Molex, Oupiin, Hirose `FH35C`, GCT, XUNPU, Amphenol, Wurth), box headers, modular jacks, USB.
- [ ] **Step 2: Run test to identify pitch, pin count, orientation, and manufacturer discrepancies**
- [ ] **Step 3: Implement fixes in `datasheet_analyzer.py` / `comparison.py`**
- [ ] **Step 4: Re-run tests to verify all Task 5 components pass**
- [ ] **Step 5: Commit Task 5 changes**

---

### Task 6: Custom Cables, Devices (`CBL`, `CHZ`, `DMB`) & Full Regression Verification

**Files:**
- Test: `tests/test_datasheet_analyzer.py`
- Process: Web server daemon restart

- [ ] **Step 1: Audit non-datasheet components (`DMB` fixtures, `CBL` internal wire harnesses, `CHZ`) and confirm skip protocol**
- [ ] **Step 2: Run full regression test suite across the entire project (`pytest tests/ -m "not slow"`)**
- [ ] **Step 3: Restart web server daemon on port 8000**
- [ ] **Step 4: Commit final verification and documentation**
