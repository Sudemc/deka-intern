# Hybrid LLM Arbitration, Trigger Redesign & Transparent Collapsible UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transform the Hybrid LLM into an active Quality Auditor and Arbitrator that catches regex contradictions and validates complex components, providing complete transparency via a collapsible inspection accordion in the web UI.

**Architecture:** 
- In `VisionLLMClient`, add status and timeout tracking (`last_status`, text timeout 5.0s, vision timeout 30.0s).
- In `datasheet_analyzer.py`, implement category-mandatory auditing for complex components (`CON`, `SWT`, `REL`, `ICR`, etc.) and red-flag contradiction heuristics (e.g. text has "socket/receptacle" but regex got "Erkek"). Grant LLM arbitration authority over critical connector/mounting fields, logging structured `ArbitrationEvent` records and `llm_status` (`"active"`, `"standby"`, `"offline"`, `"timeout"`, `"error"`).
- In `web.py`, expose `/api/health/llm` and serialize arbitration traces in analysis responses.
- In `static/index.html`, render a `<details class="llm-trace-accordion">` element (closed by default) that expands upon click to display status, arbitration decision logs, and side-by-side parameter comparisons.

**Tech Stack:** Python 3.11, FastAPI, Uvicorn, Ollama (`llama3.2:3b`, `moondream`), HTML5/CSS3 (vanilla JS), Pytest.

**Spec:** [2026-10-22-hybrid-llm-arbitration-and-transparent-ui-design.md](file:///c:/Users/sudem/OneDrive/Belgeler/deka/docs/superpowers/specs/2026-10-22-hybrid-llm-arbitration-and-transparent-ui-design.md)

## Global Constraints

- 100% offline and local execution; zero external cloud API dependencies.
- ERP codes must remain 100% deterministic (grounded in `rules.xlsx` and `map_specs_to_erp`).
- Fast text LLM timeout capped at 5.0 seconds maximum to prevent hanging during batch processing.
- All 182 existing tests across 16 test files must pass with zero regressions.
- Frontend inspection details must remain closed by default to preserve a clean, uncluttered interface.

## Review Focus

1. **Silent Fallback Deception:** When Ollama is offline or times out, `llm_status` must record `"offline"` or `"timeout"` and never falsely claim `"active"` or crash.
2. **Standard Passive Performance:** Standard chip resistors/capacitors with exact algorithmic PN decoding (`conf >= 0.98`) must skip LLM to preserve instant speed, recording `llm_status = "standby"`.
3. **Red-Flag Contradiction Override:** When text contains "socket" or "receptacle" but regex extracted `pin_type: "Erkek"`, LLM auditor must engage, override `pin_type` to `"Dişi"`, and record an `llm_override` arbitration event.
4. **Non-Connector Guarding:** Non-connector categories (`RES`, `CAP`, `DIO`, `SWT`, etc.) must not have their valid deterministic parameters corrupted by speculative LLM guesses; LLM fills empty fields or discovers missing part numbers/manufacturers.
5. **UI Accessibility & State Preservation:** Clicking the collapsible `<details>` element in `index.html` must smoothly reveal the full diagnostic data without affecting copy buttons or result formatting.

---

### Task 1: VisionLLMClient Status Tracking & Text Timeout Support

**Files:**
- Modify: `src/bom_enricher/vision_llm.py:40-150`
- Test: `tests/test_vision_client.py`

**Interfaces:**
- Produces: `VisionLLMClient.last_status: str` (`"uninitialized"`, `"success"`, `"timeout"`, `"connection_error"`, `"parse_error"`, `"empty_response"`)
- Produces: `VisionLLMClient.extract_specs_from_text(..., timeout: float = 5.0)`

- [ ] **Step 1: Write the failing tests in `tests/test_vision_client.py`**

```python
def test_vision_llm_client_status_tracking():
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()
    assert client.last_status == "uninitialized"

def test_extract_specs_from_text_timeout_handling(monkeypatch):
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()
    
    def mock_post(*args, **kwargs):
        raise httpx.TimeoutException("Read timed out")
    
    monkeypatch.setattr(httpx, "post", mock_post)
    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is None
    assert client.last_status == "timeout"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_vision_client.py::test_vision_llm_client_status_tracking tests/test_vision_client.py::test_extract_specs_from_text_timeout_handling -v`
Expected: FAIL (`last_status` attribute missing).

- [ ] **Step 3: Implement status tracking and timeout handling in `src/bom_enricher/vision_llm.py`**

Initialize `self.last_status: str = "uninitialized"` in `VisionLLMClient.__init__`.
In `extract_specs_from_text(self, text: str, target_pn: str | None = None, material_type_hint: str | None = None, timeout: float = 5.0)`:
- Catch `httpx.TimeoutException` -> `self.last_status = "timeout"` -> return `None`.
- Catch `httpx.ConnectError` -> `self.last_status = "connection_error"` -> return `None`.
- On JSON decode error or empty response -> `self.last_status = "parse_error"` / `"empty_response"`.
- On successful parsing -> `self.last_status = "success"` -> return parsed dict.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_vision_client.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/vision_llm.py tests/test_vision_client.py
git commit -m "feat: add last_status tracking and timeout configuration in VisionLLMClient"
```

---

### Task 2: Mandatory Auditor Triggers, Red-Flag Heuristics & Arbitration Engine

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:5190-5325`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `VisionLLMClient.last_status` from Task 1.
- Produces: `ArbitrationEvent` dicts in `result["llm_trace"]`: `[{"attribute": str, "regex_val": Any, "llm_val": Any, "final_val": Any, "reason": str, "source": str}]`
- Produces: `result["llm_status"]`: `"active"` | `"standby"` | `"offline"` | `"timeout"` | `"error"`
- Produces: `result["llm_specs"]`, `result["regex_specs"]`

- [ ] **Step 1: Write the failing tests in `tests/test_datasheet_analyzer.py`**

```python
def test_red_flag_gender_contradiction_triggers_llm_arbitration():
    from src.bom_enricher.datasheet_analyzer import analyze_datasheet
    # Mock text with Socket Header where regex would guess Erkek, but LLM arbitrates to Dişi
    class MockVLM:
        last_status = "success"
        def is_available(self): return True
        def extract_specs_from_text(self, text, **kwargs):
            return {"specs": {"pin_type": "Dişi", "mounting": "SMD", "type_connector": "Pin Header"}}
    
    text = "WR-PHD 2.54mm SMT Socket Header 4 Pin 61000413321"
    res = analyze_datasheet("61000413321", text=text, client=MockVLM())
    assert res["llm_status"] == "active"
    assert res["specs"]["pin_type"] == "Dişi"
    assert any(ev["attribute"] == "pin_type" and ev["source"] == "llm_override" for ev in res["llm_trace"])

def test_standard_chip_passive_bypasses_llm_with_standby_status():
    from src.bom_enricher.datasheet_analyzer import analyze_datasheet
    class MockVLM:
        last_status = "success"
        def is_available(self): return True
        def extract_specs_from_text(self, text, **kwargs):
            raise AssertionError("LLM should not be called for pure standard chip passive!")
    
    text = "RC0603FR-0710KL Yageo 10K Resistor SMD 0603 1%"
    res = analyze_datasheet("RC0603FR-0710KL", text=text, client=MockVLM())
    assert res["llm_status"] == "standby"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_datasheet_analyzer.py::test_red_flag_gender_contradiction_triggers_llm_arbitration tests/test_datasheet_analyzer.py::test_standard_chip_passive_bypasses_llm_with_standby_status -v`
Expected: FAIL (`llm_status` or `llm_trace` missing or AssertionError).

- [ ] **Step 3: Implement red-flag triggers, mandatory auditing, and arbitration in `src/bom_enricher/datasheet_analyzer.py`**

1. Preserve copy of raw regex specs: `regex_specs = copy.deepcopy(specs)`.
2. Compute `is_pure_passive` (standard `RES`/`CAP` with `pn_has_core` and `conf >= 0.98` and not catalog).
3. Compute `has_red_flag` (gender contradiction, mounting contradiction).
4. Compute `should_trigger_text_llm`:
   ```python
   should_trigger_text_llm = (
       vlm.is_available()
       and bool(clean_text)
       and (
           has_red_flag
           or not is_pure_passive
           or is_complex_or_ambiguous
           or are_critical_missing
           or is_catalog
           or not pn
           or pn.startswith("media_")
           or conf < 0.98
       )
   )
   ```
5. If `should_trigger_text_llm`:
   Call `vlm.extract_specs_from_text(clean_text[:4000], target_pn=pn, material_type_hint=mat_type, timeout=5.0)`.
   Execute arbitration rules for `pin_type`, `mounting`, `type_connector`, `style`, `pin_number`, and empty fields.
   Record each event into `llm_trace`.
   Set `llm_status = "active" if used_text_llm else vlm.last_status`.
6. Else:
   Set `llm_status = "standby"` if `vlm.is_available()` else `"offline"`.
7. Propagate `llm_status`, `llm_trace`, `llm_specs`, `regex_specs` into return dict and `DatasheetAnalysisResult`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_datasheet_analyzer.py::test_red_flag_gender_contradiction_triggers_llm_arbitration tests/test_datasheet_analyzer.py::test_standard_chip_passive_bypasses_llm_with_standby_status -v`
Expected: PASS.

- [ ] **Step 5: Run full test suite to guarantee zero regressions**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/ -m "not slow"`
Expected: All 182+ tests passing.

- [ ] **Step 6: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: implement mandatory auditor triggers, red-flag heuristics, and LLM arbitration engine"
```

---

### Task 3: Backend Health API & Payload Serialization in `web.py`

**Files:**
- Modify: `src/bom_enricher/web.py:150-260`
- Test: `tests/test_web.py`

**Interfaces:**
- Consumes: `vlm.is_available()`, `result["llm_status"]`, `result["llm_trace"]` from Task 2.
- Produces: `GET /api/health/llm` endpoint returning `{"status": "online" | "offline", "text_model": "llama3.2:3b", "vision_model": "moondream:latest", "host": str}`
- Produces: API response for `/api/datasheet-analyze` including `llm_status`, `llm_trace`, `llm_specs`, `regex_specs`.

- [ ] **Step 1: Write the failing tests in `tests/test_web.py`**

```python
def test_llm_health_endpoint():
    from fastapi.testclient import TestClient
    from src.bom_enricher.web import app
    client = TestClient(app)
    resp = client.get("/api/health/llm")
    assert resp.status_code == 200
    data = resp.json()
    assert "status" in data
    assert "text_model" in data
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_web.py::test_llm_health_endpoint -v`
Expected: FAIL with 404 Not Found.

- [ ] **Step 3: Implement `/api/health/llm` and result serialization in `src/bom_enricher/web.py`**

1. In `src/bom_enricher/web.py`, add:
   ```python
   @app.get("/api/health/llm", tags=["Health"])
   async def get_llm_health() -> dict[str, object]:
       vlm = VisionLLMClient()
       is_online = vlm.is_available()
       return {
           "status": "online" if is_online else "offline",
           "text_model": "llama3.2:3b",
           "vision_model": "moondream:latest",
           "host": "http://127.0.0.1:11434",
       }
   ```
2. In `analyze_single_datasheet`, include `llm_status`, `llm_trace`, `llm_specs`, `regex_specs` in `DatasheetAnalysisResult.to_dict()` and `analyses[0].to_dict()` returned by `/api/datasheet-analyze`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_web.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/web.py tests/test_web.py
git commit -m "feat: add LLM health endpoint and propagate arbitration trace in web API"
```

---

### Task 4: Collapsible Inspection Accordion & Status Badges in `static/index.html`

**Files:**
- Modify: `src/bom_enricher/static/index.html:1980-2110`
- Test: `tests/test_web.py`

**Interfaces:**
- Consumes: `item.llm_status`, `item.llm_trace`, `item.llm_specs`, `item.regex_specs` from Task 3.
- Produces: Collapsible `<details class="llm-trace-accordion">` rendered inside each `.component-card` in `static/index.html`.

- [ ] **Step 1: Write verification test for index.html accordion template**

```python
def test_index_html_contains_llm_accordion_template():
    from pathlib import Path
    html_text = Path("src/bom_enricher/static/index.html").read_text(encoding="utf-8")
    assert "llm-trace-accordion" in html_text
    assert "Yapay Zeka (LLM) & Karar Geçmişi" in html_text
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_web.py::test_index_html_contains_llm_accordion_template -v`
Expected: FAIL (`llm-trace-accordion` not found).

- [ ] **Step 3: Implement collapsible `<details>` and styling in `src/bom_enricher/static/index.html`**

1. In the component card rendering function in `index.html`:
   - Generate status badge:
     - `item.llm_status === 'active'` -> `🟢 LLM Aktif (Doğrulandı)`
     - `item.llm_status === 'standby'` -> `🔵 Standby (Kural Motoru Yeterli)`
     - `item.llm_status === 'timeout'` -> `⏱️ Zaman Aşımı`
     - else -> `🟡 Ollama Kapalı (Deterministik Mod)`
   - Construct `<details class="llm-trace-accordion" style="...">`:
     - `<summary class="llm-trace-summary">`: Clickable header with chevron icon and status badge. Closed by default.
     - Content container:
       - Informative banner explaining whether LLM verified the component, was kept on standby, or fell back due to Ollama being offline.
       - Arbitration Decision Table: Columns for Attribute, Regex Value, LLM Value, Final Value, Reason, Source.
       - Raw parameter inspection: side-by-side comparison of Regex specs vs LLM specs.
2. Add clean CSS for `.llm-trace-accordion`, `.llm-trace-summary`, `.llm-table`, and pills.

- [ ] **Step 4: Run test to verify it passes**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_web.py::test_index_html_contains_llm_accordion_template -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bom_enricher/static/index.html tests/test_web.py
git commit -m "feat: add collapsible LLM arbitration accordion and status badges in web UI"
```

---

### Task 5: End-to-End Real Datasheet Verification & Regression Suite

**Files:**
- Test: `tests/test_datasheet_analyzer.py`
- All regression suites across `tests/`

- [ ] **Step 1: Write integration test for Würth Socket Header and live LLM arbitration**

```python
def test_wurth_socket_header_e2e_arbitration_and_trace():
    from src.bom_enricher.datasheet_analyzer import analyze_datasheet
    from src.bom_enricher.comparison import load_rulebook
    from pathlib import Path
    
    rulebook = load_rulebook(Path("rules.xlsx"))
    text = "WR-PHD 2.54mm SMT Socket Header 4 Pin 61000413321 Wurth Elektronik"
    res = analyze_datasheet("61000413321", text=text, rulebook=rulebook)
    
    # Must correctly resolve to Female connector code
    assert res["erp_code"].startswith("CONSAPBD04")
    assert res["erp_code"].endswith("-WU")
    assert res["specs"]["pin_type"] == "Dişi"
    assert "llm_trace" in res
    assert "llm_status" in res
```

- [ ] **Step 2: Run test to verify it passes**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/test_datasheet_analyzer.py::test_wurth_socket_header_e2e_arbitration_and_trace -v`
Expected: PASS.

- [ ] **Step 3: Run full master regression suite**

Run: `.\\.venv\\Scripts\\python.exe -m pytest tests/ -m "not slow"`
Expected: All 182+ tests across all 16 test files pass cleanly with 0 failures, 0 regressions.

- [ ] **Step 4: Commit**

```bash
git add tests/test_datasheet_analyzer.py
git commit -m "test: verify end-to-end socket header LLM arbitration and full regression suite"
```
