# Vision LLM Datasheet Analyzer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an end-to-end multimodal datasheet analysis pipeline combining fast-path PN decoding, smart PDF page rendering (pypdfium2), a local Vision LLM client (Ollama / Qwen2-VL / Llama-3.2-Vision), and deterministic DEKA ERP mapping to accurately analyze vector CAD drawings and mixed-content datasheets offline.

**Architecture:** 3-tier hybrid pipeline: (1) Smart page selection & rendering using pypdfium2 for zero-text CAD PDFs; (2) Priority cascade: Exact PN Decoder (50ms fast-path) -> Text Regex -> Local Vision LLM fusion with target PN injection; (3) Deterministic DEKA ERP mapping via rules.xlsx with strict anti-hallucination guarantee.

**Tech Stack:** Python 3.10, pypdfium2, Pillow, httpx/requests (Ollama local REST API), FastAPI, pytest.

**Spec:** `docs/superpowers/specs/2026-10-19-vision-llm-datasheet-analyzer-design.md`

## Global Constraints

- 100% offline and local execution; zero external cloud API dependencies or keys.
- Safe for standard CPU servers: maximum 2 rendered pages per document, memory usage under 3.5GB RAM.
- ERP codes must remain 100% deterministic (generated strictly by rules.xlsx and map_specs_to_erp, never hallucinated by LLM).
- Preserves all 95+ existing pytest tests.
- Backward compatibility: If local Vision LLM is unavailable or offline, the system safely falls back without raising 500 errors.
- Safe failure: Never fallback unknown/0-text documents to "CAP" (capacitor); mark as "BİLİNMİYOR" with confidence 0.0.

## Review Focus

1. Zero-text CAD drawings (e.g. Degson `DG8-10-12-15HS-1.2.pdf`): Must not produce CAP; must produce CON with correct pin count (12 pin) and pitch (8.2mm) when PN is provided.
2. Multi-variant dimension tables (e.g. 2P to 24P in one CAD drawing): Target PN injection must force model to select 12P row rather than 2P or 24P.
3. Ollama server unavailable or timeout: System must smoothly fall back to deterministic extraction without throwing unhandled exceptions or crashing the web server.
4. Single-stage prompt with unselected material type: Model must return material_type + confidence + specs in a single unified JSON response.
5. Deterministic priority fusion: PN decoder values and high-confidence regex must override noisy visual tokens, while visual specs fill missing null fields.

---

### Task 1: Immediate Defense & Degson PN Decoder Quick-Fix

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py:640-660`, `src/bom_enricher/datasheet_analyzer.py:930-960`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `part_number` in `decode_manufacturer_pn` and `classify_material_type`.
- Produces: `decode_manufacturer_pn` returning connector specs for Degson `DG...HS` series; `classify_material_type` returning `"BİLİNMİYOR"`, 0.0 when text is empty and unclassified instead of `"CAP"`.

- [x] **Step 1: Write the failing test**

```python
# In tests/test_datasheet_analyzer.py
def test_degson_terminal_block_pn_decoding():
    from bom_enricher.datasheet_analyzer import decode_manufacturer_pn, classify_material_type
    pn = "DG8HS-1.2-12P-17-100SH/AH"
    specs = decode_manufacturer_pn(pn)
    assert specs.get("manufacturer") == "Degson"
    assert specs.get("pin_number") == "12 pin"
    assert specs.get("pitch") == "8.2mm"
    assert specs.get("pin_type") == "Dişi"
    assert specs.get("type_connector") == "Wire to Wire Klemens"
    assert specs.get("mounting") == "Cable"

    mat_type, conf, reason = classify_material_type("", pn, "")
    assert mat_type == "CON"


def test_empty_text_does_not_fallback_to_capacitor():
    from bom_enricher.datasheet_analyzer import classify_material_type
    mat_type, conf, reason = classify_material_type("", "UNKNOWN-PART-XYZ", "")
    assert mat_type == "BİLİNMİYOR"
    assert conf == 0.0
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_degson_terminal_block_pn_decoding tests/test_datasheet_analyzer.py::test_empty_text_does_not_fallback_to_capacitor -v`
Expected: FAIL (specs empty, mat_type == "CAP")

- [x] **Step 3: Implement Degson decoding and safe fallback in `datasheet_analyzer.py`**

1. In `classify_material_type`:
   - Add `re.match(r"^DG\d+.*", pn_clean)` check to classify as `CON`.
   - Update fallback line: when neither keyword nor rule matches, return `("BİLİNMİYOR", 0.0, "Malzeme tipi kesin belirlenemedi.")`.
2. In `decode_manufacturer_pn`:
   - Add regex pattern for Degson `DG...HS`:
     ```python
     degson_m = re.match(r"^DG(\d+)HS?-([0-9.]+)-(\d+)P", pn)
     if degson_m:
         series_num, pitch_code, pin_count = degson_m.groups()
         specs["manufacturer"] = "Degson"
         specs["type_connector"] = "Wire to Wire Klemens"
         specs["pin_type"] = "Dişi"
         specs["mounting"] = "Cable"
         specs["pin_number"] = f"{int(pin_count)} pin"
         # Pitch code mapping: 1.2 in DG8HS series represents 8.2mm pitch
         if pitch_code in ("1.2", "8.2"):
             specs["pitch"] = "8.2mm"
         elif pitch_code in ("1.0", "10.0"):
             specs["pitch"] = "10.0mm"
         elif pitch_code in ("1.5", "15.0"):
             specs["pitch"] = "15.0mm"
         return specs
     ```

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_degson_terminal_block_pn_decoding tests/test_datasheet_analyzer.py::test_empty_text_does_not_fallback_to_capacitor -v`
Expected: PASS

- [x] **Step 5: Run all existing tests to ensure no regressions**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -v`
Expected: PASS (all 95+ tests pass)

- [x] **Step 6: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "fix: add Degson PN decoder and replace blind CAP fallback with BILINMIYOR"
```

---

### Task 2: Smart PDF Page Selection & High-Speed Rendering (`pdf_renderer.py`)

**Files:**
- Create: `src/bom_enricher/pdf_renderer.py`
- Test: `tests/test_pdf_renderer.py`

**Interfaces:**
- Consumes: `pdf_bytes: bytes`, `target_pn: str | None`, `max_pages: int = 2`, `scale: float = 2.0`.
- Produces: `select_and_render_pages(pdf_bytes: bytes, target_pn: str | None = None, max_pages: int = 2, scale: float = 2.0) -> list[dict[str, Any]]`:
  Returns list of dicts with keys: `page_number` (int, 1-indexed), `base64_png` (str), `width` (int), `height` (int).

- [x] **Step 1: Write the failing test**

```python
# In tests/test_pdf_renderer.py
import pytest
from pathlib import Path
from bom_enricher.pdf_renderer import select_and_render_pages, select_target_pages

def test_select_target_pages_default_two_pages():
    # For a 5-page PDF with no match, returns [0, 1]
    pages = select_target_pages(total_pages=5, page_texts=["", "", "", "", ""], target_pn=None, max_pages=2)
    assert pages == [0, 1]

def test_select_target_pages_matches_target_pn():
    # Page 3 mentions target_pn
    texts = ["cover", "specs", "tables", "DG8HS-1.2-12P drawings", "compliance"]
    pages = select_target_pages(total_pages=5, page_texts=texts, target_pn="DG8HS-1.2-12P", max_pages=2)
    assert pages == [0, 3]

def test_render_pdf_page_to_base64_with_pypdfium2():
    # Generate a dummy 1-page PDF using reportlab or existing sample
    from reportlab.pdfgen import canvas
    import io
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "Test CAD Drawing Page")
    c.save()
    pdf_bytes = buf.getvalue()

    rendered = select_and_render_pages(pdf_bytes, target_pn="Test", max_pages=2)
    assert len(rendered) == 1
    assert rendered[0]["page_number"] == 1
    assert isinstance(rendered[0]["base64_png"], str)
    assert len(rendered[0]["base64_png"]) > 100
    assert rendered[0]["width"] > 0
    assert rendered[0]["height"] > 0
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_pdf_renderer.py -v`
Expected: FAIL with "No module named 'bom_enricher.pdf_renderer'"

- [x] **Step 3: Implement `src/bom_enricher/pdf_renderer.py`**

1. Define `select_target_pages(total_pages: int, page_texts: list[str], target_pn: str | None = None, max_pages: int = 2) -> list[int]`:
   - Page 0 is always included.
   - If `target_pn` is provided, find first page index `i > 0` where clean PN or series token appears in `page_texts[i]`.
   - If found, add `i`.
   - If not found and `total_pages > 1`, add page 1.
   - Cap result list at `max_pages`.
2. Define `select_and_render_pages(pdf_bytes: bytes, target_pn: str | None = None, max_pages: int = 2, scale: float = 2.0) -> list[dict[str, Any]]`:
   - Open with `pypdfium2.PdfDocument(pdf_bytes)`.
   - Extract lightweight text per page (for page selection).
   - Get page indices using `select_target_pages`.
   - For each selected page, call `page.render(scale=scale).to_pil()`.
   - Save to `io.BytesIO` as PNG and encode as base64 UTF-8 string.
   - Return metadata list.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_pdf_renderer.py -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/pdf_renderer.py tests/test_pdf_renderer.py
git commit -m "feat: add smart PDF page selection and pypdfium2 image renderer"
```

---

### Task 3: Local Vision LLM Client (`vision_client.py`)

**Files:**
- Create: `src/bom_enricher/vision_client.py`
- Test: `tests/test_vision_client.py`

**Interfaces:**
- Consumes: `images: list[str]` (base64 PNGs), `target_pn: str`, `material_type_hint: str | None`.
- Produces: `VisionLLMClient` with:
  - `is_available() -> bool`
  - `extract_specs_from_images(images: list[str], target_pn: str, material_type_hint: str | None = None) -> dict[str, Any] | None`

- [x] **Step 1: Write the failing test with mocked HTTP responses**

```python
# In tests/test_vision_client.py
import pytest
from unittest.mock import patch, MagicMock
from bom_enricher.vision_client import VisionLLMClient

def test_vision_client_is_available_true():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.get") as mock_get:
        mock_get.return_value = MagicMock(status_code=200)
        assert client.is_available() is True

def test_vision_client_is_available_false_on_connection_error():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.get", side_effect=Exception("Connection refused")):
        assert client.is_available() is False

def test_vision_client_extract_specs_success():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": '{"material_type": "CON", "confidence": 0.95, "manufacturer": "Degson", "specs": {"pitch": "8.2mm", "pin_number": "12 pin", "type_connector": "Wire to Wire Klemens", "mounting": "Cable", "pin_type": "Dişi"}}'
        }
    }
    with patch("httpx.post", return_value=mock_response):
        result = client.extract_specs_from_images(
            images=["base64_dummy_1"],
            target_pn="DG8HS-1.2-12P-17-100SH/AH"
        )
        assert result is not None
        assert result["material_type"] == "CON"
        assert result["specs"]["pitch"] == "8.2mm"
        assert result["specs"]["pin_number"] == "12 pin"

def test_vision_client_extract_specs_graceful_timeout():
    client = VisionLLMClient(base_url="http://localhost:11434", timeout=1.0)
    with patch("httpx.post", side_effect=Exception("Timeout")):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="TEST-PN")
        assert result is None  # Must not crash, return None
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_vision_client.py -v`
Expected: FAIL with "No module named 'bom_enricher.vision_client'"

- [x] **Step 3: Implement `src/bom_enricher/vision_client.py`**

1. Create `VisionLLMClient` with configurable `base_url`, `model` (default `"llama3.2-vision:3b"`), and `timeout` (default 15.0s).
2. `is_available()`: sends quick `httpx.get(f"{self.base_url}/api/tags", timeout=1.5)`. Returns True if 200, False otherwise.
3. Construct prompt including:
   - System instructions: Expert electronics component engineer.
   - Target Part Number injection: `Target Part Number: {target_pn}`.
   - Drawing instructions: notice about multi-variant tables (look up exact row/pin count matching target PN).
   - Unified JSON Schema: `material_type`, `confidence`, `manufacturer`, `specs` dictionary.
4. Send `POST /api/chat` with `{"model": self.model, "messages": [...], "format": "json", "stream": False}`.
5. Parse `response.json()["message"]["content"]` as JSON. Validate schema keys.
6. Return dictionary or `None` on any failure.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_vision_client.py -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/vision_client.py tests/test_vision_client.py
git commit -m "feat: add VisionLLMClient with target PN injection and timeout resilience"
```

---

### Task 4: Multimodal Hybrid Fusion in Datasheet Analyzer

**Files:**
- Modify: `src/bom_enricher/datasheet_analyzer.py`
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `pdf_renderer.select_and_render_pages`, `vision_client.VisionLLMClient`.
- Produces: `analyze_datasheet(stream_or_path, part_number=None, material_type_override=None, ...)` with multimodal fallback and fusion.

- [x] **Step 1: Write the failing test for hybrid fusion**

```python
# In tests/test_datasheet_analyzer.py
from unittest.mock import patch, MagicMock

def test_analyze_datasheet_cad_zero_text_uses_vision_fusion():
    from bom_enricher.datasheet_analyzer import analyze_datasheet
    # Simulate a zero-text CAD PDF with Degson drawing
    from reportlab.pdfgen import canvas
    import io
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    # Empty canvas simulating vector CAD drawing without extractable text
    c.save()
    cad_pdf_bytes = buf.getvalue()

    mock_vlm_output = {
        "material_type": "CON",
        "confidence": 0.90,
        "manufacturer": "Degson",
        "specs": {
            "type_connector": "Wire to Wire Klemens",
            "pitch": "8.2mm",
            "pin_number": "12 pin",
            "pin_type": "Dişi",
            "mounting": "Cable",
            "angle": "180°(Dik)"
        }
    }

    with patch("bom_enricher.vision_client.VisionLLMClient.is_available", return_value=True), \
         patch("bom_enricher.vision_client.VisionLLMClient.extract_specs_from_images", return_value=mock_vlm_output):
        result = analyze_datasheet(cad_pdf_bytes, part_number="DG8HS-1.2-12P-17-100SH/AH")
        assert result["material_type"] == "CON"
        assert result["specs"].get("pin_number") == "12 pin"
        assert result["specs"].get("pitch") == "8.2mm"
        # Deterministic ERP code must be generated matching CON rules
        assert result.get("erp_code") == "CONCAIB4120000XX-DE"
        assert result.get("analysis_method") in ("hybrid_vision", "fast_path")
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_analyze_datasheet_cad_zero_text_uses_vision_fusion -v`
Expected: FAIL (missing `erp_code` or `analysis_method` or falls back to CAP)

- [x] **Step 3: Implement priority fusion in `datasheet_analyzer.py`**

1. In `analyze_datasheet`:
   - Step A (Priority 1 - Fast-Path): Call `decode_manufacturer_pn(part_number)`.
   - Step B (Text Extraction): Call `extract_pdf_data`.
   - Step C: If text has content, run deterministic `extract_component_specs`.
   - Step D (Vision Fusion Trigger): Check condition:
     - Is text empty (`len(full_text.strip()) == 0`) OR are critical specs missing (e.g. `pitch` or `pin_number` missing for `CON`)?
     - If yes, instantiate `VisionLLMClient()`. If `client.is_available()`:
       - Render up to 2 pages via `select_and_render_pages(file_bytes, target_pn=part_number)`.
       - Call `client.extract_specs_from_images`.
       - If VLM returns specs:
         - Merge: for any field in `vlm_specs` that is missing or null in `extracted_specs`, fill it in from VLM!
         - If material type was unknown, adopt `vlm_material_type`.
         - Set `analysis_method = "hybrid_vision"`.
   - Step E: Pass merged specs into `map_specs_to_erp(rulebook, material_type, merged_specs)`.
   - Return enriched dictionary including `analysis_method`.

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_analyze_datasheet_cad_zero_text_uses_vision_fusion -v`
Expected: PASS

- [x] **Step 5: Run full test suite**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -v`
Expected: PASS

- [x] **Step 6: Commit**

```bash
git add src/bom_enricher/datasheet_analyzer.py tests/test_datasheet_analyzer.py
git commit -m "feat: implement multimodal hybrid fusion in datasheet analyzer"
```

---

### Task 5: Web API & Frontend Feedback

**Files:**
- Modify: `src/bom_enricher/web.py`
- Modify: `src/bom_enricher/static/index.html`
- Test: `tests/test_web.py`

**Interfaces:**
- Consumes: `/api/datasheet-analyze` multipart form data.
- Produces: API response returning `analysis_method` ("fast_path", "hybrid_vision", "deterministic_text") and UI status badge.

- [x] **Step 1: Write the failing test for Web API**

```python
# In tests/test_web.py
def test_datasheet_analyze_api_returns_analysis_method(client):
    from reportlab.pdfgen import canvas
    import io
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, "Test Resistor 10k 0805")
    c.save()
    pdf_bytes = buf.getvalue()

    response = client.post(
        "/api/datasheet-analyze",
        files={"file": ("test.pdf", pdf_bytes, "application/pdf")},
        data={"part_number": "RC0805FR-0710KL", "material_type": "RES"}
    )
    assert response.status_code == 200
    data = response.json()
    assert "analysis_method" in data
    assert data["material_type"] == "RES"
```

- [x] **Step 2: Run test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_web.py::test_datasheet_analyze_api_returns_analysis_method -v`
Expected: FAIL (missing `analysis_method`)

- [x] **Step 3: Update `web.py` and `index.html`**

1. In `src/bom_enricher/web.py`:
   - In `/api/datasheet-analyze` route: pass `part_number` and `material_type` to `analyze_datasheet`.
   - Ensure `analysis_method` is included in the JSON response payload.
2. In `src/bom_enricher/static/index.html`:
   - In the Datasheet Analysis card/tab:
     - Add an input field for "Hedef Parça Numarası (Target Part Number)" alongside the file upload input.
     - Display an analysis method badge next to the ERP code:
       - `fast_path`: `⚡ Hızlı Kod Çözücü` (green pill)
       - `hybrid_vision`: `👁️ Vision AI Destekli` (blue/purple pill)
       - `deterministic_text`: `📋 Metin Motoru` (slate pill)

- [x] **Step 4: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_web.py::test_datasheet_analyze_api_returns_analysis_method -v`
Expected: PASS

- [x] **Step 5: Commit**

```bash
git add src/bom_enricher/web.py src/bom_enricher/static/index.html tests/test_web.py
git commit -m "feat: expose analysis method badge and target PN input in UI and API"
```

---

### Task 6: Real CAD Sample Verification & End-to-End Regression Test

**Files:**
- Test: `tests/test_datasheet_analyzer.py`

**Interfaces:**
- Consumes: `C:\Users\sudem\Desktop\datasheets\DG8-10-12-15HS-1.2.pdf` or test fixture.
- Produces: Verified end-to-end extraction producing exact ERP code `CONCAIB4120000XX-DE`.

- [x] **Step 1: Write end-to-end integration test with the actual Degson PDF**

```python
# In tests/test_datasheet_analyzer.py
def test_real_degson_cad_datasheet_e2e():
    import os
    from pathlib import Path
    from bom_enricher.datasheet_analyzer import analyze_datasheet

    pdf_path = Path(r"C:\Users\sudem\Desktop\datasheets\DG8-10-12-15HS-1.2.pdf")
    if not pdf_path.exists():
        pytest.skip("Local Degson PDF not found on Desktop")

    with open(pdf_path, "rb") as f:
        file_bytes = f.read()

    pn = "DG8HS-1.2-12P-17-100SH/AH"
    result = analyze_datasheet(file_bytes, part_number=pn)

    # Must be classified as CON
    assert result["material_type"] == "CON"
    assert result["specs"].get("pin_number") == "12 pin"
    assert result["specs"].get("pitch") == "8.2mm"
    assert result["specs"].get("type_connector") == "Wire to Wire Klemens"
    assert result["specs"].get("pin_type") == "Dişi"
    assert result["specs"].get("mounting") == "Cable"

    # Deterministic ERP code verification
    assert result.get("erp_code") == "CONCAIB4120000XX-DE"
```

- [x] **Step 2: Run test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_datasheet_analyzer.py::test_real_degson_cad_datasheet_e2e -v`
Expected: PASS

- [x] **Step 3: Run the entire test suite**

Run: `.\.venv\Scripts\python.exe -m pytest tests/ -v`
Expected: PASS (all 97+ tests pass)

- [x] **Step 4: Commit**

```bash
git add tests/test_datasheet_analyzer.py
git commit -m "test: add real CAD datasheet end-to-end verification for Degson terminal block"
```
