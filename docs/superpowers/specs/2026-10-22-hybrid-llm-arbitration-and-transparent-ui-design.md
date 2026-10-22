# Design Specification: Hybrid LLM Arbitration, Trigger Redesign & Transparent Collapsible UI

- **Date:** 2026-10-22
- **Topic:** Hybrid LLM Trigger Conditions, Semantic Arbitration Authority, and Collapsible UI Inspection
- **Branch:** `dev-sudem`
- **Target File:** `docs/superpowers/specs/2026-10-22-hybrid-llm-arbitration-and-transparent-ui-design.md`

---

## 1. Intent and The Why

The BOM Enricher system incorporates a local multimodal Vision/Text LLM (via local Ollama with `llama3.2:3b` and `moondream`) designed to work in synergy with deterministic regex rule engines. In practice, however, two major systemic issues hindered its effectiveness:
1. **Silent LLM Bypass & Zero UI Transparency:**
   When the local Ollama background service is not running or encounters an exception, the system silently drops back to regex heuristics without notifying the user. The user observes incorrect extractions and assumes the LLM made an error, when in fact the LLM was never executed.
2. **The "Empty-Only Fill" Asymmetry:**
   In `src/bom_enricher/datasheet_analyzer.py` (lines 5271-5277), LLM output was structurally restricted to filling empty/None fields (`if k not in specs or specs[k] is None`). If an aggressive regex rule made an erroneous guess (e.g. categorizing a "WR-PHD SMT Socket Header" as `pin_type: "Erkek"`), the LLM was strictly forbidden from overriding or correcting that mistake.
3. **Trigger Incompleteness:**
   LLM execution was guarded by `if file_bytes and vlm.is_available():`. If an API caller or test passed parsed text without raw `file_bytes`, text LLM analysis was completely skipped even when the text was ambiguous.

### User Preference for UI
The user explicitly requested:
> *"arayüzde direkt çıkmasın tıklayınca açılan şekilde görebileyim"*
The analysis results should remain clean and uncluttered by default. The full LLM execution status, arbitration decisions, and parameter comparisons should be presented in a collapsible/expandable section that opens only when clicked.

---

## 2. Goals, Non-Goals, and Anti-Goals

### Goals
- **Active LLM Arbitration Authority:**
  Allow the LLM to arbitrate and override weak or conflicting regex extractions for ambiguous/complex component categories (`CON`, `SWT`, `REL`, `ICR`, etc.) on critical attributes:
  - `pin_type` (e.g. Dişi vs Erkek for Socket Headers / Receptacles)
  - `mounting` (e.g. SMD vs TH)
  - `type_connector` (e.g. specialized card/pin headers)
  - `style` / `angle` (e.g. Push-Push, Right Angle)
- **Arbitration Trace & Decision Logging:**
  Record structured decision metadata (`llm_status`, `analysis_method`, `arbitration_log`, `regex_raw_specs`, `llm_raw_specs`) in the analysis result.
- **Collapsible UI Inspection Component:**
  In `src/bom_enricher/static/index.html`, add a `<details class="llm-trace-accordion">` element to each component card. Closed by default, clicking it displays:
  - Ollama service status (Active / Offline)
  - Method used (Hybrid LLM, Vision AI, Fast-Path, Deterministic Text)
  - Step-by-step arbitration table (Attribute, Regex Value, LLM Value, Final Chosen Value, Decision Reason)
  - LLM raw extracted parameters vs. regex parameters
- **Health Check & Offline Notification:**
  Add a lightweight `/api/health/llm` endpoint. If Ollama is offline during analysis, the collapsible section explicitly explains: *"Yerel Ollama servisi (http://localhost:11434) kapalı. Analiz deterministik kural motoru ile tamamlandı."*
- **Text-Only LLM Support:**
  Ensure `vlm.extract_specs_from_text` executes when `clean_text` is available, regardless of whether `file_bytes` is provided.

### Non-Goals
- Cloud LLM / External API calls (must remain 100% offline, local Ollama only).
- Unconstrained ERP code generation by LLM (ERP codes must remain 100% deterministic, grounded in `rules.xlsx` and `map_specs_to_erp`).
- Replacing deterministic rules for high-confidence, standard passive components (e.g. standard chip resistors/capacitors where regex has >= 0.98 confidence).

### Anti-Goals
- Cluttering the main result card with open JSON dumps or intrusive diagnostic boxes.
- Breaking any existing benchmark or unit test (all 182 existing tests across 16 test files must stay green).
- Crashing or stalling the web UI or CLI when Ollama is unavailable or times out.

---

## 3. Constraints

- **Platform & Environment:** Windows 11, local Python virtualenv (`.\\.venv\\Scripts\\python.exe`), local Ollama (`http://127.0.0.1:11434`).
- **Dependencies:** FastAPI, Uvicorn, Jinja2/Static HTML, PyMuPDF, pdfplumber, httpx. No new external package dependencies.
- **Performance:** LLM timeout set to 5.0 seconds maximum to prevent hanging the web interface during batch processing. If timed out, fall back gracefully to deterministic specs with an arbitration log entry.

---

## 4. Technical Architecture & Decision Points

### A. Decision Trace Structure (`ArbitrationTrace`)
In `datasheet_analyzer.py`, introduce a structured dataclass or dict for decision logging:
```python
@dataclass
class ArbitrationEvent:
    attribute: str
    regex_value: Any
    llm_value: Any
    final_value: Any
    decision_reason: str
    source: str  # "llm_override", "regex_retained", "llm_filled_empty", "regex_default"
```
The analysis return dictionary and `DatasheetAnalysisResult` will include:
- `llm_status`: `"active"` | `"offline"` | `"timeout"` | `"error"`
- `llm_trace`: list of `ArbitrationEvent` dicts
- `llm_specs`: dict of raw LLM extractions
- `regex_specs`: dict of raw regex extractions

### B. Trigger & Arbitration Logic in `datasheet_analyzer.py`
1. **Trigger Condition:**
   ```python
   should_trigger_text_llm = (
       vlm.is_available()
       and bool(clean_text)
       and (
           is_complex_or_ambiguous
           or are_critical_missing
           or is_catalog
           or not pn
           or pn.startswith("media_")
           or conf < 0.95
           or any(k in clean_text.casefold() for k in ("detect pin", "detection switch", "sim card", "microsd", "socket header", "receptacle"))
       )
   )
   ```
2. **Arbitration Rules:**
   - **`pin_type` (Connector Gender):**
     If LLM identifies `Dişi` or `Erkek` with high contextual cues (e.g. text contains "socket", "receptacle", "female"), and regex extracted `Erkek` or `Dişi`, the LLM can override regex, logging: `"LLM bağlamsal analiz ile pin_type değerini güncelledi"`.
   - **`mounting` (Montaj Tipi):**
     If regex is uncertain or defaulted to SMD, and LLM extracted `TH` or `SMD` from package descriptions, LLM overrides and logs the decision.
   - **`type_connector` & `style`:**
     Specialized connectors (e.g., SIM/MicroSD Push-Push, Card Socket, Pin Header, Terminal Block) detected by LLM take precedence over generic "Header" regex guesses.
   - **Empty / None Fields:**
     LLM fills any missing fields that regex failed to extract.

### C. Backend API Update in `web.py`
1. Endpoint `GET /api/health/llm`:
   Calls `vlm.is_available()` and checks models (`llama3.2:3b`, `moondream`), returning:
   ```json
   {
     "status": "online",
     "text_model": "llama3.2:3b",
     "vision_model": "moondream:latest",
     "host": "http://127.0.0.1:11434"
   }
   ```
2. In `analyze_single_datasheet`, propagate `llm_trace`, `llm_status`, `llm_specs`, and `regex_specs` into the JSON payload returned to the frontend.

### D. Frontend Collapsible UI in `static/index.html`
In `static/index.html` within the component card rendering template:
```html
<details class="llm-trace-accordion">
  <summary class="llm-trace-summary">
    <span>🔍 Yapay Zeka (LLM) & Karar Geçmişi</span>
    <span class="llm-trace-pill ${item.llm_status === 'active' ? 'pill-green' : 'pill-yellow'}">
      ${item.llm_status === 'active' ? '🟢 LLM Aktif' : '🟡 Kural Motoru'}
    </span>
  </summary>
  <div class="llm-trace-content">
    <!-- Status & Info -->
    <div class="llm-status-banner">...</div>
    <!-- Arbitration Decisions Table -->
    ${arbitrationTableHtml}
    <!-- Raw Specs Comparison -->
    <div class="llm-specs-comparison">...</div>
  </div>
</details>
```
Styling:
- Subtle border, rounded corners, clean hover effect on summary.
- Arrow icon that rotates on open.
- Responsive, clean table styling matching the existing DEKA design language.
- Closed by default so the interface remains clean and fast.

---

## 5. What's Left to the Builder

- Exact CSS styles and transition animations for `<details>` in `index.html`.
- Fine-tuning the specific text prompts in `VisionLLMClient` if additional attribute extraction fields are desired.
- Unit test coverage for `ArbitrationEvent` and arbitration overrides.
