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
- **Performance & Timeouts:** 
  - Fast Text LLM (`llama3.2:3b`): 5.0 seconds maximum timeout.
  - Multimodal Vision LLM (`moondream`): 30.0 seconds timeout (page rendering + image processing on CPU).
  - If a timeout occurs, fall back gracefully to deterministic specs with an arbitration log entry recording `llm_status = "timeout"`.

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
- `llm_status`: 
  - `"active"`: LLM executed and contributed/arbitrated specs.
  - `"standby"`: Ollama is online and healthy, but component was resolved deterministically with high confidence (>= 0.95), so LLM was not needed.
  - `"offline"`: Ollama server is not running on localhost.
  - `"timeout"`: LLM execution timed out (gracefully fallen back to regex).
  - `"error"`: LLM call or JSON parsing encountered an issue.
- `llm_trace`: list of `ArbitrationEvent` dicts
- `llm_specs`: dict of raw LLM extractions
- `regex_specs`: dict of raw regex extractions

### B. Trigger & Arbitration Logic in `datasheet_analyzer.py`
1. **Auditor Trigger Conditions:**
   The text LLM acts as an active **Quality Auditor & Semantic Validator**. It engages whenever:
   - **Category-Mandatory Audit:** Any component belonging to complex, ambiguous, or multi-manufacturer categories (`CON`, `SWT`, `REL`, `ICR`, `TRA`, `IND`, `DSP`, `DIO`, `FUS`, etc.) automatically triggers the LLM auditor when Ollama is available.
   - **Semantic Contradiction / Red-Flag Heuristic:** Regardless of confidence score, if any textual keyword contradicts deterministic extraction (e.g. text has "socket", "receptacle", "female" but regex assigned "Erkek"; or text has "through hole", "dip", "tht", "leaded" but regex defaulted to "SMD"), trigger LLM immediately.
   - **Ambiguity & Missing Fields:** `conf < 0.98`, missing core attributes, catalog documents, or unresolved part numbers.
   - **Fast-Path Exemption (Pure Speed):** ONLY standard chip passive components (`RES`, `CAP`) whose part number is 100% mathematically decoded by algorithmic rules with confidence >= 0.98 bypass the LLM.

   ```python
   # Exemption check: purely standard chip passives resolved by exact mathematical PN decoders
   is_pure_passive = mat_type in ("RES", "CAP") and pn_has_core and conf >= 0.98 and not is_catalog
   
   # Red-flag contradiction check
   has_gender_conflict = (
       specs.get("pin_type") == "Erkek" and any(w in clean_text.casefold() for w in ("socket", "receptacle", "female"))
   ) or (
       specs.get("pin_type") == "Dişi" and any(w in clean_text.casefold() for w in ("male header", "pin header", "plug"))
   )
   has_mounting_conflict = (
       specs.get("mounting") == "SMD" and any(w in clean_text.casefold() for w in ("through hole", "tht", " dip ", "leaded"))
   ) or (
       specs.get("mounting") == "TH" and any(w in clean_text.casefold() for w in ("surface mount", "smd", "smt"))
   )
   has_red_flag = has_gender_conflict or has_mounting_conflict

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

2. **Arbitration Rules:**
   - **`pin_type` (Connector Gender):**
     If LLM identifies `Dişi` or `Erkek` with high contextual cues (e.g. text contains "socket", "receptacle", "female"), and regex extracted `Erkek` or `Dişi`, the LLM overrides regex, logging: `"LLM bağlamsal analiz ile pin_type değerini güncelledi"`.
   - **`mounting` (Montaj Tipi):**
     If LLM extracts explicit mounting (`"TH"` or `"SMD"`) with textual evidence and conflicts with regex, LLM arbitrates and logs the rationale.
   - **`type_connector` & `style` & `pin_number`:**
     Specialized connectors (e.g., SIM/MicroSD Push-Push, Card Socket, Pin Header, Terminal Block) and compound pin counts detected by LLM take precedence over generic regex guesses.
   - **Non-Connector Categories (`SWT`, `REL`, `ICR`, `DIO`, etc.):**
     LLM fills empty/unextracted fields (`source: "llm_filled_empty"`) and discovers part numbers and manufacturers. Conflicting non-connector values retain regex (`source: "regex_retained"`) to guarantee zero regression risk for existing benchmark suites.
   - **Empty / None Fields:**
     LLM fills any missing fields that regex failed to extract across all categories.

### C. Client Status Contract in `VisionLLMClient`
`VisionLLMClient` exposes `last_status` (`"success"`, `"timeout"`, `"connection_error"`, `"parse_error"`). This allows `datasheet_analyzer.py` to determine exact `llm_status` cleanly without breaking existing method signatures or test contracts.

### D. Backend API Update in `web.py`
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

### E. Frontend Collapsible UI in `static/index.html`
In `static/index.html` within the component card rendering template:
```html
<details class="llm-trace-accordion">
  <summary class="llm-trace-summary">
    <span>🔍 Yapay Zeka (LLM) & Karar Geçmişi</span>
    <span class="llm-trace-pill ${item.llm_status === 'active' ? 'pill-green' : (item.llm_status === 'standby' ? 'pill-blue' : 'pill-yellow')}">
      ${item.llm_status === 'active' ? '🟢 LLM Aktif' : (item.llm_status === 'standby' ? '🔵 Standby (Kural Motoru)' : '🟡 Ollama Kapalı')}
    </span>
  </summary>
  <div class="llm-trace-content">
    <!-- Status & Info Banner -->
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
- Unit test coverage for `ArbitrationEvent`, status lifecycles, and arbitration overrides.

