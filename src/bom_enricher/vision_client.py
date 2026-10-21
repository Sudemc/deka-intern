"""
Local Vision LLM Client for Datasheet & Engineering Drawing Specification Extraction.
Interacts with local Ollama REST endpoints (e.g. llama3.2-vision:3b or qwen2-vl:2b)
with single-stage target PN injection, drawing variant guidance, and robust offline fallback.
"""
from __future__ import annotations

import json
import logging
import re
from typing import Any

import httpx

logger = logging.getLogger(__name__)


class VisionLLMClient:
    """Client for local Ollama Vision LLM inference with strict timeout resilience."""

    def __init__(
        self,
        base_url: str = "http://localhost:11434",
        model: str | None = None,
        timeout: float = 60.0,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self._configured_model = model
        self.model = model or "llama3.2-vision:3b"
        self.timeout = float(timeout)

    def _resolve_model(self) -> str:
        """Resolve available vision model from Ollama, auto-detecting if preferred isn't set."""
        if self._configured_model:
            return self._configured_model
        try:
            resp = httpx.get(f"{self.base_url}/api/tags", timeout=1.5)
            if resp.status_code == 200:
                models = resp.json().get("models", [])
                names = [m.get("name", "") for m in models]
                # Check for preferred lightweight vision models in order of CPU efficiency
                for cand in ("moondream", "llama3.2-vision", "llava", "qwen2-vl", "minicpm-v"):
                    for n in names:
                        if cand in n.lower():
                            return n
                if names:
                    return names[0]
        except Exception:
            pass
        return self.model

    def is_available(self) -> bool:
        """Quickly check if the local Ollama server is running and reachable."""
        try:
            resp = httpx.get(f"{self.base_url}/api/tags", timeout=max(0.1, min(1.5, self.timeout)))
            return resp.status_code == 200
        except Exception as e:
            logger.debug("Vision LLM service not available at %s: %s", self.base_url, e)
            return False

    @staticmethod
    def _try_repair_json(text: str) -> dict[str, Any] | None:
        """Attempt to repair truncated or unclosed JSON output from LLM."""
        try:
            res = json.loads(text)
            if isinstance(res, dict):
                return res
        except Exception:
            pass

        # Try backwards recovery at punctuation boundaries
        for i in range(len(text) - 1, 0, -1):
            if text[i] in (",", "{", "}", '"'):
                candidate = text[:i].rstrip(",")
                # Close unclosed quote if odd count
                if candidate.count('"') % 2 != 0:
                    candidate += '"'
                # Balance braces
                open_cur = candidate.count("{")
                close_cur = candidate.count("}")
                if open_cur > close_cur:
                    candidate += "}" * (open_cur - close_cur)
                try:
                    res = json.loads(candidate)
                    if isinstance(res, dict):
                        return res
                except Exception:
                    continue
        return None

    def extract_specs_from_images(
        self,
        images: list[str],
        target_pn: str,
        material_type_hint: str | None = None,
    ) -> dict[str, Any] | None:
        """Extract component specifications from base64 PNG images using Vision LLM.

        Args:
            images: List of base64-encoded PNG image strings.
            target_pn: Target manufacturer part number to locate in drawings/tables.
            material_type_hint: Optional material type hint (e.g. "CON", "RES", "CAP").

        Returns:
            Dictionary matching the unified schema or None if extraction fails or times out.
        """
        if not images:
            logger.debug("No images provided for Vision LLM extraction.")
            return None

        clean_pn = target_pn.strip() if target_pn else ""
        if not clean_pn:
            logger.debug("Empty target_pn provided for Vision LLM extraction.")
            return None

        system_prompt = (
            "You are an expert electronics component engineer analyzing engineering drawings, "
            "CAD diagrams, and specification sheets from component datasheets."
        )

        user_prompt_lines = [
            f"Target Part Number: {clean_pn}",
        ]
        if material_type_hint and material_type_hint.strip():
            user_prompt_lines.append(f"Material Type Hint: {material_type_hint.strip()}")

        user_prompt_lines.extend(
            [
                "",
                "IMPORTANT DRAWING INSTRUCTIONS:",
                "1. Datasheets and engineering drawings frequently contain multi-variant tables or dimension tables "
                "with columns and rows for various pin counts, pitches, or body sizes.",
                f"2. You MUST locate and extract specifications specifically for the Target Part Number '{clean_pn}' "
                "or its exact pin count / variant row. Do not mix up values from other variants or generic series entries.",
                "3. Extract all relevant electrical and mechanical specs visible in the drawing or tables "
                "(e.g., pitch, pin_number, type_connector, mounting, pin_type, capacitance, resistance, voltage, current).",
                "4. For card connectors (SIM, MicroSIM, SD, MicroSD), sum contact circuits and any detect/switch pins "
                "(e.g. 6 circuits + 2 detect pins = 8 pin), classify type_connector ('SIM Card', 'MicroSIM Card', 'MicroSD Card'), "
                "style ('Push-Push', 'Push-Pull', 'Hinged'), and orientation ('90°(Yatık)' for horizontal front/side entry).",
                "",
                "SCHEMA INSTRUCTIONS:",
                "Return ONLY a valid JSON object matching the following structure:",
                "{",
                '  "material_type": "<e.g. CON, RES, CAP, IND, DIODE, TRANS, IC, RELAY>",',
                '  "confidence": <float between 0.0 and 1.0>,',
                '  "manufacturer": "<manufacturer name or null>",',
                '  "specs": {',
                '    "<property_name>": "<property_value>"',
                "  }",
                "}",
                "Ensure 'specs' is an object mapping property names to string values. Do not wrap with explanation or markdown.",
            ]
        )
        user_prompt = "\n".join(user_prompt_lines)

        active_model = self._resolve_model()
        is_moondream = "moondream" in active_model.lower()

        messages: list[dict[str, Any]] = []
        if not is_moondream:
            messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": user_prompt, "images": images})
        else:
            # Moondream works best with direct user prompt without separate system role
            messages.append({"role": "user", "content": f"{system_prompt}\n\n{user_prompt}", "images": images})

        payload = {
            "model": active_model,
            "messages": messages,
            "format": "json",
            "options": {
                "temperature": 0.1,
                "num_predict": 256,
            },
            "stream": False,
        }

        try:
            resp = httpx.post(
                f"{self.base_url}/api/chat",
                json=payload,
                timeout=self.timeout,
            )
            if resp.status_code != 200:
                payload.pop("format", None)
                resp = httpx.post(
                    f"{self.base_url}/api/chat",
                    json=payload,
                    timeout=self.timeout,
                )
            resp.raise_for_status()
            data = resp.json()

            raw_content = data.get("message", {}).get("content", "")
            if not raw_content or not isinstance(raw_content, str):
                logger.warning("Empty content in Vision LLM response for %s", clean_pn)
                return None

            cleaned_content = raw_content.strip()
            # Strip markdown fences if present
            if cleaned_content.startswith("```"):
                cleaned_content = re.sub(r"^```(?:json)?\s*", "", cleaned_content, flags=re.IGNORECASE)
                cleaned_content = re.sub(r"\s*```$", "", cleaned_content)
                cleaned_content = cleaned_content.strip()

            start_idx = cleaned_content.find("{")
            end_idx = cleaned_content.rfind("}")
            if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
                json_candidate = cleaned_content[start_idx : end_idx + 1]
            else:
                json_candidate = cleaned_content

            parsed = None
            try:
                parsed = json.loads(json_candidate)
            except Exception:
                parsed = self._try_repair_json(json_candidate)

            # Heuristic text fallback if JSON parsing failed completely
            if not isinstance(parsed, dict) or ("material_type" not in parsed and "specs" not in parsed):
                extracted_specs: dict[str, Any] = {}
                mat_type = None

                # Detect material type from text
                m_type = re.search(r"(?:Type|Component|Material(?:\s*Type)?)\s*[:=]\s*([A-Za-z0-9_-]+)", cleaned_content, re.IGNORECASE)
                if m_type:
                    cand = m_type.group(1).upper()
                    if cand in ("CON", "RES", "CAP", "IND", "DIO", "TRA", "CRY", "DSP", "FUS", "REL", "SWI"):
                        mat_type = cand
                    elif "CONNECTOR" in cand or "HEADER" in cand or "TERMINAL" in cand:
                        mat_type = "CON"
                    elif "RESISTOR" in cand:
                        mat_type = "RES"
                    elif "CAPACITOR" in cand:
                        mat_type = "CAP"
                    elif "DIODE" in cand:
                        mat_type = "DIO"
                    elif "TRANSISTOR" in cand or "MOSFET" in cand:
                        mat_type = "TRA"
                    elif "CRYSTAL" in cand or "OSCILLATOR" in cand:
                        mat_type = "CRY"
                    elif "DISPLAY" in cand:
                        mat_type = "DSP"

                # Pitch
                pitch_m = re.search(r"\b(?:pitch|spacing)\b\s*[:=]?\s*([0-9.]+)\s*(?:mm)?", cleaned_content, re.IGNORECASE)
                if pitch_m:
                    extracted_specs["pitch"] = f"{pitch_m.group(1)}mm"

                # Pin count
                pin_m = re.search(
                    r"(?:(?:pin(?:s|\s*number|\s*count)?|positions?|poles?)\s*[:=]?\s*(\d+))|(?:\b(\d+)\s*(?:pin|pins|position|positions|pole|poles)\b)",
                    cleaned_content,
                    re.IGNORECASE,
                )
                if pin_m:
                    p_val = pin_m.group(1) or pin_m.group(2)
                    extracted_specs["pin_number"] = f"{p_val} pin"

                # Mounting
                if re.search(r"\b(?:through[- ]hole|th)\b", cleaned_content, re.IGNORECASE):
                    extracted_specs["mounting"] = "TH"
                elif re.search(r"\b(?:surface[- ]mount|smd|smt)\b", cleaned_content, re.IGNORECASE):
                    extracted_specs["mounting"] = "SMD"

                if mat_type or extracted_specs:
                    parsed = {
                        "material_type": mat_type or material_type_hint or "CON",
                        "confidence": 0.75,
                        "manufacturer": None,
                        "specs": extracted_specs,
                    }

            if not isinstance(parsed, dict):
                logger.warning("Vision LLM response could not be parsed into a dict for %s", clean_pn)
                return None

            # Validate schema
            if "material_type" not in parsed or not isinstance(parsed["material_type"], str):
                if material_type_hint:
                    parsed["material_type"] = material_type_hint
                else:
                    logger.warning("Missing or invalid 'material_type' in Vision LLM response for %s", clean_pn)
                    return None

            if "specs" not in parsed or not isinstance(parsed["specs"], dict):
                parsed["specs"] = {}

            # Normalize property keys to internal schema & filter out placeholder strings
            raw_specs = parsed.get("specs", {})
            normalized_specs: dict[str, Any] = {}
            key_aliases = {
                "vrrm": "voltage",
                "forward_voltage": "voltage",
                "reverse_voltage": "voltage",
                "breakdown_voltage": "voltage",
                "blocking_voltage": "voltage",
                "rated_voltage": "voltage",
                "if": "current",
                "if_av": "current",
                "forward_current": "current",
                "rated_current": "current",
                "average_current": "current",
                "freq": "frequency_str",
                "frequency": "frequency_str",
                "cap": "capacitance_str",
                "capacitance": "capacitance_str",
                "res": "resistance_str",
                "resistance": "resistance_str",
                "ind": "inductance_str",
                "inductance": "inductance_str",
                "pkg": "package",
                "case": "package",
            }

            def _is_valid(val: Any) -> bool:
                if val is None:
                    return False
                s = str(val).strip()
                if not s or s in ("...", "N/A", "null", "None", "<...>", "—"):
                    return False
                if s.startswith("<") and s.endswith(">"):
                    return False
                return True

            for k, v in raw_specs.items():
                if not _is_valid(v):
                    continue
                normalized_specs[k] = v
                norm_k = key_aliases.get(k.lower().strip())
                if norm_k and norm_k != k:
                    normalized_specs[norm_k] = v
            parsed["specs"] = normalized_specs

            # Normalize confidence if present or set default
            if "confidence" in parsed and isinstance(parsed["confidence"], (int, float)):
                parsed["confidence"] = float(parsed["confidence"])
            else:
                parsed["confidence"] = 0.8

            mfr_val = parsed.get("manufacturer")
            if not _is_valid(mfr_val):
                parsed["manufacturer"] = None
            else:
                parsed["manufacturer"] = str(mfr_val).strip()

            return parsed

        except Exception as e:
            logger.warning("Vision LLM extraction failed for %s: %s", clean_pn, e)
            return None
