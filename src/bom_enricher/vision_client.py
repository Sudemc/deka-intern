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
        model: str = "llama3.2-vision:3b",
        timeout: float = 15.0,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = float(timeout)

    def is_available(self) -> bool:
        """Quickly check if the local Ollama server is running and reachable."""
        try:
            resp = httpx.get(f"{self.base_url}/api/tags", timeout=max(0.1, min(1.5, self.timeout)))
            return resp.status_code == 200
        except Exception as e:
            logger.debug("Vision LLM service not available at %s: %s", self.base_url, e)
            return False

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

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                    "images": images,
                },
            ],
            "format": "json",
            "stream": False,
        }

        try:
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
                cleaned_content = cleaned_content[start_idx : end_idx + 1]

            parsed = json.loads(cleaned_content)
            if not isinstance(parsed, dict):
                logger.warning("Vision LLM response is not a JSON object for %s", clean_pn)
                return None

            # Validate schema
            if "material_type" not in parsed or not isinstance(parsed["material_type"], str):
                logger.warning("Missing or invalid 'material_type' in Vision LLM response for %s", clean_pn)
                return None

            if "specs" not in parsed or not isinstance(parsed["specs"], dict):
                logger.warning("Missing or invalid 'specs' dict in Vision LLM response for %s", clean_pn)
                return None

            # Normalize confidence if present or set default
            if "confidence" in parsed and isinstance(parsed["confidence"], (int, float)):
                parsed["confidence"] = float(parsed["confidence"])
            else:
                parsed["confidence"] = 0.8

            if "manufacturer" not in parsed or not isinstance(parsed["manufacturer"], (str, type(None))):
                parsed["manufacturer"] = None

            return parsed

        except Exception as e:
            logger.warning("Vision LLM extraction failed for %s: %s", clean_pn, e)
            return None
