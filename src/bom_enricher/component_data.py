from __future__ import annotations

import json
import os
import time
from abc import ABC, abstractmethod
from itertools import islice
from pathlib import Path

import requests

from .models import ComponentData, ComponentLookupResult


def normalize_part_number(part_number: str) -> str:
    """Normalize only case and surrounding/internal whitespace for matching."""
    return " ".join(str(part_number).split()).casefold()


class ComponentDataProvider(ABC):
    @abstractmethod
    def lookup_many(
        self, part_numbers: list[str]
    ) -> dict[str, ComponentLookupResult]:
        """Return one explicit lookup outcome for every requested part number."""


class JsonCatalogProvider(ComponentDataProvider):
    """Read a fictional, offline component catalog used by samples and tests."""

    def __init__(self, catalog_path: str | Path) -> None:
        self._catalog = self._load_catalog(Path(catalog_path))

    def lookup_many(
        self, part_numbers: list[str]
    ) -> dict[str, ComponentLookupResult]:
        results: dict[str, ComponentLookupResult] = {}
        for requested_part_number in part_numbers:
            component = self._catalog.get(normalize_part_number(requested_part_number))
            if component is None:
                results[requested_part_number] = ComponentLookupResult(
                    status="not_found",
                    component=None,
                    message="Part number was not found in the offline fictional catalog.",
                )
            else:
                results[requested_part_number] = ComponentLookupResult(
                    status="matched", component=component
                )
        return results

    @staticmethod
    def _load_catalog(catalog_path: Path) -> dict[str, ComponentData]:
        payload = json.loads(catalog_path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("catalog must contain a JSON object")

        catalog: dict[str, ComponentData] = {}
        for part_number, record in payload.items():
            if part_number == "synthetic_data_disclaimer":
                continue
            if not isinstance(record, dict):
                raise ValueError(f"catalog entry {part_number!r} must be an object")

            normalized_part_number = normalize_part_number(part_number)
            if normalized_part_number in catalog:
                raise ValueError(
                    f"duplicate normalized part number: {part_number!r}"
                )

            raw_attributes = record.get("attributes", {})
            if not isinstance(raw_attributes, dict):
                raise ValueError(f"attributes for {part_number!r} must be an object")
            attributes = {str(key): str(value) for key, value in raw_attributes.items()}
            if "mounting" in record:
                attributes["mounting"] = str(record["mounting"])

            catalog[normalized_part_number] = ComponentData(
                part_number=str(part_number).strip(),
                manufacturer=str(record.get("manufacturer", "")),
                short_description=str(record.get("short_description", "")),
                attributes=attributes,
                datasheet_url=str(record.get("datasheet_url", "")),
            )
        return catalog


class NexarOctopartProvider(ComponentDataProvider):
    """Look up exact manufacturer part numbers through Nexar's GraphQL API."""

    _TOKEN_URL = "https://identity.nexar.com/connect/token"
    _GRAPHQL_URL = "https://api.nexar.com/graphql"
    _TIMEOUT_SECONDS = 20
    _TOKEN_SAFETY_SKEW_SECONDS = 60
    _MAX_PARTS_PER_QUERY = 100
    _QUERY = """
    query SupMultiMatch($queries: [SupPartMatchQuery!]!) {
      supMultiMatch(queries: $queries) {
        reference
        hits
        parts {
          mpn
          manufacturer { name }
          shortDescription
          octopartUrl
          specs {
            attribute { name shortname }
            displayValue
          }
        }
      }
    }
    """

    def __init__(
        self,
        client_id: str | None = None,
        client_secret: str | None = None,
        session: requests.Session | None = None,
    ) -> None:
        self._client_id = (client_id if client_id is not None else os.getenv("NEXAR_CLIENT_ID", "")).strip()
        self._client_secret = (
            client_secret if client_secret is not None else os.getenv("NEXAR_CLIENT_SECRET", "")
        ).strip()
        if not self._client_id or not self._client_secret:
            raise ValueError(
                "NEXAR_CLIENT_ID and NEXAR_CLIENT_SECRET must be set to use --provider nexar."
            )
        self._session = session or requests.Session()
        self._access_token: str | None = None
        self._token_expires_at = 0.0

    def lookup_many(
        self, part_numbers: list[str]
    ) -> dict[str, ComponentLookupResult]:
        requested = list(dict.fromkeys(part_numbers))
        results: dict[str, ComponentLookupResult] = {
            part_number: ComponentLookupResult(
                status="not_found",
                component=None,
                message="Part number is blank.",
            )
            for part_number in requested
            if not str(part_number).strip()
        }
        query_part_numbers = [
            part_number for part_number in requested if str(part_number).strip()
        ]
        if not query_part_numbers:
            return results

        try:
            token = self._get_access_token()
        except (requests.RequestException, ValueError, KeyError, TypeError) as error:
            results.update(
                self._unverified_results(query_part_numbers, self._failure_message(error))
            )
            return results

        part_iterator = iter(query_part_numbers)
        while batch := list(islice(part_iterator, self._MAX_PARTS_PER_QUERY)):
            try:
                payload = self._query_parts(batch, token)
                results.update(self._parse_results(batch, payload))
            except (requests.RequestException, ValueError, KeyError, TypeError) as error:
                results.update(
                    self._unverified_results(batch, self._failure_message(error))
                )
        return results

    def _failure_message(self, error: Exception) -> str:
        response = getattr(error, "response", None)
        status_code = getattr(response, "status_code", None)
        if status_code == 401:
            message = (
                "Nexar kimlik doğrulaması başarısız (HTTP 401). "
                "Client ID ve Client Secret değerlerini kontrol edin."
            )
        elif status_code == 403:
            message = (
                "Nexar erişimi reddedildi (HTTP 403). Uygulamanın Supply kapsamını "
                "ve etkin planını kontrol edin."
            )
        elif status_code == 429:
            message = "Nexar sorgu sınırına ulaşıldı (HTTP 429). Daha sonra tekrar deneyin."
        elif status_code is not None:
            message = f"Nexar HTTP isteği başarısız oldu (HTTP {status_code})."
        elif isinstance(error, requests.Timeout):
            message = "Nexar isteği zaman aşımına uğradı. Bağlantıyı kontrol edip tekrar deneyin."
        elif isinstance(error, requests.ConnectionError):
            message = "Nexar sunucusuna bağlanılamadı. İnternet/proxy bağlantısını kontrol edin."
        elif isinstance(error, requests.RequestException):
            message = "Nexar HTTP isteği tamamlanamadı."
        else:
            detail = " ".join(str(error).split())[:240]
            message = f"Nexar yanıtı işlenemedi: {detail or type(error).__name__}"

        # Error text from a remote service must never echo configured credentials.
        for credential in (self._client_id, self._client_secret):
            if credential:
                message = message.replace(credential, "[redacted]")
        return message

    def _get_access_token(self) -> str:
        if self._access_token and time.monotonic() < self._token_expires_at:
            return self._access_token

        response = self._session.post(
            self._TOKEN_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": self._client_id,
                "client_secret": self._client_secret,
                "scope": "supply.domain",
            },
            timeout=self._TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict):
            raise ValueError("Nexar token response is malformed.")
        token = payload.get("access_token")
        expires_in = payload.get("expires_in")
        if not isinstance(token, str) or not token.strip():
            raise ValueError("Nexar token response has no access token.")
        if not isinstance(expires_in, (int, float)):
            raise ValueError("Nexar token response has no valid expiration.")

        self._access_token = token
        self._token_expires_at = time.monotonic() + max(
            0, float(expires_in) - self._TOKEN_SAFETY_SKEW_SECONDS
        )
        return token

    def _query_parts(self, part_numbers: list[str], token: str) -> dict[str, object]:
        response = self._session.post(
            self._GRAPHQL_URL,
            json={
                "query": self._QUERY,
                "variables": {
                    "queries": [
                        {"mpn": part_number, "reference": part_number}
                        for part_number in part_numbers
                    ]
                },
            },
            headers={"Authorization": f"Bearer {token}"},
            timeout=self._TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict):
            raise ValueError("Nexar GraphQL response is not a JSON object.")
        errors = payload.get("errors")
        if errors:
            first_error = errors[0] if isinstance(errors, list) and errors else None
            detail = first_error.get("message") if isinstance(first_error, dict) else None
            if not isinstance(detail, str) or not detail.strip():
                detail = "Nexar GraphQL query was rejected."
            raise ValueError(f"Nexar GraphQL hatası: {' '.join(detail.split())[:240]}")
        return payload

    def _parse_results(
        self, requested_part_numbers: list[str], payload: dict[str, object]
    ) -> dict[str, ComponentLookupResult]:
        data = payload.get("data")
        if not isinstance(data, dict):
            raise ValueError("Nexar GraphQL response has no data object.")
        matches = data.get("supMultiMatch")
        if not isinstance(matches, list):
            raise ValueError("Nexar GraphQL response has no supMultiMatch list.")

        requested_by_reference = {
            part_number: part_number for part_number in requested_part_numbers
        }
        matches_by_reference: dict[str, object] = {}
        for match in matches:
            if not isinstance(match, dict):
                raise ValueError("Nexar GraphQL response contains an invalid match.")
            reference = match.get("reference")
            if not isinstance(reference, str) or reference not in requested_by_reference:
                raise ValueError("Nexar GraphQL response has an unclear part association.")
            if reference in matches_by_reference:
                raise ValueError("Nexar GraphQL response has duplicate part associations.")
            matches_by_reference[reference] = match

        if set(matches_by_reference) != set(requested_part_numbers):
            raise ValueError("Nexar GraphQL response is missing part associations.")

        results: dict[str, ComponentLookupResult] = {}
        for part_number in requested_part_numbers:
            match = matches_by_reference.get(part_number)
            if match is None:
                results[part_number] = ComponentLookupResult(
                    status="not_found",
                    component=None,
                    message="Nexar returned no match for this part number.",
                )
                continue
            results[part_number] = self._parse_match(part_number, match)
        return results

    def _parse_match(
        self, requested_part_number: str, match: object
    ) -> ComponentLookupResult:
        if not isinstance(match, dict):
            raise ValueError("Nexar GraphQL response contains an invalid match.")
        parts = match.get("parts")
        if not isinstance(parts, list):
            raise ValueError("Nexar GraphQL response contains invalid parts.")
        if not parts:
            return ComponentLookupResult(
                status="not_found",
                component=None,
                message="Nexar returned no matching part.",
            )

        exact_parts = [
            part
            for part in parts
            if isinstance(part, dict)
            and isinstance(part.get("mpn"), str)
            and normalize_part_number(part["mpn"])
            == normalize_part_number(requested_part_number)
        ]
        if len(exact_parts) > 1:
            return ComponentLookupResult(
                status="ambiguous",
                component=None,
                message="Nexar returned more than one exact part-number match.",
            )
        if not exact_parts:
            return ComponentLookupResult(
                status="unverified",
                component=None,
                message="Nexar returned a part, but no exact normalized MPN match.",
            )

        part = exact_parts[0]
        manufacturer = part.get("manufacturer")
        manufacturer_name = ""
        if isinstance(manufacturer, dict) and isinstance(manufacturer.get("name"), str):
            manufacturer_name = manufacturer["name"].strip()
        short_description = part.get("shortDescription")
        if not isinstance(short_description, str):
            short_description = ""
        return ComponentLookupResult(
            status="matched",
            component=ComponentData(
                part_number=str(part["mpn"]).strip(),
                manufacturer=manufacturer_name,
                short_description=short_description.strip(),
                attributes=self._parse_specs(part.get("specs")),
                # octopartUrl is a product page, not a datasheet link.
                datasheet_url="",
            ),
        )

    @classmethod
    def _parse_specs(cls, specs: object) -> dict[str, str]:
        if not isinstance(specs, list):
            return {}
        attributes: dict[str, str] = {}
        for spec in specs:
            if not isinstance(spec, dict):
                continue
            attribute = spec.get("attribute")
            display_value = spec.get("displayValue")
            if (
                not isinstance(attribute, dict)
                or not isinstance(attribute.get("name"), str)
                or not isinstance(display_value, str)
                or not display_value.strip()
            ):
                continue
            mapped_name = cls._map_spec_name(attribute["name"])
            if not mapped_name or mapped_name in attributes:
                continue
            if mapped_name == "mounting":
                mounting = cls._map_mounting(display_value)
                if mounting:
                    attributes[mapped_name] = mounting
            else:
                attributes[mapped_name] = display_value.strip()
        return attributes

    @staticmethod
    def _map_spec_name(attribute_name: str) -> str | None:
        normalized = " ".join(attribute_name.casefold().split())
        names = {
            "capacitance": "capacitance",
            "voltage": "voltage",
            "voltage - rated": "voltage",
            "dielectric": "dielectric",
            "resistance": "resistance",
            "power (watts)": "power",
            "power": "power",
            "tolerance": "tolerance",
            "package / case": "case_package",
            "mounting type": "mounting",
        }
        return names.get(normalized)

    @staticmethod
    def _map_mounting(display_value: str) -> str | None:
        normalized = " ".join(display_value.casefold().replace("-", " ").split())
        if "through hole" in normalized or normalized == "th":
            return "TH"
        if "surface mount" in normalized or normalized == "smd":
            return "SMD"
        return None

    @staticmethod
    def _unverified_results(
        part_numbers: list[str],
        message: str = "Nexar lookup failed; verify this part manually.",
    ) -> dict[str, ComponentLookupResult]:
        return {
            part_number: ComponentLookupResult(
                status="unverified",
                component=None,
                message=message,
            )
            for part_number in part_numbers
        }
