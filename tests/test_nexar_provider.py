from __future__ import annotations

from dataclasses import dataclass

import pytest
import requests

from bom_enricher.component_data import NexarOctopartProvider


@dataclass
class FakeResponse:
    payload: object
    status_code: int = 200

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            error = requests.HTTPError(f"HTTP {self.status_code}")
            error.response = self
            raise error

    def json(self):
        if isinstance(self.payload, Exception):
            raise self.payload
        return self.payload


class FakeSession:
    def __init__(self, post_results: list[object]) -> None:
        self.post_results = list(post_results)
        self.calls: list[dict[str, object]] = []

    def post(self, url: str, **kwargs):
        self.calls.append({"url": url, **kwargs})
        result = self.post_results.pop(0)
        if isinstance(result, Exception):
            raise result
        return result


def _graphql_payload(part_number: str = "ABC-123") -> dict[str, object]:
    return {
        "data": {
            "supMultiMatch": [
                {
                    "reference": part_number,
                    "hits": 1,
                    "parts": [
                        {
                            "mpn": part_number,
                            "manufacturer": {"name": "Example Manufacturer"},
                            "shortDescription": "Example 10 uF capacitor",
                            "octopartUrl": "https://octopart.com/example",
                            "specs": [
                                {
                                    "attribute": {"name": "Capacitance", "shortname": "cap"},
                                    "displayValue": "10 uF",
                                },
                                {
                                    "attribute": {"name": "Mounting Type", "shortname": "mount"},
                                    "displayValue": "Surface Mount",
                                },
                            ],
                        }
                    ],
                }
            ]
        }
    }


def _batch_payload(part_numbers: list[str]) -> dict[str, object]:
    return {
        "data": {
            "supMultiMatch": [
                {
                    "reference": part_number,
                    "hits": 1,
                    "parts": [
                        {
                            "mpn": part_number,
                            "manufacturer": {"name": "Example Manufacturer"},
                            "shortDescription": "Example component",
                            "specs": [],
                        }
                    ],
                }
                for part_number in part_numbers
            ]
        }
    }


def _provider(session: FakeSession) -> NexarOctopartProvider:
    return NexarOctopartProvider("client-id", "client-secret", session=session)


def test_nexar_auth_and_batch_query():
    payload = _graphql_payload()
    payload["data"]["supMultiMatch"].append(
        {"reference": "UNKNOWN", "hits": 0, "parts": []}
    )
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            FakeResponse(payload),
        ]
    )

    results = _provider(session).lookup_many(["ABC-123", "UNKNOWN"])

    assert results["ABC-123"].status == "matched"
    component = results["ABC-123"].component
    assert component is not None
    assert component.part_number == "ABC-123"
    assert component.manufacturer == "Example Manufacturer"
    assert component.short_description == "Example 10 uF capacitor"
    assert component.attributes == {"capacitance": "10 uF", "mounting": "SMD"}
    assert component.datasheet_url == ""
    assert results["UNKNOWN"].status == "not_found"
    assert len(session.calls) == 2
    auth_call, graphql_call = session.calls
    assert auth_call["url"] == "https://identity.nexar.com/connect/token"
    assert auth_call["data"] == {
        "grant_type": "client_credentials",
        "client_id": "client-id",
        "client_secret": "client-secret",
        "scope": "supply.domain",
    }
    assert graphql_call["url"] == "https://api.nexar.com/graphql"
    assert graphql_call["headers"] == {"Authorization": "Bearer token-value"}
    variables = graphql_call["json"]["variables"]
    assert variables["queries"] == [
        {"mpn": "ABC-123", "reference": "ABC-123"},
        {"mpn": "UNKNOWN", "reference": "UNKNOWN"},
    ]
    assert "ABC-123" not in graphql_call["json"]["query"]


def test_nexar_reuses_unexpired_token():
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            FakeResponse(_graphql_payload()),
            FakeResponse(_graphql_payload()),
        ]
    )
    provider = _provider(session)

    provider.lookup_many(["ABC-123"])
    provider.lookup_many(["ABC-123"])

    assert len(session.calls) == 3
    assert [call["url"] for call in session.calls].count(
        "https://identity.nexar.com/connect/token"
    ) == 1


def test_nexar_refreshes_expired_token():
    session = FakeSession(
        [
            FakeResponse({"access_token": "first-token", "expires_in": 3600}),
            FakeResponse(_graphql_payload()),
            FakeResponse({"access_token": "second-token", "expires_in": 3600}),
            FakeResponse(_graphql_payload()),
        ]
    )
    provider = _provider(session)

    provider.lookup_many(["ABC-123"])
    provider._token_expires_at = 0
    provider.lookup_many(["ABC-123"])

    auth_calls = [
        call
        for call in session.calls
        if call["url"] == "https://identity.nexar.com/connect/token"
    ]
    assert len(auth_calls) == 2
    assert session.calls[-1]["headers"] == {"Authorization": "Bearer second-token"}


def test_nexar_rejects_non_exact_mpn():
    payload = _graphql_payload()
    payload["data"]["supMultiMatch"][0]["parts"][0]["mpn"] = "ABC-123-REV-B"
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            FakeResponse(payload),
        ]
    )

    result = _provider(session).lookup_many(["ABC-123"])["ABC-123"]

    assert result.status == "unverified"
    assert result.component is None
    assert "exact" in result.message.casefold()


def test_nexar_partial_response_is_unverified_for_every_requested_part():
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            FakeResponse(_graphql_payload("ABC-123")),
        ]
    )

    results = _provider(session).lookup_many(["ABC-123", "XYZ-9"])

    assert {key: value.status for key, value in results.items()} == {
        "ABC-123": "unverified",
        "XYZ-9": "unverified",
    }
    assert all(result.component is None for result in results.values())


def test_nexar_splits_large_bom_into_batches_of_at_most_100():
    part_numbers = [f"PART-{index:03d}" for index in range(201)]
    responses = [FakeResponse({"access_token": "token-value", "expires_in": 3600})]
    responses.extend(
        FakeResponse(_batch_payload(part_numbers[start:start + 100]))
        for start in range(0, len(part_numbers), 100)
    )
    session = FakeSession(responses)

    results = _provider(session).lookup_many(part_numbers)

    assert len(results) == 201
    assert all(result.status == "matched" for result in results.values())
    graphql_calls = [call for call in session.calls if call["url"].endswith("/graphql")]
    assert [len(call["json"]["variables"]["queries"]) for call in graphql_calls] == [100, 100, 1]


def test_nexar_authentication_error_is_actionable_without_exposing_credentials():
    session = FakeSession([FakeResponse({}, status_code=401)])
    provider = NexarOctopartProvider("private-client-id", "private-secret", session=session)

    result = provider.lookup_many(["ABC-123"])["ABC-123"]

    assert result.status == "unverified"
    assert "401" in result.message
    assert "Client ID" in result.message
    assert "private-client-id" not in result.message
    assert "private-secret" not in result.message


def test_nexar_graphql_error_message_is_retained_and_credentials_redacted():
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            FakeResponse({"errors": [{"message": "Invalid app private-secret"}]}),
        ]
    )
    provider = NexarOctopartProvider("private-client-id", "private-secret", session=session)

    result = provider.lookup_many(["ABC-123"])["ABC-123"]

    assert result.status == "unverified"
    assert "Invalid app [redacted]" in result.message
    assert "private-secret" not in result.message


@pytest.mark.parametrize(
    "graphql_response",
    [
        FakeResponse({"errors": [{"message": "query failed"}]}),
        FakeResponse({}, status_code=429),
        FakeResponse(ValueError("bad json")),
    ],
)
def test_nexar_graphql_and_http_errors_are_unverified(graphql_response):
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            graphql_response,
        ]
    )

    results = _provider(session).lookup_many(["ABC-123", "XYZ-9"])

    assert {key: value.status for key, value in results.items()} == {
        "ABC-123": "unverified",
        "XYZ-9": "unverified",
    }
    assert all(result.component is None for result in results.values())


def test_nexar_timeout_is_unverified():
    session = FakeSession(
        [
            FakeResponse({"access_token": "token-value", "expires_in": 3600}),
            requests.Timeout("timed out"),
        ]
    )

    result = _provider(session).lookup_many(["ABC-123"])["ABC-123"]

    assert result.status == "unverified"
    assert result.component is None


def test_missing_credentials_fails_clearly(monkeypatch):
    monkeypatch.delenv("NEXAR_CLIENT_ID", raising=False)
    monkeypatch.delenv("NEXAR_CLIENT_SECRET", raising=False)

    with pytest.raises(ValueError, match="NEXAR_CLIENT_ID and NEXAR_CLIENT_SECRET"):
        NexarOctopartProvider()
