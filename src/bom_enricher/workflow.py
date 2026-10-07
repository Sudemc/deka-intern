"""Reusable BOM processing workflow shared by the CLI and web interface."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path

from bom_enricher.component_data import (
    JsonCatalogProvider,
    NexarOctopartProvider,
    normalize_part_number,
)
from bom_enricher.excel_io import read_customer_bom
from bom_enricher.models import (
    CheckIssue,
    ComponentLookupResult,
    PartNumberCandidate,
    RowResult,
    UnresolvedPartNumber,
)
from bom_enricher.output import write_erp_workbook
from bom_enricher.validation import validate_row


@dataclass(frozen=True)
class ProcessingReport:
    row_results: list[RowResult]
    unresolved_parts: list[UnresolvedPartNumber]
    counts: dict[str, int]


def process_bom(
    input_path: Path,
    output_path: Path,
    provider_name: str = "sample",
    catalog_path: Path | None = None,
    sheet_name: str | int = 0,
) -> ProcessingReport:
    """Validate a workbook and write the ERP-shaped result workbook."""
    source = Path(input_path).expanduser()
    destination = Path(output_path).expanduser()
    if source.resolve() == destination.resolve():
        raise ValueError("Output workbook must use a different path from the source.")

    rows = read_customer_bom(source, sheet_name=sheet_name)
    provider = _make_provider(provider_name, catalog_path)
    part_numbers = list(dict.fromkeys(
        candidate.value.strip()
        for row in rows
        for candidate in _row_candidates(row)
        if candidate.value.strip()
    ))
    lookups = provider.lookup_many(part_numbers)
    results: list[RowResult] = []
    unresolved_parts: list[UnresolvedPartNumber] = []
    for row in rows:
        candidates = _row_candidates(row)
        exact_candidates = [
            candidate
            for candidate in candidates
            if _is_exact_match(candidate, lookups.get(candidate.value.strip()))
            and _manufacturer_agrees(
                row.manufacturer,
                lookups[candidate.value.strip()].component.manufacturer,
            )
        ]
        if len(exact_candidates) == 1:
            selected_candidate = exact_candidates[0]
            part_number = selected_candidate.value.strip()
            lookup = lookups[part_number]
        elif len(candidates) == 1:
            selected_candidate = candidates[0]
            part_number = selected_candidate.value.strip()
            lookup = lookups.get(part_number, _missing_lookup(part_number))
        else:
            selected_candidate = None
            part_number = ""
            lookup = ComponentLookupResult(
                status="ambiguous" if candidates else "unverified",
                component=None,
                message=(
                    "More than one code-like value was found and none uniquely matched the provider."
                    if candidates
                    else "No product-code value could be identified from this row."
                ),
            )

        resolved_row = replace(row, part_number=part_number)
        results.append(validate_row(resolved_row, lookup))
        selected_is_verified = (
            selected_candidate is not None
            and _is_exact_match(selected_candidate, lookup)
            and _manufacturer_agrees(row.manufacturer, lookup.component.manufacturer)
        )
        if not selected_is_verified and candidates:
            for candidate in candidates:
                candidate_lookup = lookups.get(candidate.value.strip())
                reason = (
                    candidate_lookup.message
                    if candidate_lookup is not None and candidate_lookup.message
                    else "No unique exact product-code match was verified."
                )
                if candidate_lookup is not None and candidate_lookup.status == "not_found":
                    reason = "No exact online/catalog match was found for this code."
                elif candidate_lookup is not None and candidate_lookup.status == "matched":
                    reason = "Candidate code matched, but manufacturer or identity was ambiguous."
                unresolved_parts.append(
                    UnresolvedPartNumber(
                        part_number=candidate.value,
                        source_column=candidate.source_column,
                        source_row=row.source_row,
                        part=row.part,
                        status="unverified",
                        reason=reason,
                    )
                )

    write_erp_workbook(results, destination, unresolved_parts)
    return ProcessingReport(
        row_results=results,
        unresolved_parts=unresolved_parts,
        counts=_summary_counts([result.issues for result in results], len(results)),
    )


def _make_provider(provider_name: str, catalog_path: Path | None):
    if provider_name == "sample":
        selected_catalog = catalog_path or (
            Path(__file__).resolve().parents[2]
            / "examples"
            / "component_catalog_sample.json"
        )
        return JsonCatalogProvider(selected_catalog)
    if provider_name == "nexar":
        return NexarOctopartProvider()
    raise ValueError(f"Unsupported component data provider: {provider_name}")


def _summary_counts(issue_rows: list[list[CheckIssue]], row_count: int) -> Counter[str]:
    counts: Counter[str] = Counter()
    for issues in issue_rows:
        for issue in issues:
            if issue.status in {"added", "corrected", "invalid"}:
                counts[issue.status] += 1
            elif issue.status in {"conflict", "unverified"}:
                counts["unresolved"] += 1
    counts["unchanged"] = max(0, row_count * 4 - sum(counts.values()))
    return counts


def _missing_lookup(part_number: str) -> ComponentLookupResult:
    return ComponentLookupResult(
        status="unverified",
        component=None,
        message=f"Provider returned no result for part number {part_number!r}.",
    )


def _row_candidates(row) -> tuple[PartNumberCandidate, ...]:
    if row.product_code_candidates:
        return row.product_code_candidates
    if row.part_number.strip():
        return (PartNumberCandidate(row.part_number, "Part Number"),)
    return ()


def _is_exact_match(candidate: PartNumberCandidate, lookup) -> bool:
    component = lookup.component if lookup is not None else None
    return bool(
        lookup is not None
        and lookup.status == "matched"
        and component is not None
        and normalize_part_number(candidate.value)
        == normalize_part_number(component.part_number)
    )


def _manufacturer_agrees(source: str, matched: str) -> bool:
    if not source.strip():
        return True
    return " ".join(source.casefold().split()) == " ".join(matched.casefold().split())
