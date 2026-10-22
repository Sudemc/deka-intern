from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


LookupStatus = Literal["matched", "not_found", "ambiguous", "unverified"]
IssueStatus = Literal["added", "corrected", "conflict", "unverified", "invalid"]


@dataclass(frozen=True)
class PartNumberCandidate:
    value: str
    source_column: str


@dataclass(frozen=True)
class CustomerBomRow:
    part: str
    qty: object
    value: str
    part_number: str
    manufacturer: str = ""
    mounting: str = ""
    package: str = ""
    product_code_candidates: tuple[PartNumberCandidate, ...] = ()
    source_row: int = 0
    source_sheet: str = ""
    source_file: str = ""
    status: str = ""


@dataclass(frozen=True)
class ComponentData:
    part_number: str
    manufacturer: str
    short_description: str
    attributes: dict[str, str]
    datasheet_url: str = ""


@dataclass(frozen=True)
class ComponentLookupResult:
    status: LookupStatus
    component: ComponentData | None
    message: str = ""


@dataclass(frozen=True)
class CheckIssue:
    field: str
    status: IssueStatus
    original: str
    proposed: str
    message: str


@dataclass(frozen=True)
class RowResult:
    source: CustomerBomRow
    description: str
    mounting: str
    issues: list[CheckIssue]


@dataclass(frozen=True)
class UnresolvedPartNumber:
    part_number: str
    source_column: str
    source_row: int
    part: str
    status: str
    reason: str
