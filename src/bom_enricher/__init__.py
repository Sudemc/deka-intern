"""Offline-first BOM comparison and validation engine."""

from .comparison import ComparisonResult, RuleBook, compare_workbooks, load_rulebook
from .models import (
    CheckIssue,
    ComponentData,
    ComponentLookupResult,
    CustomerBomRow,
    RowResult,
)

__all__ = [
    "CheckIssue",
    "ComparisonResult",
    "ComponentData",
    "ComponentLookupResult",
    "CustomerBomRow",
    "RowResult",
    "RuleBook",
    "compare_workbooks",
    "load_rulebook",
]
