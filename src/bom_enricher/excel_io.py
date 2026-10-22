"""Excel input helpers for customer bill-of-material workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any
import unicodedata

import pandas as pd
from openpyxl import load_workbook

from bom_enricher.models import CustomerBomRow, PartNumberCandidate


@dataclass(frozen=True)
class ColumnMapping:
    """Map customer workbook headings to the fields used by the application."""

    part: str = "Part"
    qty: str = "Qty"
    value: str = "Value"
    part_number: str = "Part Number"
    manufacturer: str | None = "Manufacture"


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("İ", "i").replace("I", "i").replace("ı", "i").lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.split())


def _string_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value)


def _worksheet_name(path: Path, sheet_name: str | int) -> str:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if isinstance(sheet_name, int):
            try:
                return workbook.sheetnames[sheet_name]
            except IndexError as error:
                raise ValueError(
                    f"Workbook '{path}' has no worksheet at index {sheet_name}. "
                    f"Available worksheets: {workbook.sheetnames}."
                ) from error
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Workbook '{path}' has no worksheet named '{sheet_name}'. "
                f"Available worksheets: {workbook.sheetnames}."
            )
        return sheet_name
    finally:
        workbook.close()


_HEADER_ALIASES = {
    "item": {
        "item", "item no", "line", "line no", "sira no", "kalem no", "no", "parca",
    },
    "part": {
        "part", "reference", "ref", "refdes", "designator", "component",
        "component reference", "reference designator", "position", "location",
        "referans", "pozisyon", "konum",
    },
    "qty": {
        "qty", "quantity", "quantity required", "required qty", "count", "amount",
        "pcs", "pieces", "adet", "miktar", "sayi",
    },
    "value": {
        "value", "description", "desc", "spec", "specification", "details",
        "aciklama", "deger", "teknik ozellik",
    },
    "part_number": {
        "part number", "part no", "part n", "p n", "pn", "mpn",
        "manufacturer part number", "manufacturer part no", "manufacturer pn",
        "mfr pn", "mfr p n", "supplier part number", "vendor part number",
        "catalog number", "product id", "product code", "product number", "item code",
        "component code", "stock code", "stock number", "material number", "material no",
        "code", "kod", "urun kodu", "urun no", "parca kodu",
        "uretici kodu", "malzeme kodu", "malzeme no",
    },
    "manufacturer": {
        "manufacturer", "manufacture", "maker", "brand", "uretici", "marka",
    },
    "mounting": {
        "mounting", "th/smd", "thsmd", "th smd", "montaj", "montaj tipi", "montaj sekli", "montajtipi",
    },
    "package": {
        "package", "footprint", "case", "kilif", "paket", "size",
    },
    "status": {
        "status", "durum", "full", "dnp", "populate", "populated", "population",
        "montaj durumu", "dizgi durumu", "montajdurumu", "dizgidurumu",
    },
}
_CODE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/+()\-]{3,63}$")
_EXPLICIT_CODE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/+()\-]{1,63}$")
_REFERENCE_PATTERN = re.compile(
    r"(?i)^(?:(?:R|C|L|D|Q|U|J|FB|LS|LED|TP|SW)\s*\d+[A-Z]?)(?:\s*,\s*(?:(?:R|C|L|D|Q|U|J|FB|LS|LED|TP|SW)\s*\d+[A-Z]?))*$"
)


def _is_ref_like(val: Any) -> bool:
    s = _string_value(val).strip()
    if not s:
        return False
    first = re.split(r"[,;\s]+", s)[0]
    return bool(re.match(r"^(?:R|C|L|D|Q|U|IC|J|CN|P|TP|SW|F|FB|T|LED|BT|Y|X|K|VR|ZD|TR|TH|MOV|TS|BRK|CON|DST|FL)\d+", first, re.IGNORECASE))


def _header_field(value: Any) -> str | None:
    normalized = _normalize_header(value).replace("#", " number ")
    normalized = re.sub(r"[/._-]+", " ", normalized)
    normalized = " ".join(normalized.split())
    for field, aliases in _HEADER_ALIASES.items():
        if normalized in aliases:
            return field
    item_aliases = _HEADER_ALIASES.get("item", set())
    if normalized not in item_aliases:
        tokens = normalized.split()
        if tokens:
            last = tokens[-1]
            if last in {"kodu", "kod", "code", "pn", "mpn"} or (
                last in {"no", "number"} and not any(t in {"sira", "kalem", "line", "row", "item"} for t in tokens)
            ):
                return "part_number"
    return None



def _code_like(value: Any) -> bool:
    candidate = _string_value(value).strip()
    if len(candidate) < 5 or not _CODE_PATTERN.fullmatch(candidate):
        return False
    if not any(character.isalpha() for character in candidate) or not any(
        character.isdigit() for character in candidate
    ):
        return False
    return not _REFERENCE_PATTERN.fullmatch(candidate)


def clean_part_number(s: str) -> str:
    """Clean auxiliary notes from part numbers (e.g. bin codes or mounting orientation notes)."""
    if not s:
        return ""
    # Strip parenthetical notes that do NOT contain '/' (e.g. '(AA-BB bin)')
    cleaned = re.sub(r"\s*\([^)/]*\)", "", s)
    cleaned = re.sub(r"\s+-\s+(?:Right-Angle|Vertical|Horizontal|Dik|Yatık|SMD|TH).*$", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+\d+-\d+\s+color\)?.*$", "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


_NOTE_PREFIX_PATTERN = re.compile(
    r"(?i)^(?:SEE\b|LOOK\b|REFER\b|BAK\b|BAKINIZ\b|PLEASE\b|NOTE\b|NOT\b|TBD\b|N/?A\b|NONE\b)"
)


def _explicit_code_like(value: Any) -> bool:
    """Reject prose/footer text even when it lands under a named code column."""
    raw_candidate = _string_value(value).strip()
    if not raw_candidate or len(raw_candidate) > 128:
        return False
    if _REFERENCE_PATTERN.fullmatch(raw_candidate):
        return False
    candidate = clean_part_number(raw_candidate)
    if not candidate or _NOTE_PREFIX_PATTERN.search(candidate):
        return False
    tokens = candidate.split()
    if len(tokens) > 4:
        return False
    for token in tokens:
        if not re.fullmatch(r"[A-Za-z0-9._/+()\-*]+", token):
            return False
    if len(tokens) > 1 and not any(character.isdigit() for character in candidate):
        return False
    if not any(character.isdigit() for character in candidate) and len(candidate) > 20:
        return False
    return True



def _auto_layout(
    path: Path, sheet_name: str | int
) -> tuple[str, int, list[Any], dict[str, int], list[int]]:
    resolved_sheet = _worksheet_name(path, sheet_name)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[resolved_sheet]
        sample_rows = list(worksheet.iter_rows(min_row=1, max_row=30, values_only=True))
        if not sample_rows:
            raise ValueError(f"Workbook '{path}', worksheet '{resolved_sheet}' is empty.")
        best_row = 0
        best_score = -1
        best_fields: dict[str, list[int]] = {}
        for row_index, values in enumerate(sample_rows):
            found: dict[str, list[int]] = {}
            for column_index, value in enumerate(values):
                field = _header_field(value)
                if field:
                    found.setdefault(field, []).append(column_index)
            score = sum(min(2, len(indexes)) for indexes in found.values())
            # Recognized part-number and BOM field labels outweigh title-like rows.
            score += 2 * bool(found.get("part_number"))
            score += bool(found.get("part")) + bool(found.get("qty")) + bool(found.get("value"))
            if score > best_score:
                best_row, best_score, best_fields = row_index, score, found
        if best_score <= 0:
            # Header labels may be entirely customer-specific. Keep the conventional
            # first row as header and infer candidate-code columns from the rows below.
            best_row = 0
            best_fields = {}

        # Disambiguate 'part' (reference designator) vs item/sequence numbers
        if "part" in best_fields and len(best_fields["part"]) > 1:
            ref_scores: dict[int, int] = {}
            for col_idx in best_fields["part"]:
                ref_score = sum(
                    1 for r in sample_rows[best_row + 1:]
                    if col_idx < len(r) and r[col_idx] and _is_ref_like(r[col_idx])
                )
                ref_scores[col_idx] = ref_score
            sorted_by_ref = sorted(best_fields["part"], key=lambda c: ref_scores.get(c, 0), reverse=True)
            if ref_scores[sorted_by_ref[0]] > ref_scores.get(sorted_by_ref[1], 0):
                promoted = sorted_by_ref[0]
                demoted = [c for c in best_fields["part"] if c != promoted]
                best_fields["part"] = [promoted]
                best_fields.setdefault("item", []).extend(demoted)

        # If 'part' not found but 'item' exists, check if any 'item' column has circuit designators
        if "part" not in best_fields and "item" in best_fields:
            for col_idx in list(best_fields["item"]):
                has_refs = any(
                    col_idx < len(r) and r[col_idx] and _is_ref_like(r[col_idx])
                    for r in sample_rows[best_row + 1:]
                )
                if has_refs:
                    best_fields["part"] = [col_idx]
                    best_fields["item"].remove(col_idx)
                    break

        ambiguous_fields = {
            field: indexes
            for field, indexes in best_fields.items()
            if field != "part_number" and field != "item" and len(indexes) > 1
        }
        if ambiguous_fields:
            labels = [
                str(sample_rows[best_row][indexes[0]])
                for indexes in ambiguous_fields.values()
            ]
            raise ValueError(
                f"Workbook '{path}', worksheet '{resolved_sheet}' has ambiguous "
                f"column(s): {', '.join(labels)}. Rename duplicate columns before processing."
            )
        raw_headers = list(sample_rows[best_row])
        resolved_fields = {
            field: indexes[0]
            for field, indexes in best_fields.items()
            if len(indexes) == 1
        }
        ambiguous_code_columns = best_fields.get("part_number", [])
        code_columns = list(ambiguous_code_columns)
        excluded = set(resolved_fields.values()) | set(code_columns)
        nonempty_counts: dict[int, int] = {}
        code_counts: dict[int, int] = {}
        for values in sample_rows[best_row + 1:]:
            for column_index, value in enumerate(values):
                if value is None or not str(value).strip() or column_index in excluded:
                    continue
                nonempty_counts[column_index] = nonempty_counts.get(column_index, 0) + 1
                if _code_like(value):
                    code_counts[column_index] = code_counts.get(column_index, 0) + 1
        for column_index, count in code_counts.items():
            if count >= max(1, (nonempty_counts.get(column_index, 0) + 1) // 2):
                if column_index not in code_columns:
                    code_columns.append(column_index)
        return resolved_sheet, best_row, raw_headers, resolved_fields, code_columns
    finally:
        workbook.close()


def _validate_headers(
    path: Path, sheet_name: str | int, mapping: ColumnMapping
) -> tuple[str, dict[str, int]]:
    resolved_sheet = _worksheet_name(path, sheet_name)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[resolved_sheet]
        header_values = [cell.value for cell in next(worksheet.iter_rows(max_row=1))]
    finally:
        workbook.close()

    positions: dict[str, list[int]] = {}
    labels: dict[str, str] = {}
    for index, value in enumerate(header_values):
        normalized = _normalize_header(value)
        if not normalized:
            continue
        positions.setdefault(normalized, []).append(index)
        labels.setdefault(normalized, str(value).strip())

    duplicates = [labels[key] for key, indexes in positions.items() if len(indexes) > 1]
    if duplicates:
        duplicate_list = ", ".join(sorted(duplicates))
        raise ValueError(
            f"Workbook '{path}', worksheet '{resolved_sheet}' has duplicate "
            f"column headers: {duplicate_list}. Rename or remove the duplicate columns."
        )

    required = {
        "part": mapping.part,
        "qty": mapping.qty,
        "value": mapping.value,
        "part_number": mapping.part_number,
    }
    configured_required = [_normalize_header(name) for name in required.values()]
    if len(set(configured_required)) != len(configured_required):
        raise ValueError(
            f"Workbook '{path}', worksheet '{resolved_sheet}': column mapping assigns "
            "the same source header to multiple required fields."
        )

    missing = [
        name for name in required.values() if _normalize_header(name) not in positions
    ]
    if missing:
        missing_list = ", ".join(missing)
        raise ValueError(
            f"Workbook '{path}', worksheet '{resolved_sheet}' is missing required "
            f"column header(s): {missing_list}. Check the header row or provide a custom mapping."
        )

    resolved: dict[str, int] = {
        field: positions[_normalize_header(name)][0]
        for field, name in required.items()
    }
    if mapping.manufacturer is not None:
        manufacturer_key = _normalize_header(mapping.manufacturer)
        if manufacturer_key in positions:
            resolved["manufacturer"] = positions[manufacturer_key][0]
    return resolved_sheet, resolved


def read_customer_bom(
    path: Path,
    sheet_name: str | int | None = None,
    mapping: ColumnMapping | None = None,
) -> list[CustomerBomRow]:
    """Read customer BOM rows while validating raw worksheet headers first."""

    workbook_path = Path(path)
    if mapping is None:
        if sheet_name is None:
            sheet_name = _select_customer_sheet(workbook_path)
        resolved_sheet, header_row, raw_headers, detected, code_columns = _auto_layout(
            workbook_path, sheet_name
        )
        columns = detected
        dataframe = pd.read_excel(
            workbook_path,
            sheet_name=resolved_sheet,
            header=header_row,
            dtype=object,
        )
    else:
        resolved_sheet, columns = _validate_headers(workbook_path, sheet_name, mapping)
        header_row = 0
        raw_headers = []
        code_columns = [columns["part_number"]]
        dataframe = pd.read_excel(
            workbook_path,
            sheet_name=resolved_sheet,
            header=0,
            dtype=object,
        )
    rows: list[CustomerBomRow] = []
    for dataframe_index, values in enumerate(dataframe.itertuples(index=False, name=None)):
        get_value = lambda field: values[columns[field]] if field in columns else None
        part_number = _string_value(get_value("part_number")).strip()
        part_number_column = columns.get("part_number", -1)
        explicit_code_header = (
            mapping is not None
            or (
                0 <= part_number_column < len(raw_headers)
                and _header_field(raw_headers[part_number_column]) == "part_number"
            )
        )
        if explicit_code_header and part_number:
            if not _explicit_code_like(part_number):
                part_number = ""
            else:
                part_number = clean_part_number(part_number)
        candidates: list[PartNumberCandidate] = []
        for column_index in code_columns:
            if column_index >= len(values):
                continue
            candidate_value = _string_value(values[column_index]).strip()
            if not candidate_value:
                continue
            source_header = (
                raw_headers[column_index] if column_index < len(raw_headers) else None
            )
            candidate_has_explicit_header = _header_field(source_header) == "part_number"
            if mapping is not None and candidate_value != part_number:
                continue
            if candidate_has_explicit_header and not _explicit_code_like(candidate_value):
                continue
            if candidate_has_explicit_header and column_index == part_number_column and part_number and candidate_value != part_number:
                continue
            if mapping is None and not candidate_has_explicit_header and not _code_like(candidate_value):
                continue
            candidates.append(
                PartNumberCandidate(
                    candidate_value,
                    _string_value(source_header).strip() or f"Column {column_index + 1}",
                )
            )
        if part_number and not candidates:
            column_index = columns.get("part_number", -1)
            header = raw_headers[column_index] if 0 <= column_index < len(raw_headers) else mapping.part_number if mapping else "Part Number"
            candidates.append(PartNumberCandidate(part_number, _string_value(header).strip()))
        row_values = (get_value(field) for field in ("part", "qty", "value", "manufacturer", "mounting", "package", "status"))
        part, qty, value, manufacturer, mounting, package, status = row_values
        if not any(
            item is not None and str(item).strip()
            for item in (part, qty, value, manufacturer, mounting, package, status, part_number)
        ) and not candidates:
            continue
        rows.append(
            CustomerBomRow(
                part=_string_value(part),
                qty=qty,
                value=_string_value(value),
                part_number=part_number,
                manufacturer=_string_value(manufacturer),
                mounting=_string_value(mounting),
                package=_string_value(package),
                product_code_candidates=tuple(candidates),
                source_row=header_row + dataframe_index + 2,
                source_sheet=str(resolved_sheet),
                source_file=workbook_path.name,
                status=_string_value(status),
            )
        )
    return rows


_IGNORED_SHEET_KEYWORDS = {"bin", "revision", "revizyon", "history", "cover", "kapak", "tarihce", "doc"}


def find_customer_bom_sheets(path: Path) -> list[str]:
    """Find all worksheets in a workbook that appear to contain a BOM table."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        bom_sheets: list[str] = []
        for worksheet in workbook.worksheets:
            title_lower = worksheet.title.casefold()
            if any(kw in title_lower for kw in _IGNORED_SHEET_KEYWORDS):
                continue
            sample_rows = list(
                worksheet.iter_rows(min_row=1, max_row=30, values_only=True)
            )
            sheet_score = 0
            for values in sample_rows:
                fields = {
                    field
                    for value in values
                    if (field := _header_field(value)) is not None
                }
                if not fields:
                    continue
                score = len(fields)
                score += 8 if "part_number" in fields else 0
                score += 2 if {"part", "qty", "value"}.issubset(fields) else 0
                if score > sheet_score:
                    sheet_score = score
            if sheet_score >= 3:
                bom_sheets.append(worksheet.title)
        if not bom_sheets:
            return [_select_customer_sheet(path)]
        return bom_sheets
    finally:
        workbook.close()


def _select_customer_sheet(path: Path) -> str:
    """Choose the worksheet most likely to contain the BOM table, not its cover."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        best_sheet = workbook.sheetnames[0]
        best_score = -1
        for worksheet in workbook.worksheets:
            title_lower = worksheet.title.casefold()
            if any(kw in title_lower for kw in _IGNORED_SHEET_KEYWORDS):
                continue
            sample_rows = list(
                worksheet.iter_rows(min_row=1, max_row=30, values_only=True)
            )
            for row_index, values in enumerate(sample_rows):
                fields = {
                    field
                    for value in values
                    if (field := _header_field(value)) is not None
                }
                if not fields:
                    continue
                # A real BOM header contains several table fields. An explicit
                # manufacturer-code heading gets extra weight over cover metadata.
                score = len(fields)
                score += 8 if "part_number" in fields else 0
                score += 2 if {"part", "qty", "value"}.issubset(fields) else 0
                if score > best_score:
                    best_sheet, best_score = worksheet.title, score
        return best_sheet
    finally:
        workbook.close()

