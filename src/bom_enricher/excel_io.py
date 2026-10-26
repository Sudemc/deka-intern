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
        "sira", "kalem", "row", "row no", "entry",
    },
    "part": {
        "part", "parts", "reference", "references", "ref", "refs", "refdes", "ref des",
        "designator", "designators", "component", "components",
        "component reference", "reference designator", "position", "positions",
        "location", "locations", "component locations",
        "referans", "referanslar", "pozisyon", "konum", "devre elemani",
        "sematik referans", "eleman", "etiket", "sematik etiket",
    },
    "qty": {
        "qty", "quantity", "quantity required", "required qty", "count", "amount",
        "pcs", "pieces", "adet", "miktar", "sayi", "total qty", "qty board",
        "birim adet", "toplam adet", "parca adedi", "adedi", "sayisi",
    },
    "value": {
        "value", "val", "description", "desc", "spec", "specification", "details",
        "comment", "comments", "device", "part name", "device description",
        "aciklama", "deger", "tanim", "stok adi", "malzeme adi", "eleman degeri",
        "teknik ozellik",
    },
    "part_number": {
        "part number", "part no", "part n", "p n", "pn", "mpn",
        "manufacturer part number", "manufacturer part no", "manufacturer pn",
        "mfr pn", "mfr p n", "mfr part number", "mfr part no", "supplier part number",
        "vendor part number", "order code", "ordering code", "siparis kodu",
        "tedarikci kodu", "libref", "library reference",
        "catalog number", "product id", "product code", "product number", "item code",
        "component code", "stock code", "stock number", "material number", "material no",
        "code", "kod", "urun kodu", "urun no", "parca kodu",
        "uretici kodu", "uretici parca kodu", "malzeme kodu", "malzeme no",
    },
    "manufacturer": {
        "manufacturer", "manufacture", "maker", "brand", "uretici", "marka", "vendor", "supplier",
    },
    "mounting": {
        "mounting", "mounting type", "th smd", "thsmd", "th/smd", "montaj", "montaj tipi",
        "montaj sekli", "montajtipi", "dizgi tipi", "dizgitipi",
    },
    "package": {
        "package", "footprint", "case", "kilif", "paket", "size",
        "pcb footprint", "package case", "kasa", "boy", "govde",
    },
    "status": {
        "status", "durum", "full", "dnp", "np", "dnp np", "populate", "populated", "population",
        "fitted", "placement", "variant", "dizim", "dizim durumu",
        "montaj durumu", "dizgi durumu", "montajdurumu", "dizgidurumu",
        "process", "proses", "islem", "islem tipi", "process type",
    },
}
_DATE_PATTERN = re.compile(r"^\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4}$")


def _is_date_like(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip()
    return bool(_DATE_PATTERN.match(s))


_CODE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/+()\-]{3,63}$")
_EXPLICIT_CODE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/+()\-]{1,63}$")
_REFERENCE_PATTERN = re.compile(
    r"(?i)^(?:(?:R|C|L|D|Q|U|J|FB|LS|LED|TP|SW)\s*\d+[A-Z]?)(?:\s*,\s*(?:(?:R|C|L|D|Q|U|J|FB|LS|LED|TP|SW)\s*\d+[A-Z]?))*$"
)
_ELECTRICAL_VALUE_PATTERN = re.compile(
    r"(?i)^\d+(?:[.,]\d+)?\s*(?:[pnumkMG]?F|[kMGT]?Hz|[kMGT]?ohm|[kMR]?|V|A|W)(?:\s*[/_\s].*)?$"
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
    if len(candidate) < 5 or not _CODE_PATTERN.fullmatch(candidate) or _is_date_like(candidate):
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
    cleaned = re.sub(r"(?i)\s+BIN(?:\s*TABLE\d*|\s*:\s*[A-Za-z0-9_\-]+|\s+[A-Za-z0-9_\-]+)?.*$", "", cleaned)
    return cleaned.strip()


_NOTE_PREFIX_PATTERN = re.compile(
    r"(?i)^(?:SEE\b|LOOK\b|REFER\b|BAK\b|BAKINIZ\b|PLEASE\b|NOTE\b|NOT\b|TBD\b|N/?A\b|NONE\b)"
)


def _explicit_code_like(value: Any) -> bool:
    """Reject prose/footer text even when it lands under a named code column."""
    raw_candidate = _string_value(value).strip()
    if not raw_candidate or len(raw_candidate) > 128:
        return False
    if _REFERENCE_PATTERN.fullmatch(raw_candidate) or _is_date_like(raw_candidate):
        return False
    candidate = clean_part_number(raw_candidate)
    if not candidate or _NOTE_PREFIX_PATTERN.search(candidate) or _is_date_like(candidate):
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



def expand_ref_range(token: str) -> list[str]:
    """Expand designator ranges like 'R1-R5' or 'C10-C14' into ['R1', 'R2', 'R3', 'R4', 'R5']."""
    raw = token.strip()
    match = re.match(r"^([A-Za-z]+)(\d+)\s*[-~]\s*([A-Za-z]+)?(\d+)$", raw)
    if not match:
        return [raw]
    p1, n1_s, p2, n2_s = match.groups()
    if p2 and p1.upper() != p2.upper():
        return [raw]
    try:
        n1, n2 = int(n1_s), int(n2_s)
    except ValueError:
        return [raw]
    if n1 > n2 or n2 - n1 > 500:
        return [raw]
    prefix = p1.upper()
    return [f"{prefix}{i}" for i in range(n1, n2 + 1)]


def split_references(s: str) -> list[str]:
    """Split reference designator string (supporting commas, spaces, semicolons, and ranges)."""
    if not s:
        return []
    raw_tokens = [t.strip() for t in re.split(r"[,;|\s/]+", s) if t.strip()]
    refs: list[str] = []
    for tok in raw_tokens:
        expanded = expand_ref_range(tok)
        for r in expanded:
            if r and r not in refs:
                refs.append(r)
    return refs


def pair_lr_references(refs: list[str]) -> tuple[bool, list[tuple[str, str]]]:
    """Detect whether a designator list contains symmetrical Left and Right circuits.

    In automotive lighting BOMs, customers frequently specify 'Quantity' as the
    per-lamp / per-side count (e.g. 2 pcs), but list designators for both Left and Right
    (e.g. ST1L_B1, ST1L_B2, ST1R_B1, ST1R_B2 -> 4 pcs).

    Supports:
    - Prefix tags: 'ST1L_B1' / 'ST1R_B1', 'T1L_C1' / 'T1R_C1', 'L_R1' / 'R_R1'
    - Suffix tags: 'R1_L' / 'R1_R', 'C2-L' / 'C2-R'
    - Infix tags: 'ST-L-R1' / 'ST-R-R1', 'FOG_L_D1' / 'FOG_R_D1'
    - LH/RH variations: 'LH_R1' / 'RH_R1', 'ST_LH_R1' / 'ST_RH_R1'
    """
    if len(refs) < 2 or len(refs) % 2 != 0:
        return False, []

    ref_upper_map = {x.upper(): x for x in refs}
    l_to_r: list[tuple[str, str]] = []
    matched_r: set[str] = set()

    for r in refs:
        m_lh = re.match(r"^([A-Za-z0-9]*?)(?:_)?LH(?:_)?([A-Za-z0-9_]*)$", r, re.IGNORECASE)
        m_pre = re.match(r"^([A-Za-z0-9]*?)(L)_([A-Za-z0-9]+)$", r, re.IGNORECASE)
        m_suf = re.match(r"^([A-Za-z0-9]+)([_\-])(L)$", r, re.IGNORECASE)
        m_inf = re.match(r"^([A-Za-z0-9]+)([_\-])(L)([_\-])([A-Za-z0-9]+)$", r, re.IGNORECASE)

        target = None
        if m_lh:
            p1, p2 = m_lh.groups()
            delim1 = "_" if "_" in r[:len(p1)+1] else ""
            delim2 = "_" if "_" in r[len(p1):] else ""
            target = f"{p1}{delim1}RH{delim2}{p2}"
        elif m_pre:
            target = f"{m_pre.group(1)}R_{m_pre.group(3)}"
        elif m_suf:
            target = f"{m_suf.group(1)}{m_suf.group(2)}R"
        elif m_inf:
            target = f"{m_inf.group(1)}{m_inf.group(2)}R{m_inf.group(4)}{m_inf.group(5)}"

        if target and target.upper() in ref_upper_map and target.upper() not in matched_r:
            actual_r = ref_upper_map[target.upper()]
            l_to_r.append((r, actual_r))
            matched_r.add(target.upper())

    is_fully_paired = (len(l_to_r) * 2 == len(refs))
    return is_fully_paired, l_to_r


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
            else:
                best_fields["part"] = [sorted_by_ref[0]]

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

        # If 'part' still not found, check all unmapped columns for reference-like values
        if "part" not in best_fields:
            candidate_part_scores: dict[int, int] = {}
            mapped_cols = {c for idxs in best_fields.values() for c in idxs}
            for col_idx in range(len(sample_rows[best_row])):
                if col_idx in mapped_cols:
                    continue
                score = sum(
                    1 for r in sample_rows[best_row + 1:]
                    if col_idx < len(r) and r[col_idx] and _is_ref_like(r[col_idx])
                )
                if score >= 2:
                    candidate_part_scores[col_idx] = score
            if candidate_part_scores:
                best_fields["part"] = [max(candidate_part_scores, key=candidate_part_scores.get)]

        # Disambiguate 'qty'
        if "qty" in best_fields and len(best_fields["qty"]) > 1:
            qty_scores: dict[int, int] = {}
            for col_idx in best_fields["qty"]:
                score = sum(
                    1 for r in sample_rows[best_row + 1:]
                    if col_idx < len(r) and r[col_idx] is not None and any(c.isdigit() for c in str(r[col_idx]))
                )
                qty_scores[col_idx] = score
            def _qty_rank(c: int) -> tuple[int, int]:
                hdr = _normalize_header(sample_rows[best_row][c]) if c < len(sample_rows[best_row]) else ""
                is_total = 1 if ("toplam" in hdr or "total" in hdr) else 0
                return (qty_scores.get(c, 0), is_total)
            best_fields["qty"] = [max(best_fields["qty"], key=_qty_rank)]

        # Disambiguate 'mounting'
        if "mounting" in best_fields and len(best_fields["mounting"]) > 1:
            mount_scores: dict[int, int] = {}
            for col_idx in best_fields["mounting"]:
                score = sum(
                    1 for r in sample_rows[best_row + 1:]
                    if col_idx < len(r) and r[col_idx] and re.search(r"(?i)\b(?:smd|th|tht|surface|through|hole)\b", str(r[col_idx]))
                )
                mount_scores[col_idx] = score
            best_fields["mounting"] = [max(best_fields["mounting"], key=lambda c: mount_scores.get(c, 0))]

        # Disambiguate 'status'
        if "status" in best_fields and len(best_fields["status"]) > 1:
            status_scores: dict[int, int] = {}
            for col_idx in best_fields["status"]:
                score = sum(
                    1 for r in sample_rows[best_row + 1:]
                    if col_idx < len(r) and r[col_idx] and re.search(r"(?i)\b(?:dnp|np|fitted|not fitted|fit|no fit|nc|montaj|dizim)\b", str(r[col_idx]))
                )
                status_scores[col_idx] = score
            best_fields["status"] = [max(best_fields["status"], key=lambda c: status_scores.get(c, 0))]

        # Preserve multi-column fields (e.g. Value and Description, Footprint and Package)
        all_value_columns = list(best_fields.get("value", []))
        all_package_columns = list(best_fields.get("package", []))

        # Fallback resolution for any remaining duplicate columns (pick highest non-empty rate)
        for field, indexes in list(best_fields.items()):
            if field not in {"part_number", "item"} and len(indexes) > 1:
                non_empty = {
                    c: sum(1 for r in sample_rows[best_row + 1:] if c < len(r) and r[c] is not None and str(r[c]).strip())
                    for c in indexes
                }
                best_fields[field] = [max(indexes, key=lambda c: non_empty.get(c, 0))]

        raw_headers = list(sample_rows[best_row])
        resolved_fields = {
            field: indexes[0]
            for field, indexes in best_fields.items()
            if len(indexes) == 1
        }
        if all_value_columns:
            resolved_fields["_value_columns"] = all_value_columns
        if all_package_columns:
            resolved_fields["_package_columns"] = all_package_columns

        ambiguous_code_columns = list(best_fields.get("part_number", []))
        if len(ambiguous_code_columns) > 1:
            def _pn_priority(col_idx: int) -> int:
                hdr = _normalize_header(sample_rows[best_row][col_idx]) if col_idx < len(sample_rows[best_row]) else ""
                if any(kw in hdr for kw in ("part number", "part no", "mpn", "mfr pn", "uretici", "manufacturer")):
                    return 10
                if any(kw in hdr for kw in ("feka", "internal", "stok", "stock", "catalog", "siparis", "order")):
                    return 5
                return 1
            best_pn_col = max(ambiguous_code_columns, key=_pn_priority)
            resolved_fields["part_number"] = best_pn_col

        code_columns = list(ambiguous_code_columns)
        excluded = {c for c in resolved_fields.values() if isinstance(c, int)} | set(code_columns) | set(all_value_columns) | set(all_package_columns)
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
    _METADATA_ROW_KEYWORDS = (
        "project name", "creation date", "revision number", "design responsible",
        "last modified", "proje adi", "revizyon no", "bin table", "revision description",
        "see pcb spec", "pcb spec. document", "pcb spec document",
    )
    for dataframe_index, values in enumerate(dataframe.itertuples(index=False, name=None)):
        row_str_lower = " ".join(str(v).lower() for v in values if v is not None)
        if any(kw in row_str_lower for kw in _METADATA_ROW_KEYWORDS):
            if not any(_is_ref_like(v) for v in values if v is not None):
                continue

        get_value = lambda field: values[columns[field]] if field in columns and isinstance(columns[field], int) else None
        part_number = _string_value(get_value("part_number")).strip()
        if _is_date_like(part_number):
            part_number = ""
        part_number_column = columns.get("part_number", -1) if isinstance(columns.get("part_number"), int) else -1
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
            if not candidate_value or _is_date_like(candidate_value):
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
            column_index = columns.get("part_number", -1) if isinstance(columns.get("part_number"), int) else -1
            header = raw_headers[column_index] if 0 <= column_index < len(raw_headers) else mapping.part_number if mapping else "Part Number"
            candidates.append(PartNumberCandidate(part_number, _string_value(header).strip()))

        # Combine multiple value columns (e.g. Value and Description) without duplication
        val_cols = columns.get("_value_columns") or ([columns["value"]] if "value" in columns and isinstance(columns["value"], int) else [])
        val_parts: list[str] = []
        for col_idx in val_cols:
            if col_idx < len(values):
                val_txt = _string_value(values[col_idx]).strip()
                if val_txt and val_txt.lower() != "nan" and val_txt not in val_parts:
                    val_parts.append(val_txt)
        combined_value = " - ".join(val_parts)

        # In CAD exports without a dedicated Part Number column (e.g. KiCad / Eagle),
        # check if the Value column holds a component part number rather than an electrical value
        if not code_columns and mapping is None and not part_number and not candidates:
            if combined_value and _code_like(combined_value) and not _ELECTRICAL_VALUE_PATTERN.match(combined_value):
                candidates.append(PartNumberCandidate(combined_value, "Value"))
                part_number = combined_value

        # Combine multiple package columns (e.g. Footprint and Package)
        pkg_cols = columns.get("_package_columns") or ([columns["package"]] if "package" in columns and isinstance(columns["package"], int) else [])
        pkg_parts: list[str] = []
        for col_idx in pkg_cols:
            if col_idx < len(values):
                pkg_txt = _string_value(values[col_idx]).strip()
                if pkg_txt and pkg_txt.lower() != "nan" and pkg_txt not in pkg_parts:
                    pkg_parts.append(pkg_txt)
        combined_package = " ".join(pkg_parts)

        part = get_value("part")
        qty = get_value("qty")
        manufacturer = get_value("manufacturer")
        mounting = get_value("mounting")
        status = get_value("status")

        # Fallback: if qty is missing or empty, infer count from circuit references
        part_str = _string_value(part).strip()
        if (qty is None or str(qty).strip() == "" or str(qty).lower() == "nan") and part_str:
            refs = split_references(part_str)
            if refs:
                qty = len(refs)
        elif part_str and qty is not None:
            # Check for automotive L/R symmetrical pairing
            # In automotive lighting BOMs, customers frequently specify 'Quantity' as the
            # per-lamp / per-side count (e.g. 2 pcs), but list designators for both Left and Right
            # (e.g. ST1L_B1, ST1L_B2, ST1R_B1, ST1R_B2 -> 4 pcs).
            try:
                raw_qty_num = int(float(str(qty).strip()))
                refs = split_references(part_str)
                if refs and len(refs) == 2 * raw_qty_num:
                    is_paired, _ = pair_lr_references(refs)
                    if is_paired:
                        qty = len(refs)
            except (ValueError, TypeError):
                pass

        if not any(
            item is not None and str(item).strip()
            for item in (part, qty, combined_value, manufacturer, mounting, combined_package, status, part_number)
        ) and not candidates:
            continue
        rows.append(
            CustomerBomRow(
                part=_string_value(part),
                qty=qty,
                value=combined_value,
                part_number=part_number,
                manufacturer=_string_value(manufacturer),
                mounting=_string_value(mounting),
                package=combined_package,
                product_code_candidates=tuple(candidates),
                source_row=header_row + dataframe_index + 2,
                source_sheet=str(resolved_sheet),
                source_file=workbook_path.name,
                status=_string_value(status),
            )
        )
    return rows


_IGNORED_SHEET_KEYWORDS = {
    "bin", "revision", "revizyon", "history", "cover", "kapak", "tarihce", "tarihçe",
    "doc", "dokuman", "özeti", "ozeti", "summary", "kickoff", "spec", "specification",
    "gerber", "readme", "data", "lookup", "dropdown", "secim", "seçim", "parametre",
}


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
            valid_sheets = [
                ws for ws in workbook.worksheets
                if not any(kw in ws.title.casefold() for kw in _IGNORED_SHEET_KEYWORDS)
            ]
            if not valid_sheets:
                return []
            cand = _select_customer_sheet(path)
            cand_rows = list(workbook[cand].iter_rows(min_row=1, max_row=20, values_only=True))
            has_refs = any(_is_ref_like(c) for r in cand_rows for c in r if c is not None)
            has_codes = any(_code_like(c) or _explicit_code_like(c) for r in cand_rows for c in r if c is not None)
            has_headers = any(_header_field(c) is not None for r in cand_rows for c in r if c is not None)
            if has_refs or has_codes or has_headers:
                return [cand]
            return []
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

