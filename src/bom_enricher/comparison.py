"""Excel-to-Excel BOM comparison and ERP rule validation."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .excel_io import (
    clean_part_number,
    find_customer_bom_sheets,
    read_customer_bom,
    split_references,
)
from .semantic_parser import compare_descriptions, normalize_code_key, normalize_mounting


OUTPUT_HEADERS = (
    "Durum",
    "Müşteri Part Number",
    "ERP Part Number",
    "Müşteri Qty",
    "ERP Qty",
    "Müşteri Value",
    "ERP Description",
    "Müşteri TH/SMD",
    "ERP TH/SMD",
    "Kaynak Kart / BOM",
    "Müşteri Referans",
    "ERP Referans",
    "Hata/Uyarı Durumu",
)
_MAX_SHEET_SCAN_ROWS = 30


def _classify_severity(issues: list[str]) -> str:
    """Classify issues into 'HATA' (red), 'UYARI' (yellow), or 'UYGUN' (normal)."""
    if not issues:
        return "UYGUN"
    error_keywords = (
        "Miktar Uyuşmazlığı",
        "Referans sayısı uyumsuz",
        "Sistemde Yok",
        "Montaj Tipi Çelişkisi",
        "Voltaj farklı",
        "Kılıf/Paket farklı",
        "yinelenen Part Number",
        "TH/SMD ile Description",
        "Kart Ön Eki",
        "LED ve Renkler",
    )
    for issue in issues:
        if any(tag in issue for tag in ("Montaj Dışı", "ERP'ye Girilmez", "ERP Sarf Malzemesi", "Ham PCB Kalemi", "Dummy")):
            continue
        if any(keyword in issue for keyword in error_keywords):
            return "HATA"
    return "UYARI"


@dataclass(frozen=True)
class PartRecord:
    part_number: str
    description: str
    mounting: str = ""
    reference: str = ""
    material_type: str = ""
    qty: str = ""
    manufacturer: str = ""
    package: str = ""
    duplicate: bool = False
    source_file: str = ""
    source_sheet: str = ""
    source_breakdown: str = ""
    unmapped_prefix: bool = False
    all_codes: tuple[str, ...] = ()
    code_labels: dict[str, str] = field(default_factory=dict)
    status: str = ""



@dataclass(frozen=True)
class ComparisonResult:
    rows: list[dict[str, str]]
    total_part_count: int


@dataclass(frozen=True)
class RuleEntry:
    """One complete rule row, including type, ordering, language and code affixes."""

    material_type: str
    attribute_code: str
    attribute_name: str
    value: str
    value_description: str
    language: str
    sort_order: str
    prefix: str
    suffix: str


@dataclass(frozen=True)
class RuleBook:
    """The source dictionary indexed by material type and attribute name."""

    entries: tuple[RuleEntry, ...] = ()
    by_type_attribute: dict[str, dict[str, tuple[RuleEntry, ...]]] = field(default_factory=dict)
    by_attribute: dict[str, tuple[RuleEntry, ...]] = field(default_factory=dict)
    manufacturers: tuple[tuple[str, str, str], ...] = ()

    @property
    def material_types(self) -> set[str]:
        return set(self.by_type_attribute)



def _is_product_code_token(token: str) -> bool:
    """Return True if *token* looks like a standalone product code (not a short word like MDD)."""
    stripped = re.sub(r"[^A-Za-z0-9]", "", token)
    return len(stripped) >= 4 and any(c.isdigit() for c in stripped)


def expand_parentheses_variants(code: str) -> list[str]:
    """Expand slash packaging variants inside parentheses.

    For example:
    - 'CL10B104KB8NNN(C/D)' -> ['CL10B104KB8NNN(C/D)', 'CL10B104KB8NNNC', 'CL10B104KB8NNND']
    - '0805(S/W)8F1002T(5/D)E' -> ['0805(S/W)8F1002T(5/D)E', '0805S8F1002T5E', ...]
    """
    match = re.search(r"\(([A-Za-z0-9]+(?:/[A-Za-z0-9]+)+)\)", code)
    if not match:
        return [code]
    options = match.group(1).split("/")
    prefix = code[: match.start()]
    suffix = code[match.end() :]
    results: list[str] = []
    for opt in options:
        variant = f"{prefix}{opt}{suffix}"
        results.extend(expand_parentheses_variants(variant))
    return list(dict.fromkeys([code] + results))


def _split_multi_code(raw: str) -> list[str]:
    """Split an ERP Part Number cell that contains space/comma/slash separated alternative codes.

    Returns a list of individual code strings including packaging variants.
    If the cell holds only a single code (e.g. 'BAV99 MDD', 'TE 2324 DPN'),
    it returns the full code (plus any parentheses variants).
    """
    cleaned_raw = re.sub(r"(?i)\s+BIN(?:\s*TABLE\d*|\s*:\s*[A-Za-z0-9_\-]+|\s+[A-Za-z0-9_\-]+)?.*$", "", raw)
    cleaned_raw = re.sub(r"\s+-\s+(?:Right-Angle|Vertical|Horizontal|Dik|Yatık|SMD|TH).*$", "", cleaned_raw, flags=re.IGNORECASE)
    cleaned = re.sub(r"[,;|\n\r]+", " ", cleaned_raw)
    cleaned = re.sub(r"\s+/\s+", " ", cleaned)
    tokens = cleaned.split()

    results: list[str] = []
    # Search for parenthetical or bracketed codes (e.g. 'DELTA-V (IXTT140P10T)' or '[4E22000001]')
    for m in re.finditer(r"[\(\[]([A-Za-z0-9._/+()\-]+)[\)\]]", cleaned_raw):
        inner = m.group(1).strip()
        if _is_product_code_token(inner):
            results.extend(expand_parentheses_variants(inner))

    for t in tokens:
        m = re.match(r"^([A-Za-z0-9._/+()\-]+)\((.+)\)$", t)
        if m:
            outer = m.group(1).strip()
            inner = m.group(2).strip()
            if _is_product_code_token(outer):
                results.extend(expand_parentheses_variants(outer))
            if _is_product_code_token(inner):
                results.extend(expand_parentheses_variants(inner))
            results.append(t)
        else:
            t_clean = t.strip("()[]")
            if _is_product_code_token(t_clean):
                results.extend(expand_parentheses_variants(t_clean))
            elif _is_product_code_token(t):
                results.extend(expand_parentheses_variants(t))

    if results:
        results.append(cleaned_raw.strip())
        return list(dict.fromkeys(results))

    return expand_parentheses_variants(cleaned_raw.strip())


def _normalize_ref_token(token: str) -> str:
    """Normalize delimiter in designator token like TR1.C1 or TR1-C1 to TR1_C1."""
    tok = token.strip().upper()
    parts = re.split(r"[_\.\-]", tok, maxsplit=1)
    if len(parts) == 2 and 2 <= len(parts[0]) <= 10 and re.match(r"^[A-Z]+\d+", parts[1]):
        return f"{parts[0]}_{parts[1]}"
    return tok


def _normalize_ref_string(s: str) -> str:
    if not s:
        return ""
    cleaned = re.sub(r"\b([A-Za-z0-9]+_)\s+([A-Za-z0-9]+)\b", r"\1\2", s)
    tokens = split_references(cleaned)
    return ", ".join(_normalize_ref_token(t) for t in tokens if t.strip())


def _split_refs(s: str) -> set[str]:
    """Split reference designators into normalized uppercase tokens, expanding ranges."""
    if not s:
        return set()
    cleaned = re.sub(r"\b([A-Za-z0-9]+_)\s+([A-Za-z0-9]+)\b", r"\1\2", s)
    return {_normalize_ref_token(t) for t in split_references(cleaned) if t.strip()}


def _ref_base_component(ref: str) -> tuple[str, str, str]:
    """Return (board_prefix, base_ref, comp_type) from a designator token."""
    norm = _normalize_ref_token(ref)
    if "_" in norm:
        prefix, base = norm.split("_", 1)
    else:
        prefix, base = "", norm
    base = base.strip().upper()
    match = re.match(r"^([A-Z]+)\d+", base)
    comp_type = match.group(1) if match else (base[:1] if base else "")
    return prefix.strip().upper(), base, comp_type


def _should_pair_refs(crefs: set[str], erefs: set[str]) -> bool:
    """Determine whether two sets of reference designators should be paired."""
    if not crefs or not erefs:
        return False
    if crefs == erefs:
        return True
    overlap = crefs & erefs
    if overlap:
        c_types = {_ref_base_component(r)[2] for r in crefs if r}
        e_types = {_ref_base_component(r)[2] for r in erefs if r}
        if len(c_types) == 1 and c_types == e_types and list(c_types)[0] in "CRLDQUJSF":
            if crefs.issubset(erefs) or erefs.issubset(crefs):
                return True
            if len(overlap) / min(len(crefs), len(erefs)) >= 0.5:
                return True

    # Check prefix-stripped base references
    c_bases = {_ref_base_component(r)[1] for r in crefs if r}
    e_bases = {_ref_base_component(r)[1] for r in erefs if r}
    if c_bases and e_bases and c_bases == e_bases:
        c_types = {_ref_base_component(r)[2] for r in crefs if r}
        e_types = {_ref_base_component(r)[2] for r in erefs if r}
        if len(c_types) == 1 and c_types == e_types and list(c_types)[0] in "CRLDQUJSF":
            return True

    return False


def _extract_erp_prefixes(erp_records: list[PartRecord]) -> tuple[set[str], dict[str, set[str]]]:
    """Extract board prefixes (e.g. TR1, ST1, FM) and their base references from ERP."""
    refs_by_prefix: dict[str, set[str]] = defaultdict(set)
    pfx_record_count: dict[str, int] = defaultdict(int)
    for rec in erp_records:
        seen_in_rec: set[str] = set()
        for ref in _split_refs(rec.reference):
            if "_" in ref:
                parts = ref.split("_", 1)
                prefix = parts[0].strip().upper()
                base = parts[1].strip().upper()
                # A PCB board/sub-panel prefix must be at least 2 characters (single letters like J, C, R, D
                # are schematic component class identifiers, e.g. J_LCD1, D_LED1, not board prefixes).
                if 2 <= len(prefix) <= 10 and re.match(r"^[A-Z]+\d+", base):
                    refs_by_prefix[prefix].add(base)
                    seen_in_rec.add(prefix)
        for pfx in seen_in_rec:
            pfx_record_count[pfx] += 1

    # Genuine board prefixes must appear across multiple component designators or records (at least 2)
    valid_prefixes = {
        pfx
        for pfx, bases in refs_by_prefix.items()
        if len(bases) >= 2 or pfx_record_count[pfx] >= 2
    }
    return valid_prefixes, {pfx: refs_by_prefix[pfx] for pfx in valid_prefixes}


def _resolve_record_prefix(
    source_sheet: str,
    source_file: str,
    record_refs: set[str],
    erp_prefixes: set[str],
    erp_refs_by_prefix: dict[str, set[str]],
) -> str | None:
    """Determine the ERP board prefix corresponding to a customer sheet or file."""
    if not erp_prefixes:
        return None

    norm_sheet = re.sub(r"[^A-Za-z0-9]", "", source_sheet.upper())
    norm_file = re.sub(r"[^A-Za-z0-9]", "", Path(source_file).stem.upper())

    # 1. Exact match on normalized sheet name
    for pfx in erp_prefixes:
        norm_pfx = re.sub(r"[^A-Za-z0-9]", "", pfx.upper())
        if norm_sheet == norm_pfx:
            return pfx

    # 2. Sheet name starts with prefix
    for pfx in sorted(erp_prefixes, key=len, reverse=True):
        norm_pfx = re.sub(r"[^A-Za-z0-9]", "", pfx.upper())
        if norm_sheet.startswith(norm_pfx) and len(norm_pfx) >= 2:
            return pfx

    # 2b. Token & Acronym matching from sheet name (e.g. FOG1 PCB -> F1, FOGMAIN PCB -> FM, STOP TAIL 1 PCB -> ST1, TURN REV 2 PCB -> TR2)
    cleaned_sheet = re.sub(r"(?i)\b(?:PCB|ASSY|BOARD|DEVRE|KART|PANEL|SHEET|SAYFA)\b", "", source_sheet).strip()
    spaced_sheet = re.sub(r"(?i)(FOG|STOP|TAIL|TURN|REV|MAIN|LAMP|UPPER|LOWER|REAR|FRONT|LEFT|RIGHT)", r" \1 ", cleaned_sheet)
    sheet_tokens = [t.strip().upper() for t in re.split(r"[\s_\-]+", spaced_sheet) if t.strip()]
    if sheet_tokens:
        sheet_acro = "".join(t if t.isdigit() else t[0] for t in sheet_tokens)
        if sheet_acro in erp_prefixes:
            return sheet_acro
        if f"{sheet_acro}L" in erp_prefixes and f"{sheet_acro}R" in erp_prefixes:
            return f"{sheet_acro}L+{sheet_acro}R"
        raw_upper_sheet = "".join(c for c in cleaned_sheet if c.isupper() or c.isdigit()).upper()
        if raw_upper_sheet in erp_prefixes:
            return raw_upper_sheet
        if f"{raw_upper_sheet}L" in erp_prefixes and f"{raw_upper_sheet}R" in erp_prefixes:
            return f"{raw_upper_sheet}L+{raw_upper_sheet}R"
        for pfx in erp_prefixes:
            if sheet_acro.startswith(pfx) or pfx == sheet_acro:
                return pfx

    # 3. Exact match on filename stem
    for pfx in erp_prefixes:
        norm_pfx = re.sub(r"[^A-Za-z0-9]", "", pfx.upper())
        if norm_file == norm_pfx:
            return pfx

    # 4. Token & Acronym matching from filename stem (e.g. StopTail1 -> ST1, FogMain -> FM, TurnRev1 -> TR1)
    tokens = re.split(r"[_ \-]+", Path(source_file).stem)
    for tok in reversed(tokens):
        tok_clean = tok.strip()
        if not tok_clean:
            continue
        tok_upper = tok_clean.upper()
        if tok_upper in erp_prefixes:
            return tok_upper
        acronym = "".join(c for c in tok_clean if c.isupper() or c.isdigit()).upper()
        if acronym in erp_prefixes:
            return acronym
        for pfx in erp_prefixes:
            if tok_upper.startswith(pfx.upper()) and (len(tok_upper) == len(pfx) or not tok_upper[len(pfx)].isalpha()):
                return pfx

    # 5. Filename stem starts with prefix
    for pfx in sorted(erp_prefixes, key=len, reverse=True):
        norm_pfx = re.sub(r"[^A-Za-z0-9]", "", pfx.upper())
        if norm_file.startswith(norm_pfx) and len(norm_pfx) >= 2:
            return pfx

    # 6. Fallback: match by RefDes overlap
    if record_refs:
        best_pfx = None
        best_overlap = 0
        for pfx, bases in erp_refs_by_prefix.items():
            overlap = len(record_refs & bases)
            if overlap > best_overlap:
                best_overlap = overlap
                best_pfx = pfx
        if best_pfx and best_overlap >= 3 and (
            best_overlap / len(record_refs) >= 0.25 or best_overlap / len(erp_refs_by_prefix[best_pfx]) >= 0.4
        ):
            return best_pfx

    return None


def _qualify_customer_references(
    customer_records: list[PartRecord],
    erp_prefixes: set[str],
    erp_refs_by_prefix: dict[str, set[str]],
) -> list[PartRecord]:
    """Qualify customer RefDes with ERP board prefixes (e.g. C10 -> TR1_C10)."""
    if not erp_prefixes:
        return customer_records

    sheet_refs: dict[tuple[str, str], set[str]] = defaultdict(set)
    for rec in customer_records:
        key = (rec.source_file, rec.source_sheet)
        for ref in _split_refs(rec.reference):
            _, base, _ = _ref_base_component(ref)
            if base:
                sheet_refs[key].add(base)

    resolved_prefixes: dict[tuple[str, str], str] = {}
    for key, refs in sheet_refs.items():
        src_file, src_sheet = key
        pfx = _resolve_record_prefix(src_sheet, src_file, refs, erp_prefixes, erp_refs_by_prefix)
        if pfx:
            resolved_prefixes[key] = pfx

    satellite_pairs: dict[tuple[str, str], tuple[str, str]] = {}
    for key, pfx in list(resolved_prefixes.items()):
        if "+" in pfx:
            parts = pfx.split("+")
            satellite_pairs[key] = (parts[0], parts[1])

    for key in sheet_refs:
        if key not in resolved_prefixes:
            src_file, src_sheet = key
            cleaned = re.sub(r"(?i)\b(?:PCB|ASSY|BOARD|DEVRE|KART|PANEL|SHEET|SAYFA)\b", "", src_sheet).strip()
            spaced = re.sub(r"(?i)(FOG|STOP|TAIL|TURN|REV|MAIN|LAMP|UPPER|LOWER|REAR|FRONT|LEFT|RIGHT)", r" \1 ", cleaned)
            tokens = [t.strip().upper() for t in re.split(r"[\s_\-]+", spaced) if t.strip()]
            if tokens:
                acro = "".join(t if t.isdigit() else t[0] for t in tokens)
                l_pfx, r_pfx = f"{acro}L", f"{acro}R"
                if l_pfx in erp_prefixes and r_pfx in erp_prefixes:
                    satellite_pairs[key] = (l_pfx, r_pfx)
                    resolved_prefixes[key] = f"{l_pfx}+{r_pfx}"

    unmapped_keys = {key for key in sheet_refs if key not in resolved_prefixes and key not in satellite_pairs}

    updated: list[PartRecord] = []
    for rec in customer_records:
        key = (rec.source_file, rec.source_sheet)
        rec_refs = _split_refs(rec.reference)
        already_qualified = bool(
            rec_refs and any(
                any(r.upper().startswith(f"{pfx.upper()}_") for pfx in erp_prefixes)
                for r in rec_refs
            )
        )

        if key in satellite_pairs and not already_qualified:
            l_pfx, r_pfx = satellite_pairs[key]
            clean_tokens = [r for r in rec_refs if "_" not in r]
            if clean_tokens:
                expanded = [f"{l_pfx}_{t}" for t in clean_tokens] + [f"{r_pfx}_{t}" for t in clean_tokens]
                new_ref = ", ".join(expanded)
                new_qty = rec.qty
                if rec.qty:
                    try:
                        new_qty = str(int(float(rec.qty)) * 2)
                    except (ValueError, TypeError):
                        new_qty = str(len(expanded))
                else:
                    new_qty = str(len(expanded))
                updated.append(PartRecord(
                    part_number=rec.part_number,
                    description=rec.description,
                    mounting=rec.mounting,
                    reference=new_ref,
                    material_type=rec.material_type,
                    qty=new_qty,
                    manufacturer=rec.manufacturer,
                    package=rec.package,
                    duplicate=rec.duplicate,
                    source_file=rec.source_file,
                    source_sheet=rec.source_sheet,
                    source_breakdown=rec.source_breakdown,
                    unmapped_prefix=False,
                    all_codes=rec.all_codes,
                    code_labels=rec.code_labels,
                    status=rec.status,
                ))
                continue

        pfx = resolved_prefixes.get(key)
        is_unmapped = bool(unmapped_keys and key in unmapped_keys and not already_qualified)
        if not pfx or not rec.reference.strip() or already_qualified:
            updated.append(PartRecord(
                part_number=rec.part_number,
                description=rec.description,
                mounting=rec.mounting,
                reference=_normalize_ref_string(rec.reference),
                material_type=rec.material_type,
                qty=rec.qty,
                manufacturer=rec.manufacturer,
                package=rec.package,
                duplicate=rec.duplicate,
                source_file=rec.source_file,
                source_sheet=rec.source_sheet,
                source_breakdown=rec.source_breakdown,
                unmapped_prefix=is_unmapped,
                all_codes=rec.all_codes,
                code_labels=rec.code_labels,
                status=rec.status,
            ))
            continue

        clean_ref = re.sub(r"\b([A-Za-z0-9]+_)\s+([A-Za-z0-9]+)\b", r"\1\2", rec.reference)
        raw_tokens = re.split(r"([,;|\s/]+)", clean_ref)
        new_tokens: list[str] = []
        erp_bases = erp_refs_by_prefix.get(pfx, set())
        for token in raw_tokens:
            stripped = _normalize_ref_token(token.strip().upper())
            if stripped and re.match(r"^[A-Z]+\d+", stripped):
                if not stripped.startswith(f"{pfx}_") and "_" not in stripped:
                    new_tokens.append(f"{pfx}_{stripped}")
                else:
                    new_tokens.append(stripped)
            elif stripped and re.match(r"^[A-Za-z0-9]+_[A-Za-z]+\d+", stripped):
                new_tokens.append(stripped)
            elif token in (",", ";", "|", "/", " "):
                new_tokens.append(token)
            elif not stripped or stripped.endswith("_"):
                continue
            else:
                new_tokens.append(token)

        new_ref = "".join(new_tokens)
        updated.append(PartRecord(
            part_number=rec.part_number,
            description=rec.description,
            mounting=rec.mounting,
            reference=new_ref,
            material_type=rec.material_type,
            qty=rec.qty,
            manufacturer=rec.manufacturer,
            package=rec.package,
            duplicate=rec.duplicate,
            source_file=rec.source_file,
            source_sheet=rec.source_sheet,
            source_breakdown=rec.source_breakdown,
            unmapped_prefix=False,
            all_codes=rec.all_codes,
            code_labels=rec.code_labels,
            status=rec.status,
        ))
    return updated


def _check_certain_non_assembly(record: PartRecord) -> tuple[bool, str]:
    """Determine with certainty if an item should not be entered in the PCB assembly ERP BOM.

    Returns (True, reason) only when there is unequivocal evidence that the item
    is not intended for board assembly (e.g. explicit DNP/NP, Fiducial optical marks,
    or dummy mechanical accessories without product codes).
    """
    status_upper = (record.status or "").strip().upper()
    val_upper = (record.description or "").strip().upper()
    ref_upper = (record.reference or "").strip().upper()
    pkg_upper = (record.package or "").strip().upper()
    mount_upper = (record.mounting or "").strip().upper()
    has_code = bool((record.part_number or "").strip()) or bool(record.all_codes)

    # 1. Explicit DNP / NP / NM marker in status column (e.g. FULL column = DNP / NP / Process = NM)
    dnp_statuses = {
        "DNP", "NP", "N.P.", "D.N.P.", "DO NOT POPULATE", "NOT POPULATED",
        "NO POP", "DİZİLMEYECEK", "DİZİLMEZ", "DIZILMEYECEK", "DIZILMEZ", "MONTAJ YOK",
        "MONTAJ DIŞI", "MONTAJ DISI", "NM", "N.M.", "NOT MOUNTED", "NOT MOUNT",
    }
    if status_upper in dnp_statuses:
        return True, f"Montaj Dışı / Dizilmez ({record.status.strip()})"

    # 2. Fiducial optical camera alignment markers (PCB copper features, not components)
    if (
        ref_upper.startswith("FL_STD")
        or ref_upper.startswith("FID")
        or val_upper == "FIDUCIAL"
        or pkg_upper.startswith("FL_STD")
    ):
        return True, "Fiducial / Optik Hizalama İşareti (Fiziksel parça değildir)"


    # 3. Explicit DNP / NP in Description/Value or Mounting without any product code
    if not has_code:
        if (
            val_upper in dnp_statuses
            or mount_upper in {"DNP", "NP"}
            or any(tag in val_upper for tag in ("(DNP)", "[DNP]", "(NP)", "[NP]", "NOT POPULATED", "DO NOT POPULATE"))
        ):
            return True, "DNP / Monte Edilmeyecek Parça (ERP'ye girilmedi)"

    # 4. Purely mechanical accessories / dummy items without any product code
    if not has_code:
        if mount_upper == "MECH" or pkg_upper in {"DUMMY", "DGHT PCB"}:
            detail = f"{record.package or record.mounting}".strip()
            return True, f"Mekanik / Harici Montaj Parçası ({detail})" if detail else "Mekanik / Harici Montaj Parçası"

    return False, ""


def _check_erp_consumable_or_pcb(record: PartRecord) -> tuple[bool, str]:
    """Check if an ERP record is an SMT consumable, bare PCB, or dummy component."""
    ref_upper = (record.reference or "").strip().upper()
    val_upper = (record.description or "").strip().upper()
    pn_upper = (record.part_number or "").strip().upper()
    combined = f"{ref_upper} {val_upper} {pn_upper}"

    # Solder paste / production consumables
    if any(k in combined for k in ("KREM LEHIM", "KREMLEHIM", "SOLDER PASTE", "LEHIM", "M705", "FLUX", "TINER", "THINNER", "SERIT", "TAPE", "SARF")):
        return True, "Sarf Malzeme (Krem Lehim / SMT Üretim Sarfı)"

    # Bare PCB
    if any(k in combined for k in ("BARE PCB", "HAM PCB", "PCB ODAK", "FR4", "D047-", "D047", "BASKI DEVRE", "BASKI KART")) or ref_upper == "PCB":
        return True, "Ham PCB Kalemi (Müşteri spesifikasyon dokümanında tanımlı çıplak devre kartı)"

    # Dummy items
    if "DUMMY" in combined or "NONE/PN" in combined or "DOLGU" in combined:
        return True, "Dummy / Dolgu Kalemi (ERP üretim reçetesi yer tutucu malzeme)"

    return False, ""


def compare_workbooks(
    customer_path: Path | list[Path] | str | list[str],
    erp_path: Path | str,
    rules_path: Path | str | None = None,
    output_path: Path | str | None = None,
) -> ComparisonResult:
    """Compare customer/ERP rows by MPN and RefDes, saving only rows requiring review."""
    rules = load_rulebook(Path(rules_path)) if rules_path else None
    customer = load_records(customer_path, kind="customer")
    erp_raw = load_records(erp_path, kind="erp")

    erp_prefixes, erp_refs_by_prefix = _extract_erp_prefixes(erp_raw)
    if erp_prefixes:
        customer = _qualify_customer_references(customer, erp_prefixes, erp_refs_by_prefix)

    # ── ERP çoklu Part Number genişletme ──
    # ERP hücresinde boşlukla ayrılmış birden fazla alternatif üretici kodu
    # varsa (ör. "CL21B104KBCNNNC 885012207098 CC0805KRX7R9BB104") her kodu
    # ayrı PartRecord olarak genişlet.  Aynı hücreden gelen kardeş kodları
    # izle: eğer birisi müşteri tarafında eşleşirse diğerlerini raporlama.
    erp: list[PartRecord] = []
    sibling_groups: list[dict[str, object]] = []   # her grup: primary kod ve grubun tüm normalized kodları
    for rec in erp_raw:
        codes = _split_multi_code(rec.part_number)
        norm_ref = _normalize_ref_string(rec.reference)
        cleaned_first = clean_part_number(rec.part_number)
        if cleaned_first and cleaned_first not in codes:
            codes.append(cleaned_first)
        if len(codes) > 1:
            group: set[str] = set()
            for code in codes:
                erp.append(PartRecord(
                    part_number=code,
                    description=rec.description,
                    mounting=rec.mounting,
                    reference=norm_ref,
                    material_type=rec.material_type,
                    qty=rec.qty,
                    manufacturer=rec.manufacturer,
                    package=rec.package,
                    source_file=rec.source_file,
                    source_sheet=rec.source_sheet,
                    all_codes=tuple(codes),
                ))
                group.add(_normalize_code(code))
            sibling_groups.append({
                "primary": _normalize_code(codes[0]),
                "all": group,
            })
        else:
            erp.append(PartRecord(
                part_number=cleaned_first or rec.part_number,
                description=rec.description,
                mounting=rec.mounting,
                reference=norm_ref,
                material_type=rec.material_type,
                qty=rec.qty,
                manufacturer=rec.manufacturer,
                package=rec.package,
                duplicate=rec.duplicate,
                source_file=rec.source_file,
                source_sheet=rec.source_sheet,
            ))

    customer_by_code = _unique_records(customer)
    erp_by_code = _unique_records(erp)

    # ── Üç Aşamalı Eşleme (Three-Pass Matching) ──
    # Aşama 1: Normalized Part Number üzerinden birincil eşleme
    matched_pairs: list[tuple[PartRecord | None, PartRecord | None]] = []
    unmatched_cust: dict[str, PartRecord] = dict(customer_by_code)
    unmatched_erp: dict[str, PartRecord] = dict(erp_by_code)

    active_cust_keys = [k for k in unmatched_cust.keys() if not k.endswith("__NM")]
    for ck in active_cust_keys:
        if ck.startswith("__NO_CODE_"):
            continue
        c_rec = unmatched_cust.get(ck)
        if not c_rec:
            continue
        cand_keys = [ek for ek in unmatched_erp if ek == ck or ek.startswith(f"{ck}__sub_")]
        if not cand_keys:
            for alt_code in c_rec.all_codes:
                alt_norm = _normalize_code(alt_code)
                sub_cands = [ek for ek in unmatched_erp if ek == alt_norm or ek.startswith(f"{alt_norm}__sub_")]
                if sub_cands:
                    cand_keys = sub_cands
                    break

        if cand_keys:
            best_ek = cand_keys[0]
            if len(cand_keys) > 1 and c_rec.reference:
                crefs = _split_refs(c_rec.reference)
                for ek_cand in cand_keys:
                    erefs = _split_refs(unmatched_erp[ek_cand].reference)
                    if crefs == erefs:
                        best_ek = ek_cand
                        break
            c_matched = unmatched_cust.pop(ck)
            e_matched = unmatched_erp.pop(best_ek)

            # Check if this ERP record aggregates multiple customer rows of alternative MPNs
            # (e.g. 38 pcs Yageo + 2 pcs Vishay matching an ERP line of 40 pcs)
            if c_matched.qty and e_matched.qty and e_matched.reference and c_matched.reference:
                try:
                    c_q = float(c_matched.qty)
                    e_q = float(e_matched.qty)
                    if c_q < e_q:
                        erefs = _split_refs(e_matched.reference)
                        crefs = _split_refs(c_matched.reference)
                        missing_refs = erefs - crefs
                        for other_ck, other_c in list(unmatched_cust.items()):
                            if other_ck.endswith("__NM") or not other_c.reference:
                                continue
                            o_refs = _split_refs(other_c.reference)
                            if o_refs and o_refs.issubset(missing_refs):
                                e_codes = set(e_matched.all_codes) | set(_split_multi_code(e_matched.part_number))
                                e_codes_norm = {_normalize_code(c) for c in e_codes if c}
                                o_codes_norm = {_normalize_code(c) for c in other_c.all_codes if c}
                                if (e_codes_norm & o_codes_norm) or _normalize_code(other_c.part_number) in e_codes_norm:
                                    merged_refs = sorted(crefs | o_refs)
                                    new_c_q = c_q + (float(other_c.qty) if other_c.qty else len(o_refs))
                                    c_matched = PartRecord(
                                        part_number=f"{c_matched.part_number} / {other_c.part_number}",
                                        description=c_matched.description,
                                        mounting=c_matched.mounting,
                                        reference=", ".join(merged_refs),
                                        material_type=c_matched.material_type,
                                        qty=str(int(new_c_q)) if new_c_q.is_integer() else str(new_c_q),
                                        package=c_matched.package,
                                        manufacturer=f"{c_matched.manufacturer} / {other_c.manufacturer}" if other_c.manufacturer and other_c.manufacturer != c_matched.manufacturer else c_matched.manufacturer,
                                        duplicate=c_matched.duplicate,
                                        source_file=c_matched.source_file,
                                        source_sheet=f"{c_matched.source_sheet}, {other_c.source_sheet}",
                                        source_breakdown=f"{c_matched.source_breakdown}; {other_c.source_breakdown}",
                                        unmapped_prefix=c_matched.unmapped_prefix,
                                        all_codes=tuple(dict.fromkeys(c_matched.all_codes + other_c.all_codes)),
                                        code_labels={**c_matched.code_labels, **other_c.code_labels},
                                        status=c_matched.status,
                                    )
                                    crefs = set(merged_refs)
                                    missing_refs = erefs - crefs
                                    c_q = new_c_q
                                    unmatched_cust.pop(other_ck)
                                    if not missing_refs or c_q >= e_q:
                                        break
                except (ValueError, TypeError):
                    pass

            # Check if this ERP record aggregates an active customer row and an NM customer row
            nm_candidate_keys = [f"{ck}__NM"] + [
                f"{_normalize_code(ac)}__NM" for ac in c_matched.all_codes if ac
            ]
            matched_nm_key = next((k for k in dict.fromkeys(nm_candidate_keys) if k in unmatched_cust), None)
            if matched_nm_key and c_matched.reference and e_matched.reference:
                c_nm = unmatched_cust[matched_nm_key]
                crefs = _split_refs(c_matched.reference)
                nm_refs = _split_refs(c_nm.reference)
                erefs = _split_refs(e_matched.reference)
                if crefs and nm_refs and crefs.issubset(erefs) and nm_refs.issubset(erefs) and not (crefs & nm_refs):
                    unmatched_cust.pop(matched_nm_key)
                    e_active = PartRecord(
                        part_number=e_matched.part_number,
                        description=e_matched.description,
                        mounting=e_matched.mounting,
                        reference=", ".join(sorted(crefs)),
                        material_type=e_matched.material_type,
                        qty=str(len(crefs)),
                        manufacturer=e_matched.manufacturer,
                        package=e_matched.package,
                        duplicate=e_matched.duplicate,
                        source_file=e_matched.source_file,
                        source_sheet=e_matched.source_sheet,
                        source_breakdown=e_matched.source_breakdown,
                        unmapped_prefix=e_matched.unmapped_prefix,
                        all_codes=e_matched.all_codes,
                        code_labels=e_matched.code_labels,
                        status=e_matched.status,
                    )
                    remaining_erefs = erefs - crefs
                    e_nm = PartRecord(
                        part_number=e_matched.part_number,
                        description=e_matched.description,
                        mounting=e_matched.mounting,
                        reference=", ".join(sorted(remaining_erefs)),
                        material_type=e_matched.material_type,
                        qty=str(len(remaining_erefs)),
                        manufacturer=e_matched.manufacturer,
                        package=e_matched.package,
                        duplicate=e_matched.duplicate,
                        source_file=e_matched.source_file,
                        source_sheet=e_matched.source_sheet,
                        source_breakdown=e_matched.source_breakdown,
                        unmapped_prefix=e_matched.unmapped_prefix,
                        all_codes=e_matched.all_codes,
                        code_labels=e_matched.code_labels,
                        status=e_matched.status,
                    )
                    matched_pairs.append((c_matched, e_active))
                    matched_pairs.append((c_nm, e_nm))
                    continue

            matched_pairs.append((c_matched, e_matched))
            continue

    # Aşama 2: RefDes eşleme (Part Number farklı fakat devre pozisyonu aynı olanlar)
    # Önce tam eşleşenler (exact match crefs == erefs), ardından kısmi örtüşenler
    for pass_mode in ("exact", "fuzzy"):
        for ck in list(unmatched_cust.keys()):
            c_rec = unmatched_cust.get(ck)
            if not c_rec:
                continue
            crefs = _split_refs(c_rec.reference)
            if not crefs:
                continue
            for ek in list(unmatched_erp.keys()):
                e_rec = unmatched_erp.get(ek)
                if not e_rec:
                    continue
                erefs = _split_refs(e_rec.reference)
                if not erefs:
                    continue
                match = False
                if pass_mode == "exact" and crefs == erefs:
                    match = True
                elif pass_mode == "fuzzy" and _should_pair_refs(crefs, erefs):
                    match = True

                if match:
                    matched_pairs.append((c_rec, e_rec))
                    del unmatched_cust[ck]
                    del unmatched_erp[ek]
                    break

    # Aşama 3: Açıklama (Description) ile Eşleme
    # Eğer ERP'deki bir parça müşteri tarafında kod veya RefDes ile bulunamadıysa,
    # fakat müşteri listesinde Part Number'ı boş olan bir satırın açıklaması ERP açıklamasıyla
    # birebir aynıysa (veya normalize edilmiş hali uyuşuyorsa) bunları eşleştir:
    for ek in list(unmatched_erp.keys()):
        e_rec = unmatched_erp.get(ek)
        if not e_rec or not e_rec.description.strip():
            continue
        e_desc_norm = _normalize_description(e_rec.description)
        for ck in list(unmatched_cust.keys()):
            c_rec = unmatched_cust.get(ck)
            if not c_rec:
                continue
            if not c_rec.part_number.strip() and c_rec.description.strip():
                c_desc_norm = _normalize_description(c_rec.description)
                if c_desc_norm == e_desc_norm:
                    matched_pairs.append((c_rec, e_rec))
                    del unmatched_cust[ck]
                    del unmatched_erp[ek]
                    break

    # ── Kardeş kod filtreleme ──
    # Çoklu Part Number hücresinden gelen alternatif kodlardan biri müşteri
    # tarafında eşleşmişse, eşleşmeyen kardeş kodları rapordan çıkar.
    matched_erp_normalized_codes = {
        _normalize_code(e.part_number)
        for c, e in matched_pairs
        if e is not None
    }
    filtered_unmatched_erp: list[PartRecord] = []
    for ek, e_rec in unmatched_erp.items():
        code_norm = _normalize_code(e_rec.part_number)
        skip = False
        for group in sibling_groups:
            if code_norm in group["all"]:
                if group["all"] & matched_erp_normalized_codes:
                    skip = True
                elif code_norm != group["primary"]:
                    skip = True
                break
        if not skip:
            filtered_unmatched_erp.append(e_rec)

    # Eşleşmeyen yalnız-müşteri ve yalnız-ERP kayıtları
    for c_rec in unmatched_cust.values():
        matched_pairs.append((c_rec, None))
    for e_rec in filtered_unmatched_erp:
        matched_pairs.append((None, e_rec))

    # Sabit ve tutarlı sıralama
    matched_pairs.sort(key=lambda pair: ((pair[0] or pair[1]).part_number.upper() if (pair[0] or pair[1]) else ""))

    report: list[dict[str, str]] = []
    for customer_record, erp_record in matched_pairs:
        c_pn = customer_record.part_number.strip() if customer_record else ""
        e_pn = erp_record.part_number.strip() if erp_record else ""

        pn_matches = False
        if customer_record and erp_record and e_pn:
            e_norm = _normalize_code(e_pn)
            if c_pn and _normalize_code(c_pn) == e_norm:
                pn_matches = True
            elif any(_normalize_code(c) == e_norm for c in customer_record.all_codes):
                pn_matches = True

        # Part Number gösterimi
        if customer_record and erp_record:
            if pn_matches:
                display_code = c_pn or e_pn
            elif c_pn and e_pn and _normalize_code(c_pn) != _normalize_code(e_pn):
                display_code = f"{c_pn} / {e_pn}"
            else:
                display_code = c_pn or e_pn
        elif customer_record:
            display_code = c_pn or customer_record.reference or "(Part Number Yok)"
        else:
            display_code = e_pn

        customer_value = customer_record.description if customer_record else ""
        erp_description = erp_record.description if erp_record else ""
        customer_qty = customer_record.qty if customer_record else ""
        erp_qty = erp_record.qty if erp_record else ""
        customer_mounting = customer_record.mounting if customer_record else ""
        erp_mounting = erp_record.mounting if erp_record else ""

        # ERP veya müşteride montaj sütunu boşsa açıklamadan (TH / SMD) otomatik tespit et:
        if not erp_mounting and erp_description:
            inferred_erp = normalize_mounting(erp_description)
            if inferred_erp:
                erp_mounting = inferred_erp

        if not customer_mounting and customer_value:
            inferred_cust = normalize_mounting(customer_value)
            if inferred_cust:
                customer_mounting = inferred_cust
        issues: list[str] = []

        if customer_record is None:
            is_consumable, cons_reason = _check_erp_consumable_or_pcb(erp_record)
            if is_consumable:
                issues.append(f"Montaj Dışı / ERP Sarf Malzemesi: {cons_reason}")
            else:
                issues.append("Sistemde Yok: müşteri dosyasında bulunamadı")
        elif erp_record is None:
            is_non_erp, non_erp_reason = _check_certain_non_assembly(customer_record)
            if is_non_erp:
                issues.append(f"Montaj Dışı / ERP'ye Girilmez: {non_erp_reason}")
            else:
                issues.append("Sistemde Yok: ERP dosyasında bulunamadı")
                if not c_pn:
                    issues.append("Part Number eksik (Müşteri dosyasında Part Number belirtilmemiş)")

        if customer_record and customer_record.duplicate:
            issues.append("Müşteri dosyasında yinelenen Part Number")
        if erp_record and erp_record.duplicate:
            issues.append("ERP dosyasında yinelenen Part Number")
        if customer_record and customer_record.unmapped_prefix:
            src = customer_record.source_file or "Müşteri dosyası"
            issues.append(f"Önek Eşleşemedi: ERP kart öneki tespit edilemedi ({src})")

        # Part Number eksik uyarısı (müşteri ve ERP eşleşti ama müşteri part number boş)
        if customer_record and erp_record and not c_pn and not customer_record.all_codes:
            issues.append("Part Number eksik (Müşteri dosyasında Part Number belirtilmemiş)")

        pn_diff_note = ""
        if customer_record and erp_record and c_pn and e_pn and not pn_matches:
            pn_diff_note = f"Part Number farklı (Müşteri: {c_pn}, ERP: {e_pn})"

        if customer_record and erp_record:
            if "NM" in (customer_record.status or "").upper():
                issues.append("Montaj Dışı (Process: NM): Müşteri listesinde bu parça 'Process: NM' (Dizilmez) olarak tanımlıdır. SMT dizgide dizilmemelidir.")

            # 1. Qty (Miktar) Kıyaslaması
            if customer_qty and erp_qty:
                try:
                    c_num = float(customer_qty)
                    e_num = float(erp_qty)
                    if c_num != e_num:
                        issues.append(f"Miktar Uyuşmazlığı (Müşteri: {customer_qty}, ERP: {erp_qty})")
                except ValueError:
                    if customer_qty.strip() != erp_qty.strip():
                        issues.append(f"Miktar Uyuşmazlığı (Müşteri: {customer_qty}, ERP: {erp_qty})")

            # 2. Müşteri Qty vs Referans Adedi
            if customer_record.reference and customer_qty:
                ref_tokens = [r.strip() for r in re.split(r"[,;|\s/]+", customer_record.reference) if r.strip()]
                valid_refs = [
                    r for r in ref_tokens
                    if re.match(r"^(?:[A-Za-z0-9]+_)?[A-Za-z][A-Za-z0-9_]*$", r)
                ]
                if valid_refs:
                    try:
                        c_qty_val = float(customer_qty)
                        if c_qty_val != len(valid_refs) and c_qty_val != len(ref_tokens):
                            issues.append(
                                f"Müşteri Qty ile Referans sayısı uyumsuz (Qty: {customer_qty}, Ref: {len(valid_refs)})"
                            )
                    except ValueError:
                        pass

            # 3. Dizgi/Referans Farkı (Müşteri vs ERP)
            c_refs = _split_refs(customer_record.reference)
            e_refs = _split_refs(erp_record.reference)
            if c_refs and e_refs and c_refs != e_refs:
                c_only = sorted(c_refs - e_refs)
                e_only = sorted(e_refs - c_refs)
                c_bases = {_ref_base_component(r)[1]: _ref_base_component(r)[0] for r in c_only if _ref_base_component(r)[1]}
                e_bases = {_ref_base_component(r)[1]: _ref_base_component(r)[0] for r in e_only if _ref_base_component(r)[1]}
                prefix_mismatches = [
                    f"{c_bases[b]}_{b} (Müşteri) vs {e_bases[b]}_{b} (ERP)"
                    for b in c_bases if b in e_bases and c_bases[b] != e_bases[b]
                ]
                if prefix_mismatches:
                    issues.append(
                        f"Kart Ön Eki / Referans Uyuşmazlığı: {', '.join(prefix_mismatches)}"
                    )
                else:
                    diff_parts = []
                    if c_only:
                        diff_parts.append(f"Müşteride Farklı/Fazla: {', '.join(c_only)}")
                    if e_only:
                        diff_parts.append(f"ERP'de Farklı/Fazla: {', '.join(e_only)}")
                    diff_summary = f"Fark: {' | '.join(diff_parts)}; " if diff_parts else ""
                    issues.append(
                        f"Dizgi/Referans farkı ({diff_summary}Müşteri: {customer_record.reference}, ERP: {erp_record.reference})"
                    )

            # 4. Müşteri Montaj vs ERP Montaj
            if customer_mounting and erp_mounting:
                c_mount = normalize_mounting(customer_mounting)
                e_mount = normalize_mounting(erp_mounting)
                if c_mount and e_mount and c_mount != e_mount:
                    issues.append(f"Montaj Tipi Çelişkisi (Müşteri: {customer_mounting}, ERP: {erp_mounting})")

            # 5. Semantik Açıklama ve Teknik Değer Kıyaslaması
            issues.extend(compare_descriptions(customer_value, erp_description))

        if erp_record:
            if customer_record and customer_record.material_type and erp_record.material_type:
                if _normalize_label(customer_record.material_type) != _normalize_label(erp_record.material_type):
                    issues.append("Müşteri ve ERP malzeme tipi çelişkisi")
            validation_record = PartRecord(
                part_number=erp_record.part_number,
                description=erp_record.description,
                mounting=erp_record.mounting,
                reference=erp_record.reference or (customer_record.reference if customer_record else ""),
                material_type=erp_record.material_type or (customer_record.material_type if customer_record else ""),
                qty=erp_record.qty,
                package=erp_record.package or (customer_record.package if customer_record else ""),
                manufacturer=erp_record.manufacturer or (customer_record.manufacturer if customer_record else ""),
            )
            issues.extend(validate_erp_rules(validation_record, rules) if rules else [])

        if issues and pn_diff_note:
            issues.insert(0, pn_diff_note)

        durum = _classify_severity(issues)
        issue_str = "; ".join(dict.fromkeys(issues))

        if customer_record:
            if len(customer_record.all_codes) > 1:
                codes_list = list(customer_record.all_codes)
                if e_pn and any(_normalize_code(c) == _normalize_code(e_pn) for c in codes_list):
                    matched_c = next(c for c in codes_list if _normalize_code(c) == _normalize_code(e_pn))
                    other_codes = [c for c in codes_list if c != matched_c]
                    customer_pn = f"{matched_c} [{' / '.join(other_codes)}]"
                else:
                    customer_pn = " / ".join(codes_list)
            else:
                customer_pn = c_pn
        else:
            customer_pn = ""
        erp_pn = e_pn
        source_label = (
            (customer_record.source_breakdown or customer_record.source_file)
            if customer_record
            else ""
        )
        cust_reference = customer_record.reference if customer_record else ""
        erp_reference = erp_record.reference if erp_record else ""

        report.append({
            "Durum": durum,
            "Müşteri Part Number": customer_pn,
            "ERP Part Number": erp_pn,
            "Part Number": display_code,
            "Müşteri Qty": customer_qty,
            "ERP Qty": erp_qty,
            "Müşteri Value": customer_value,
            "ERP Description": erp_description,
            "Müşteri TH/SMD": customer_mounting,
            "ERP TH/SMD": erp_mounting,
            "Kaynak Kart / BOM": source_label,
            "Müşteri Referans": cust_reference,
            "ERP Referans": erp_reference,
            "Hata/Uyarı Durumu": issue_str,
        })

    if output_path:
        write_comparison_report(report, output_path)
    return ComparisonResult(report, len(matched_pairs))


def load_records(path: Path | list[Path] | str | list[str], kind: str) -> list[PartRecord]:
    """Load customer or ERP rows, finding a header row and common header aliases."""
    if kind not in {"customer", "erp"}:
        raise ValueError("kind must be 'customer' or 'erp'.")
    if kind == "customer":
        paths = [Path(p) for p in path] if isinstance(path, (list, tuple)) else [Path(path)]
        all_records: list[PartRecord] = []
        for file_path in paths:
            bom_sheets = find_customer_bom_sheets(file_path)
            for sheet in bom_sheets:
                customer_rows = read_customer_bom(file_path, sheet_name=sheet)
                for row in customer_rows:
                    code = row.part_number
                    candidates = list(dict.fromkeys(
                        candidate.value.strip()
                        for candidate in row.product_code_candidates
                        if candidate.value.strip()
                    ))
                    code_labels: dict[str, str] = {}
                    for candidate in row.product_code_candidates:
                        val = candidate.value.strip()
                        if val and val not in code_labels:
                            code_labels[val] = candidate.source_column

                    if not code:
                        if len(candidates) == 1:
                            code = candidates[0]
                        elif len(candidates) > 1:
                            mpn_cand = next(
                                (c for c in candidates if any(k in code_labels.get(c, "").lower() for k in ("üretici", "uretici", "mfr", "manufacturer", "mpn"))),
                                None
                            )
                            code = mpn_cand or candidates[0]

                    all_codes_list = [code] if code else []
                    for c in candidates:
                        if c not in all_codes_list:
                            all_codes_list.append(c)
                    all_codes = tuple(clean_part_number(c) for c in all_codes_list if clean_part_number(c))

                    if code:
                        code = clean_part_number(code)
                        qty_val = _cell_text(row.qty).strip()
                        if qty_val.lower() == "nan":
                            qty_val = ""
                        if not qty_val and row.part:
                            refs = _split_refs(row.part)
                            if refs:
                                qty_val = str(len(refs))
                        all_records.append(PartRecord(
                            part_number=code,
                            description=row.value,
                            reference=row.part,
                            mounting=row.mounting,
                            qty=qty_val,
                            manufacturer=row.manufacturer,
                            package=row.package,
                            source_file=file_path.name,
                            source_sheet=sheet,
                            all_codes=all_codes,
                            code_labels=code_labels,
                            status=row.status,
                        ))
                    else:
                        qty_str = _cell_text(row.qty).strip()
                        if qty_str.lower() == "nan":
                            qty_str = ""
                        if not qty_str and row.part:
                            refs = _split_refs(row.part)
                            if refs:
                                qty_str = str(len(refs))
                        has_qty = bool(qty_str) and any(c.isdigit() for c in qty_str)
                        ref_str = (row.part or "").strip()
                        circuit_ref_pattern = r"^(?:R|C|L|D|Q|U|IC|J|CN|P|TP|SW|F|FB|T|LED|BT|Y|X|K|VR|ZD|TR|TH|MOV|TS|BRK)\d+"
                        has_ref = bool(ref_str and re.match(circuit_ref_pattern, ref_str, re.IGNORECASE))
                        if has_qty or (has_ref and bool((row.value or "").strip())):
                            all_records.append(PartRecord(
                                part_number="",
                                description=row.value,
                                reference=row.part,
                                mounting=row.mounting,
                                qty=qty_str,
                                manufacturer=row.manufacturer,
                                package=row.package,
                                source_file=file_path.name,
                                source_sheet=sheet,
                                status=row.status,
                            ))
        if not all_records:
            first_name = paths[0].name if paths else "Müşteri dosyası"
            raise ValueError(f"{first_name}: ürün kodu veya malzeme satırı bulunamadı.")

        # Cross-sheet internal stock code to MPN resolution
        # If a row's part_number was filled from an internal stock code (e.g. 4E12000022),
        # but other rows in the customer files with the same internal stock code have a real MPN,
        # adopt that MPN so rows of the same component across sheets group seamlessly.
        internal_to_mpn: dict[str, str] = {}
        for r in all_records:
            if r.part_number and not r.part_number.startswith("4E"):
                for c in r.all_codes:
                    if c.startswith("4E") and c not in internal_to_mpn:
                        internal_to_mpn[c] = r.part_number
        if internal_to_mpn:
            updated_records: list[PartRecord] = []
            for r in all_records:
                if r.part_number in internal_to_mpn:
                    updated_records.append(PartRecord(
                        part_number=internal_to_mpn[r.part_number],
                        description=r.description,
                        mounting=r.mounting,
                        reference=r.reference,
                        material_type=r.material_type,
                        qty=r.qty,
                        package=r.package,
                        manufacturer=r.manufacturer,
                        duplicate=r.duplicate,
                        source_file=r.source_file,
                        source_sheet=r.source_sheet,
                        source_breakdown=r.source_breakdown,
                        unmapped_prefix=r.unmapped_prefix,
                        all_codes=r.all_codes,
                        code_labels=r.code_labels,
                        status=r.status,
                    ))
                else:
                    updated_records.append(r)
            all_records = updated_records

        return all_records

    erp_path = Path(path[0]) if isinstance(path, (list, tuple)) else Path(path)
    frame, headers = _read_table_with_detected_header(erp_path, sheet_name=None)
    part_col = _find_column(headers, "part_number")
    description_col = (
        _find_column(headers, "description") or _find_column(headers, "value")
    )
    mounting_col = _find_column(headers, "mounting")
    reference_col = _find_column(headers, "reference")
    material_type_col = _find_column(headers, "material_type")
    qty_col = _find_column(headers, "qty")
    package_col = _find_column(headers, "package")
    manufacturer_col = _find_column(headers, "manufacturer")
    if part_col is None:
        raise ValueError(f"{erp_path.name}: Part Number / ürün kodu sütunu bulunamadı.")
    if description_col is None:
        raise ValueError(f"{erp_path.name}: Description / açıklama sütunu bulunamadı.")

    records: list[PartRecord] = []
    for _, row in frame.iterrows():
        code = _cell_text(row.get(part_col))
        if not code:
            continue
        records.append(PartRecord(
            part_number=code,
            description=_cell_text(row.get(description_col)),
            mounting=_cell_text(row.get(mounting_col)) if mounting_col else "",
            reference=_cell_text(row.get(reference_col)) if reference_col else "",
            material_type=_cell_text(row.get(material_type_col)) if material_type_col else "",
            qty=_cell_text(row.get(qty_col)) if qty_col else "",
            package=_cell_text(row.get(package_col)) if package_col else "",
            manufacturer=_cell_text(row.get(manufacturer_col)) if manufacturer_col else "",
            source_file=erp_path.name,
            source_sheet="",
        ))
    if not records:
        raise ValueError(f"{erp_path.name}: başlık bulundu ancak ürün kodu içeren satır yok.")
    return records


KNOWN_MANUFACTURERS: tuple[tuple[str, str, str], ...] = (
    ("samsung", "Samsung", "SS"),
    ("yageo", "Yageo", "YG"),
    ("murata", "Murata", "MT"),
    ("tdk", "TDK", "TD"),
    ("kemet", "Kemet", "KM"),
    ("royal ohm", "Royal Ohm", "RO"),
    ("royal", "Royal Ohm", "RO"),
    ("vishay", "Vishay", "VH"),
    ("panasonic", "Panasonic", "PS"),
    ("avx", "AVX", "AV"),
    ("bourns", "Bourns", "BO"),
    ("diodes inc", "Diodes Inc", "DI"),
    ("diodes incorporated", "Diodes Inc", "DI"),
    ("mdd", "MDD", "DD"),
    ("stmicroelectronics", "ST", "ST"),
    ("stmicro", "ST", "ST"),
    ("texas instruments", "TI", "TI"),
    ("on semiconductor", "On Semi", "ON"),
    ("onsemi", "On Semi", "ON"),
    ("infineon", "Infineon", "IN"),
    ("microchip", "Microchip", "MC"),
    ("nexperia", "Nexperia", "NP"),
    ("littelfuse", "Littelfuse", "LT"),
    ("rohm", "Rohm", "ROH"),
    ("taiwo yuden", "Taiyo Yuden", "TY"),
    ("taiyo", "Taiyo Yuden", "TY"),
    ("wuerth", "Wurth", "WU"),
    ("würth", "Wurth", "WU"),
    ("kls", "Kls", "KL"),
    ("molex", "Molex", "MO"),
    ("phoenix contact", "Phoenix", "PNX"),
    ("ttaf", "Ttaf", "TT"),
    ("jst", "JST", "JST"),
    ("amphenol", "Amphenol", "AM"),
    ("te connectivity", "TE Connectivity", "TC"),
    ("tyco electronics", "TE Connectivity", "TC"),
    ("tyco", "TE Connectivity", "TC"),
    ("fusheng", "FUET", "FT"),
    ("fuet", "FUET", "FT"),
    ("huaneng", "Huaneng", "HUA"),
    ("huawha", "Huaneng", "HUA"),
    ("hottech", "Hottech", "HO"),
    ("cui devices", "Same Sky", "SKY"),
    ("same sky", "Same Sky", "SKY"),
    ("cui", "Same Sky", "SKY"),
    ("yixinzy", "YIXINZY", "YIX"),
    ("epcos", "Epcos", "EP"),
    ("multicomp pro", "Multicomp Pro", "MCP"),
    ("multicomp", "Multicomp Pro", "MCP"),
    ("amtek", "Amtek", "AH"),
    ("wago", "Wago", "WA"),
    ("harting", "Harting", "HR"),
    ("hirose", "Hirose", "HIR"),
    ("samtec", "Samtec", "ZZ3"),
    ("raltron", "Raltron", "RA"),
    ("foryard", "Foryard", "FO"),
    ("panjit semiconductor", "Panjit", "PA"),
    ("panjit", "Panjit", "PA"),
    ("ezurio", "Ezurio", "EZU"),
    ("laird connectivity", "Ezurio", "EZU"),
    ("laird", "Ezurio", "EZU"),
    ("simsukian", "SIMSUKIAN", "SMK"),
    ("connfly", "Connfly", "CF"),
    ("eemb", "Eemb", "EM"),
    ("stocko", "Stocko", "SC"),
    ("gaosong", "Gaosong", "GA"),
    ("hirose electric", "Hirose Electric", "HRE"),
    ("gct", "GCT", "ZZ6"),
    ("xunpu", "XUNPU", "XUN"),
    ("iriso", "IRISO", "IR"),
    ("tianli", "Tianli", "TL"),
    ("oupiin", "Oupiin", "OP"),
    ("on shore", "On Shore", "MIK"),
    ("ismak", "İsmak", "ISM"),
    ("i\u0307smak", "İsmak", "ISM"),
    ("core", "CORE", "COR"),
    ("dtsis", "DTSİS", "DTS"),
    ("dts", "DTSİS", "DTS"),
    ("mikrodev", "Mikrodev", "MIK"),
    ("enervon", "Enervon", "ENV"),
    ("isipark", "Isipark", "ISP"),
)

_CORPORATE_SUFFIX_RE = re.compile(
    r"\b(?:Inc\.?|Co\.?,\s*Ltd\.?|Co\.?\s*Ltd\.?|Ltd\.?|Corp\.?|Corporation|Cor\.?|Electronics?|Elec\.?|Elektronik|Semiconductor|Semi|Technologies|Technology|Connectors?|Components?|LLC|GmbH|ALT\s*\d+|ALT\d+)\b",
    flags=re.IGNORECASE,
)

_DISALLOWED_MFR_TERMS = {
    "diode", "diodes", "resistor", "resistors", "capacitor", "capacitors",
    "connector", "connectors", "inductor", "inductors", "relay", "relays",
    "switch", "switches", "fuse", "fuses", "transformer", "transformers",
    "taiwan", "china", "japan", "korea", "germany", "usa", "turkey", "europe", "asia",
    "on", "it", "at", "in", "to", "by", "of", "or", "as", "is", "an", "am",
    "if", "so", "no", "we", "the", "and", "top",
}


def _clean_mfr_term(term: str) -> str:
    cleaned = re.sub(r"[\s,.-]+$", "", term).strip()
    cleaned = re.sub(r"^[\s,.-]+", "", cleaned).strip()
    return re.sub(r"\s+", " ", cleaned)


def _build_manufacturer_registry(parsed: list[RuleEntry]) -> tuple[tuple[str, str, str], ...]:
    """Build dynamic universal manufacturer registry from rules.xlsx merged with curated aliases.

    Returns tuple of (search_term, display_name, erp_code) sorted by search_term length descending.
    """
    registry: dict[str, tuple[str, str, str]] = {}

    # 1. Curated aliases from KNOWN_MANUFACTURERS take precedence
    for kw, name, code in KNOWN_MANUFACTURERS:
        term_clean = kw.lower().strip()
        if term_clean and term_clean not in registry:
            registry[term_clean] = (term_clean, name, code)

    # 2. Extract dynamic entries from rules.xlsx
    for entry in parsed:
        if entry.attribute_name in ("Q", "Manufacturer") and entry.value and entry.value_description:
            code = entry.value.strip()
            raw_desc = entry.value_description.strip().rstrip(",").strip()
            if not code or not raw_desc:
                continue

            display_name = re.sub(r"\s+", " ", raw_desc)
            raw_lower = _clean_mfr_term(display_name.lower())
            no_parens = re.sub(r"\([^)]*\)", "", display_name).strip()
            stripped = _CORPORATE_SUFFIX_RE.sub("", no_parens).strip()
            stripped_lower = _clean_mfr_term(stripped.lower())

            for term in (raw_lower, stripped_lower):
                if not term or len(term) < 3:
                    continue
                if term in _DISALLOWED_MFR_TERMS:
                    continue
                if code.upper() == "DI" and term in ("diode", "diodes"):
                    continue
                if term not in registry:
                    registry[term] = (term, display_name, code)
                elif "alt" in registry[term][1].lower() and "alt" not in display_name.lower():
                    registry[term] = (term, display_name, code)

    # Sort manufacturer match entries by search term length descending
    return tuple(sorted(registry.values(), key=lambda item: len(item[0]), reverse=True))


def load_rulebook(path: Path) -> RuleBook:
    """Read every populated rule field from the combined materials dictionary."""
    with pd.ExcelFile(path) as workbook:
        sheet = _find_rules_sheet(workbook.sheet_names)
        raw = pd.read_excel(workbook, sheet_name=sheet, header=None, dtype=object)
    header_index: int | None = None
    required_columns = {
        "material_type": {"malzemetipi"},
        "attribute_code": {"ozellikattribute", "attribute"},
        "attribute_name": {"ozellikaciklamasi", "attributeaciklamasi"},
        "value": {"value"},
        "value_description": {"valueaciklamasi"},
        "language": {"langu", "language"},
        "sort_order": {"siralamasortno", "sortno"},
        "prefix": {"prefix"},
        "suffix": {"suffix"},
    }
    column_indexes: dict[str, int] = {}
    # Sözlükte başlık satırı sabit satırda olmayabileceği için ilk bölüm taranır.
    for row_index in range(min(_MAX_SHEET_SCAN_ROWS, len(raw))):
        row = [_header_key(value) for value in raw.iloc[row_index].tolist()]
        detected = {
            field: next((i for i, name in enumerate(row) if name in aliases), None)
            for field, aliases in required_columns.items()
        }
        if all(index is not None for index in detected.values()):
            header_index = row_index
            column_indexes = {field: int(index) for field, index in detected.items()}
            break
    if header_index is None:
        raise ValueError(
            f"Kurallar dosyasının '{sheet}' sayfasında 9 kural sütununun tamamı bulunamadı."
        )

    parsed: list[RuleEntry] = []
    for _, row in raw.iloc[header_index + 1:].iterrows():
        values = {field: _cell_text(row.iloc[index]) for field, index in column_indexes.items()}
        if not any(values.values()):
            continue
        parsed.append(RuleEntry(**values))

    grouped: dict[str, dict[str, list[RuleEntry]]] = defaultdict(lambda: defaultdict(list))
    by_attribute: dict[str, list[RuleEntry]] = defaultdict(list)
    for entry in parsed:
        if entry.material_type and entry.attribute_name:
            attribute = _normalize_attribute(entry.attribute_name)
            grouped[_normalize_label(entry.material_type)][attribute].append(entry)
            by_attribute[attribute].append(entry)
    indexed = {
        material_type: {
            attribute: tuple(entries)
            for attribute, entries in attributes.items()
        }
        for material_type, attributes in grouped.items()
    }
    all_attributes = {name: tuple(entries) for name, entries in by_attribute.items()}
    return RuleBook(tuple(parsed), indexed, all_attributes, _build_manufacturer_registry(parsed))


def validate_erp_rules(record: PartRecord, rulebook: RuleBook) -> list[str]:
    """Validate ERP mounting and recognized rule values for the inferred material type."""
    issues: list[str] = []
    material_type = _infer_material_type(record, rulebook)
    if material_type is None:
        issues.append("Malzeme tipi belirlenemedi; tipe bağlı sözlük kuralları doğrulanamadı")
        return issues

    rules = rulebook.by_type_attribute[material_type]
    mounting_entries = rules.get("mounting", ())
    declared = _resolve_rule_entry(record.mounting, mounting_entries)
    description_mount = _resolve_rule_entry(record.description, mounting_entries)
    if record.mounting.strip() and declared is None:
        issues.append(f"ERP TH/SMD değeri {material_type} sözlüğünde tanınmıyor")
    if declared and description_mount and _normalize_label(declared) != _normalize_label(description_mount):
        issues.append("TH/SMD ile Description montaj tipi çelişkisi")
    issues.extend(_validate_description_values(record.description, rules, rulebook, material_type))
    return issues


def write_comparison_report(rows: list[dict[str, str]], output_path: Path) -> None:
    """Write comparison report containing all components with severity-based coloring."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Karşılaştırma Raporu"
    sheet.append(OUTPUT_HEADERS)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="173B32")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    red_fill = PatternFill("solid", fgColor="FCE8E6")     # Yanlış bilgi / hata için açık kırmızı
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")  # Standart uyarı için açık sarı
    blue_fill = PatternFill("solid", fgColor="D9E1F2")    # Normalde girilmemesi gereken / montaj dışı satırlar için açık mavi

    for row in rows:
        sheet.append([row.get(header, "") for header in OUTPUT_HEADERS])
        current_row = sheet[sheet.max_row]
        durum = row.get("Durum", "UYGUN")
        issues_text = row.get("Hata/Uyarı Durumu", "")
        row_fill = None
        if durum == "HATA":
            row_fill = red_fill
        elif durum == "UYARI":
            if any(k in issues_text for k in ("Montaj Dışı", "ERP'ye Girilmez", "ERP Sarf Malzemesi", "Ham PCB", "Dummy")):
                row_fill = blue_fill
            else:
                row_fill = yellow_fill

        for cell in current_row:
            if row_fill:
                cell.fill = row_fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = (12, 28, 28, 14, 14, 32, 32, 16, 16, 32, 28, 28, 44)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 28
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()


def _unique_records(records: list[PartRecord]) -> dict[str, PartRecord]:
    grouped: dict[str, list[PartRecord]] = defaultdict(list)
    no_code_count = 0
    dnp_statuses = {"NM", "N.M.", "NOT MOUNTED", "DNP", "NO POPULATE", "NC", "DO NOT POPULATE"}
    for record in records:
        if not record.part_number.strip():
            no_code_count += 1
            grouped[f"__NO_CODE_{no_code_count}__"].append(record)
        else:
            is_nm = (record.status or "").strip().upper() in dnp_statuses or "NOT MOUNTED" in (record.status or "").upper()
            norm = _normalize_code(record.part_number)
            if is_nm:
                grouped[f"{norm}__NM"].append(record)
            else:
                grouped[norm].append(record)
    result: dict[str, PartRecord] = {}
    for normalized, group in grouped.items():
        descriptions = list(dict.fromkeys(item.description for item in group if item.description))
        mountings = list(dict.fromkeys(item.mounting for item in group if item.mounting))
        references = list(dict.fromkeys(item.reference for item in group if item.reference))
        material_types = list(dict.fromkeys(item.material_type for item in group if item.material_type))
        packages = list(dict.fromkeys(item.package for item in group if item.package))
        manufacturers = list(dict.fromkeys(item.manufacturer for item in group if item.manufacturer))

        total_qty = 0
        has_qty = False
        non_numeric_qtys = []
        for item in group:
            if item.qty:
                has_qty = True
                try:
                    total_qty += int(float(item.qty))
                except ValueError:
                    non_numeric_qtys.append(item.qty)
        if has_qty:
            if non_numeric_qtys:
                qty_str = ", ".join(non_numeric_qtys)
            else:
                qty_str = str(total_qty)
        else:
            qty_str = ""

        is_duplicate = False
        if not normalized.startswith("__NO_CODE_") and len(group) > 1:
            has_any_ref = any(bool(item.reference.strip()) for item in group)
            if not has_any_ref:
                sources = {(item.source_file, item.source_sheet) for item in group}
                is_duplicate = len(sources) < len(group)
            else:
                all_refs = [_split_refs(item.reference) for item in group if item.reference]
                for i, ra in enumerate(all_refs):
                    for rb in all_refs[i + 1 :]:
                        if ra and rb and (ra & rb):
                            is_duplicate = True
                            break
                    if is_duplicate:
                        break

        # Source breakdown and combined references
        source_counts: dict[str, int] = defaultdict(int)
        for item in group:
            if not item.source_file:
                continue
            src_label = f"{item.source_file} ({item.source_sheet})" if item.source_sheet else item.source_file
            try:
                source_counts[src_label] += int(float(item.qty)) if item.qty else 1
            except (ValueError, TypeError):
                source_counts[src_label] += 1

        if len(source_counts) > 1:
            source_breakdown = ", ".join(f"{src}: {cnt} adet" for src, cnt in source_counts.items())
        elif len(source_counts) == 1:
            source_breakdown = list(source_counts.keys())[0]
        else:
            source_breakdown = ""

        all_refs_list: list[str] = []
        for item in group:
            for r in _split_refs(item.reference):
                if r and r not in all_refs_list:
                    all_refs_list.append(r)
        combined_ref = ", ".join(all_refs_list) if all_refs_list else " | ".join(references)

        is_unmapped = any(item.unmapped_prefix for item in group)

        merged_all_codes = tuple(dict.fromkeys(c for item in group for c in item.all_codes if c))
        merged_code_labels = {k: v for item in group for k, v in item.code_labels.items()}
        merged_status = " | ".join(dict.fromkeys(item.status for item in group if item.status))

        result[normalized] = PartRecord(
            part_number=group[0].part_number,
            description=" | ".join(descriptions),
            mounting=" | ".join(mountings),
            reference=combined_ref,
            material_type=material_types[0] if len(material_types) == 1 else "",
            qty=qty_str,
            manufacturer=" | ".join(manufacturers),
            package=" | ".join(packages),
            duplicate=is_duplicate,
            source_file=", ".join(dict.fromkeys(item.source_file for item in group if item.source_file)),
            source_sheet=", ".join(dict.fromkeys(item.source_sheet for item in group if item.source_sheet)),
            source_breakdown=source_breakdown,
            unmapped_prefix=is_unmapped,
            all_codes=merged_all_codes,
            code_labels=merged_code_labels,
            status=merged_status,
        )
    return result


def _read_table_with_detected_header(path: Path, sheet_name: str | None) -> tuple[pd.DataFrame, dict[str, str]]:
    with pd.ExcelFile(path) as book:
        selected_sheet = sheet_name or book.sheet_names[0]
        raw = pd.read_excel(book, sheet_name=selected_sheet, header=None, dtype=object)
        for row_index in range(min(_MAX_SHEET_SCAN_ROWS, len(raw))):
            headers = [_cell_text(value) for value in raw.iloc[row_index].tolist()]
            keys = {_header_key(header) for header in headers}
            if any(_header_field(header) == "part_number" for header in keys):
                frame = pd.read_excel(book, sheet_name=selected_sheet, header=row_index, dtype=object)
                resolved: dict[str, str] = {}
                for column in frame.columns:
                    role = _header_field(_cell_text(column))
                    if role and role not in resolved:
                        resolved[role] = str(column)
                return frame, resolved
    raise ValueError(f"{path.name}: ilk {_MAX_SHEET_SCAN_ROWS} satırda ürün kodu başlığı bulunamadı.")


def _find_rules_sheet(sheet_names: list[str]) -> str:
    for name in sheet_names:
        if _header_key(name) == "birlestirilmis":
            return name
    for name in sheet_names:
        if "birlestirilmis" in _header_key(name):
            return name
    raise ValueError("Kurallar dosyasında 'BİRLEŞTİRİLMİŞ' sayfası bulunamadı.")


def _find_column(headers: dict[str, str], role: str) -> str | None:
    return headers.get(role)


def _header_field(header: str) -> str | None:
    key = _header_key(header)
    if key in {
        "partnumber", "manufacturerpartnumber", "mpn", "productcode", "urunkodu",
        "malzemekodu", "stokkodu", "ureticiparcakodu", "ureticikodu", "ordercode",
        "orderingcode", "sipariskodu", "tedarikcikodu", "libref", "libraryreference",
        "pn", "mfrpn", "mfrpartno", "mfrpartnumber", "partno", "itemcode", "catalogno",
    }:
        return "part_number"
    if key in {
        "part", "parts", "reference", "references", "ref", "refs", "refdes",
        "designator", "designators", "parca", "referans", "referanslar", "devreelemani",
        "sematikreferans", "locations", "location", "position", "positions", "eleman", "etiket",
    }:
        return "reference"
    if key in {"malzemetipi", "materialtype", "componenttype", "parttype"}:
        return "material_type"
    if key in {
        "value", "customerdescription", "musterivalue", "deger", "val", "comment", "comments",
        "device", "tanim", "stokadi", "malzemeadi", "elemandegeri", "partname", "devicedescription",
    }:
        return "value"
    if key in {
        "description", "erpdescription", "aciklama", "aciklama1", "urunaciklamasi",
        "desc", "spec", "specification", "details",
    }:
        return "description"
    if key in {"thsmd", "mounting", "mountingtype", "montajtipi", "montaj", "montajsekli", "dizgitipi"}:
        return "mounting"
    if key in {
        "qty", "quantity", "adet", "miktar", "sayi", "count", "totalqty", "birimadet",
        "toplamadet", "parcaadedi", "amount", "pcs", "pieces", "sayisi", "adedi", "qtyboard",
    }:
        return "qty"
    if key in {
        "package", "footprint", "pcbfootprint", "packagecase", "case", "kilif", "paket",
        "size", "kasa", "boy", "govde",
    }:
        return "package"
    if key in {"manufacturer", "manufacture", "maker", "brand", "uretici", "marka", "vendor", "supplier"}:
        return "manufacturer"
    if key in {
        "status", "durum", "fitted", "placement", "variant", "dnp", "np", "dnpnp",
        "montajdurumu", "dizgidurumu", "dizim", "dizimdurumu", "populate", "population",
    }:
        return "status"
    return None


def _header_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", _cell_text(value).casefold())
    return "".join(char for char in text if char.isalnum())


def _normalize_code(value: str) -> str:
    return normalize_code_key(value)


def _normalize_description(value: str) -> str:
    return " ".join(value.split()).casefold()


def _normalize_label(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.casefold())
    return "".join(character for character in normalized if character.isalnum())


def _normalize_attribute(value: str) -> str:
    label = _normalize_label(value)
    if label == "mounting":
        return "mounting"
    if label.startswith("package"):
        return "package"
    if label.startswith("valu"):
        return "value"
    if label.startswith("serial"):
        return "serial"
    return label


def _cell_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _resolve_rule_entry(text: str, entries: tuple[RuleEntry, ...]) -> str | None:
    """Resolve a display value or value code using only the selected type's rules."""
    normalized = _normalize_description(text).strip(" ,;:/")
    if not normalized:
        return None
    matches: list[RuleEntry] = []
    for entry in entries:
        for alias in (entry.value, entry.value_description):
            candidate = _normalize_description(alias).strip(" ,;:/")
            if candidate and normalized == candidate:
                matches.append(entry)
                break
    if len(matches) != 1:
        # Description alanı uzun metin olduğunda yalnız açıkça bulunan sözlük
        # açıklamasını kabul et; kısa kodları metin içinde arama.
        matches = [
            entry for entry in entries
            if entry.value_description
            and len(_normalize_description(entry.value_description).strip(" ,;:/")) >= 2
            and _contains_phrase(text, _entry_text(entry))
        ]
    if not matches:
        return None
    descriptions = list(dict.fromkeys(entry.value_description for entry in matches if entry.value_description))
    return descriptions[0] if len(descriptions) == 1 else None


def _entry_text(entry: RuleEntry) -> str:
    description = entry.value_description.strip().strip(" ,;:/")
    if _normalize_attribute(entry.attribute_name) == "tolerance":
        try:
            numeric_value = float(description)
        except ValueError:
            numeric_value = None
        if numeric_value is not None and 0 < numeric_value < 1:
            description = f"{numeric_value * 100:g}%"
    return description


def _contains_phrase(text: str, phrase: str) -> bool:
    """Match a normalized phrase without allowing it inside a larger word/value."""
    normalized_text = _normalize_description(text)
    normalized_phrase = _normalize_description(phrase).strip(" ,;:/")
    if not normalized_phrase:
        return False
    start = 0
    while (position := normalized_text.find(normalized_phrase, start)) >= 0:
        end = position + len(normalized_phrase)
        left_ok = position == 0 or not (normalized_text[position - 1].isalnum() or normalized_text[position - 1] in "-_")
        right_ok = end == len(normalized_text) or not (normalized_text[end].isalnum() or normalized_text[end] in "-_")
        if left_ok and right_ok:
            return True
        start = position + 1
    return False


def _infer_material_type(record: PartRecord, rulebook: RuleBook) -> str | None:
    explicit = _normalize_label(record.material_type)
    if explicit in rulebook.by_type_attribute:
        return explicit

    ref_clean = _normalize_label(record.reference)
    if ref_clean in rulebook.by_type_attribute:
        return ref_clean

    desc_clean = record.description.casefold()
    ref_lower = record.reference.casefold()

    if "lehim" in ref_lower or "lehim" in desc_clean:
        if "sarf" in rulebook.by_type_attribute:
            return "sarf"

    if "stm32" in desc_clean or "stm32" in record.part_number.casefold():
        if "mcu" in rulebook.by_type_attribute:
            return "mcu"

    if "voltage regulator" in desc_clean or "regulator" in desc_clean:
        if "icr" in rulebook.by_type_attribute:
            return "icr"

    if "humidity sensor" in desc_clean or "shtc3" in desc_clean:
        if "isl" in rulebook.by_type_attribute:
            return "isl"

    if "makaron" in desc_clean or "makaron" in ref_lower:
        if "sarf" in rulebook.by_type_attribute:
            return "sarf"

    if "display board" in desc_clean or "displayboard" in desc_clean or "display board" in ref_lower:
        if "ymml" in rulebook.by_type_attribute:
            return "ymml"
        if "pcb" in rulebook.by_type_attribute:
            return "pcb"

    if "coil:" in desc_clean or "relay" in desc_clean or "röle" in desc_clean or "role" in desc_clean:
        if "rel" in rulebook.by_type_attribute:
            return "rel"

    if "opto" in desc_clean or "ltv-" in desc_clean or "isp817" in desc_clean or "pc817" in desc_clean:
        if "opt" in rulebook.by_type_attribute:
            return "opt"

    if "konnektör" in desc_clean or "connector" in desc_clean or "terminal block" in desc_clean or "modular jack" in desc_clean:
        if "con" in rulebook.by_type_attribute:
            return "con"

    if "ac/dc converter" in desc_clean or "ac/dc" in desc_clean or "tny" in desc_clean or "tny" in record.part_number.casefold():
        if "icr" in rulebook.by_type_attribute:
            return "icr"

    if "bridge" in desc_clean or "köprü" in desc_clean or "sdip-4l" in desc_clean:
        if "dio" in rulebook.by_type_attribute:
            return "dio"

    # Part/Ref designator prefixes give a deterministic class for common BOM items.
    reference_map = {
        "c": "cap", "r": "res", "l": "ind", "d": "dio", "fb": "fer",
        "f": "fus", "j": "con", "led": "led", "sw": "swt", "x": "cry",
        "y": "cry", "ls": "spk", "spk": "spk", "bz": "buz", "mic": "mic",
        "pcb": "pcb", "vr": "var", "rv": "var", "var": "var",
        "ds": "dsp", "dsp": "dsp", "rt": "res", "ts": "spr",
        "q": "tra", "ld": "led", "s": "swt",
        "cbl": "cbl", "cable": "cbl",
        "rp": "res", "tft": "dsp", "spc": "mknk", "bant": "sarf",
        "p": "con", "rl": "rel", "k": "rel", "hs": "con", "mk": "sarf",
        "br": "dio", "csn": "cap", "rsn": "res",
    }
    inferred: set[str] = set()
    for segment in re.split(r"[,;|]", record.reference.upper()):
        match = re.match(r"\s*([A-Za-z0-9_]+)", segment)
        if match:
            raw_token = match.group(1).casefold()
            # Try exact prefix match first
            candidate = reference_map.get(raw_token)
            if not candidate:
                # Strip trailing digits or underscores
                alpha_only = re.sub(r"[\d_].*$", "", raw_token)
                candidate = reference_map.get(alpha_only)
            if not candidate:
                if raw_token.startswith("r") and not raw_token.startswith(("rl", "rv", "rp", "rt")):
                    candidate = "res"
                elif raw_token.startswith("c") and not raw_token.startswith(("cbl",)):
                    candidate = "cap"
                elif raw_token.startswith("d") or raw_token.startswith("br"):
                    candidate = "dio"
                elif raw_token.startswith("l"):
                    candidate = "ind"
                elif raw_token.startswith("j") or raw_token.startswith("p"):
                    candidate = "con"
            if candidate and candidate in rulebook.by_type_attribute:
                inferred.add(candidate)
    if len(inferred) == 1:
        return next(iter(inferred))

    # Ambiguous reference designators (for example Q3) are resolved only by a
    # unique type-specific Value/Type term present in the ERP description.
    scores: dict[str, int] = defaultdict(int)
    for material_type, attributes in rulebook.by_type_attribute.items():
        for attribute, entries in attributes.items():
            if attribute not in {"value", "type"}:
                continue
            for entry in entries:
                phrase = _entry_text(entry)
                if len(phrase) >= 3 and (
                    _contains_phrase(record.description, phrase)
                    or _contains_phrase(record.part_number, phrase)
                ):
                    scores[material_type] += len(phrase)
    if not scores:
        return None
    best_score = max(scores.values())
    winners = [material_type for material_type, score in scores.items() if score == best_score]
    return winners[0] if len(winners) == 1 else None


def _validate_description_values(
    description: str,
    typed_rules: dict[str, tuple[RuleEntry, ...]],
    rulebook: RuleBook,
    material_type: str,
) -> list[str]:
    """Flag recognized dictionary terms that are not allowed for this material type."""
    findings: list[str] = []
    matched_for_order: list[tuple[int, float, str]] = []
    valid_type_phrases = {
        _normalize_description(_entry_text(entry))
        for entries in typed_rules.values()
        for entry in entries
        if _entry_text(entry)
    }

    # Only validate attributes that this material type defines (excluding mounting and manufacturer Q)
    for attribute, allowed in typed_rules.items():
        if attribute in {"mounting", "q"}:
            continue
        all_entries = rulebook.by_attribute.get(attribute, ())
        detected = [
            entry for entry in all_entries
            if len(_entry_text(entry)) >= 3 and _contains_phrase(description, _entry_text(entry))
        ]
        if not detected:
            continue
        allowed_phrases = {_normalize_description(_entry_text(entry)) for entry in allowed}
        invalid = [
            entry for entry in detected
            if _normalize_description(_entry_text(entry)) not in allowed_phrases
        ]
        if invalid:
            invalid = [
                entry for entry in invalid
                if _normalize_description(_entry_text(entry)) not in valid_type_phrases
            ]
        if invalid:
            example = invalid[0]
            sample = list(dict.fromkeys(_entry_text(entry) for entry in invalid))[:3]
            findings.append(
                f"Description sözlük uyumsuzluğu ({material_type}/"
                f"{example.attribute_name}, attribute {example.attribute_code}): "
                + ", ".join(sample)
            )

    # Sıralama kontrolü: yalnızca bu malzeme tipinin kendi attribute'ları
    for attribute, entries in typed_rules.items():
        if attribute == "mounting":
            continue
        for entry in entries:
            phrase = _entry_text(entry)
            if len(phrase) < 3 or not _contains_phrase(description, phrase):
                continue
            try:
                sort_order = float(entry.sort_order)
            except ValueError:
                continue
            position = _normalize_description(description).find(_normalize_description(phrase))
            if position >= 0:
                matched_for_order.append((position, sort_order, attribute))

    matched_for_order.sort()
    sort_orders = [item[1] for item in matched_for_order]
    if any(later < earlier for earlier, later in zip(sort_orders, sort_orders[1:])):
        findings.append("Description içindeki özellik sırası sözlük SORTNO sırasıyla uyuşmuyor")
    return findings
