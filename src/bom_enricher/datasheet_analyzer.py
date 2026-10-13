"""Datasheet PDF analysis and ERP property mapping engine."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
import io
from pathlib import Path
import re
import unicodedata
from typing import Any, BinaryIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader

from bom_enricher.comparison import RuleBook, RuleEntry, _normalize_label, _normalize_attribute, _entry_text, _contains_phrase
from bom_enricher.semantic_parser import (
    _extract_capacitance,
    _extract_resistance,
    _extract_inductance,
    _extract_voltage,
    _extract_tolerance,
    _extract_mounting,
    _extract_package,
    normalize_mounting,
    _format_cap,
    _format_res,
)


@dataclass
class PropertyMapping:
    """Mapping of a single electrical/physical property to an ERP dictionary attribute."""
    attribute_name: str
    sort_no: float
    extracted_value: str
    matched_erp_value: str
    erp_code: str
    prefix: str = ""
    status: str = "TAM EŞLEŞTİ"  # "TAM EŞLEŞTİ", "YAKIN EŞLEŞTİ", "SÖZLÜKTE YOK"


@dataclass
class DatasheetAnalysisResult:
    """Full analysis result for a single component datasheet."""
    filename: str
    part_number: str
    material_type: str                  # e.g. "CAP"
    material_type_name: str             # e.g. "Kapasitör (CAP)"
    inferred_mounting: str              # e.g. "SMD"
    confidence: float                   # 0.0 - 1.0
    suggested_description: str          # e.g. "SMD 0805 Package, 100nF, %10, 50V, X7R, Samsung"
    suggested_erp_code: str             # e.g. "CAPSS31003KD4X7R-SA"
    properties: list[PropertyMapping] = field(default_factory=list)
    raw_specs: dict[str, Any] = field(default_factory=dict)
    page_count: int = 1
    analyzed_pages: list[int] = field(default_factory=list)
    summary_text: str = ""
    error: str | None = None
    is_catalog: bool = False
    detected_part_numbers: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "filename": self.filename,
            "part_number": self.part_number,
            "material_type": self.material_type,
            "material_type_name": self.material_type_name,
            "inferred_mounting": self.inferred_mounting,
            "confidence": round(self.confidence, 2),
            "suggested_description": self.suggested_description,
            "suggested_erp_code": self.suggested_erp_code,
            "page_count": self.page_count,
            "analyzed_pages": self.analyzed_pages,
            "summary_text": self.summary_text,
            "error": self.error,
            "is_catalog": self.is_catalog,
            "detected_part_numbers": self.detected_part_numbers,
            "properties": [
                {
                    "attribute_name": p.attribute_name,
                    "sort_no": p.sort_no,
                    "extracted_value": p.extracted_value,
                    "matched_erp_value": p.matched_erp_value,
                    "erp_code": p.erp_code,
                    "prefix": p.prefix,
                    "status": p.status,
                }
                for p in self.properties
            ],
            "raw_specs": {
                k: str(v) if isinstance(v, Decimal) else v
                for k, v in self.raw_specs.items()
            },
        }


MATERIAL_TYPE_NAMES = {
    "CAP": "Kapasitör (CAP)",
    "RES": "Direnç (RES)",
    "DIO": "Diyot (DIO)",
    "FET": "MOSFET / FET",
    "TRA": "Transistör (TRA)",
    "ICR": "Voltaj Regülatörü / LDO (ICR)",
    "ICS": "SMPS / Güç Entegresi (ICS)",
    "MCU": "Mikrodenetleyici (MCU)",
    "IND": "Bobin / İndüktör (IND)",
    "LED": "LED",
    "CON": "Konnektör (CON)",
    "OPT": "Optokuplör (OPT)",
    "REL": "Röle (REL)",
    "FUS": "Sigorta (FUS)",
    "VAR": "Varistör (VAR)",
    "TER": "Termistör (TER)",
    "CRY": "Kristal / Osilatör (CRY)",
    "SWT": "Anahtar / Switch (SWT)",
    "FER": "Ferrit Boncuk (FER)",
}


KNOWN_MANUFACTURERS = [
    ("samsung", "Samsung", "SA"),
    ("yageo", "Yageo", "YA"),
    ("murata", "Murata", "MU"),
    ("tdk", "TDK", "TD"),
    ("kemet", "Kemet", "KE"),
    ("royal ohm", "Royal Ohm", "RO"),
    ("royal", "Royal Ohm", "RO"),
    ("vishay", "Vishay", "VI"),
    ("panasonic", "Panasonic", "PA"),
    ("avx", "AVX", "AV"),
    ("bourns", "Bourns", "BO"),
    ("diodes inc", "Diodes Inc", "DI"),
    ("diodes incorporated", "Diodes Inc", "DI"),
    ("mdd", "MDD", "MD"),
    ("stmicroelectronics", "ST", "ST"),
    ("stmicro", "ST", "ST"),
    ("texas instruments", "TI", "TI"),
    ("on semiconductor", "On Semi", "ON"),
    ("onsemi", "On Semi", "ON"),
    ("infineon", "Infineon", "IN"),
    ("microchip", "Microchip", "MC"),
    ("nexperia", "Nexperia", "NX"),
    ("littelfuse", "Littelfuse", "LF"),
    ("rohm", "Rohm", "ROH"),
    ("taiwo yuden", "Taiyo Yuden", "TY"),
    ("taiyo", "Taiyo Yuden", "TY"),
    ("wuerth", "Wurth", "WE"),
    ("würth", "Wurth", "WE"),
    ("kls", "Kls", "KL"),
    ("molex", "Molex", "MO"),
    ("phoenix contact", "Phoenix", "PH"),
    ("ttaf", "Ttaf", "TT"),
    ("jst", "JST", "JST"),
    ("amphenol", "Amphenol", "AM"),
    ("te connectivity", "TE", "TE"),
    ("te", "TE", "TE"),
    ("amtek", "Amtek", "AMT"),
    ("wago", "Wago", "WAG"),
    ("harting", "Harting", "HAR"),
    ("hirose", "Hirose", "HIR"),
    ("samtec", "Samtec", "SAM"),
]


def extract_pdf_data(stream_or_path: BinaryIO | bytes | Path, part_number: str | None = None) -> tuple[str, int, list[int]]:
    """Extract informative text from a PDF, intelligently filtering long documents."""
    if isinstance(stream_or_path, (bytes, bytearray)):
        stream = io.BytesIO(stream_or_path)
    elif isinstance(stream_or_path, Path):
        stream = open(stream_or_path, "rb")
    else:
        stream = stream_or_path

    reader = PdfReader(stream)
    total_pages = len(reader.pages)
    if total_pages == 0:
        return "", 0, []

    # If small PDF (<= 8 pages), read all pages
    if total_pages <= 8:
        pages_to_read = list(range(total_pages))
    else:
        # Prioritize key pages: always pages 0 and 1
        pages_set = {0, 1}
        pn_clean = re.sub(r"[^A-Za-z0-9]", "", part_number or "").casefold()
        keywords = (
            "ordering", "part number", "specifications", "electrical", "ratings",
            "absolute maximum", "dimension", "package", "characteristics", "marking"
        )
        for i in range(total_pages):
            if len(pages_set) >= 12:
                break
            page_text = ""
            try:
                page_text = reader.pages[i].extract_text() or ""
            except Exception:
                continue
            text_lower = page_text.casefold()
            if pn_clean and pn_clean in re.sub(r"[^A-Za-z0-9]", "", text_lower):
                pages_set.add(i)
            elif any(k in text_lower for k in keywords):
                pages_set.add(i)

        pages_to_read = sorted(pages_set)[:12]

    extracted_parts = []
    analyzed_pages = []
    for idx in pages_to_read:
        try:
            page_txt = reader.pages[idx].extract_text()
            if page_txt:
                extracted_parts.append(f"--- PAGE {idx + 1} ---\n" + page_txt)
                analyzed_pages.append(idx + 1)
        except Exception:
            continue

    return "\n\n".join(extracted_parts), total_pages, analyzed_pages


def is_catalog_document_name(name: str) -> bool:
    """Return True if the name/code indicates a general catalog or document rather than a specific MPN."""
    if not name:
        return False
    clean = re.sub(r"\.pdf$", "", name, flags=re.I).strip().upper()

    # 1. Distributor document prefixes/patterns
    # e.g. PYU-RC_51_ROHS_P, PYU-RC_GROUP_51_ROHS_L, DOC-1234, DS-001, DATASHEET_V1
    if re.match(r"^(?:PYU|DOC|DS|CATALOG|DATASHEET|SPEC|MANUAL)[-_]", clean):
        return True

    # 2. Distributor numeric document IDs (e.g. 6-8 digit PDF codes from DigiKey/Mouser)
    if re.fullmatch(r"\d{6,8}", clean):
        return True

    # 3. Contains series/catalog indicator keywords
    catalog_keywords = {
        "GROUP", "ROHS", "SERIES", "FAMILY", "CATALOG", "OVERVIEW",
        "GENERAL", "SELECTION", "GUIDE", "DATASHEET", "SPECIFICATION",
        "SUMMARY", "BROCHURE",
    }
    tokens = {t for t in re.split(r"[-_\s.]+", clean) if t}
    if tokens & catalog_keywords:
        return True

    # 4. Standalone series or package prefix without ordering code
    # e.g. "RC", "RC0805", "CC0805", "CL21", "GRM18", "0805", "0603"
    if re.fullmatch(r"(?:RC|CC|AC|CL|GRM|CGA)?\d{4}", clean):
        return True
    if clean in ("RC", "CC", "AC", "CL", "GRM", "CGA", "RESISTOR", "CAPACITOR"):
        return True

    return False


def is_automotive_qualified(part_number: str, text: str) -> bool:
    """Determine if a component is truly Automotive AEC-Q200 qualified.

    Avoids false positives from footnotes, testing standards, or cross-references
    (e.g., Yageo RC series mentioning AC series or AEC-Q200 test methods).
    """
    pn_clean = part_number.strip().upper()

    # Explicit General Purpose series prefixes -> NEVER automotive
    if pn_clean.startswith("RC") or pn_clean.startswith("PYU-RC") or pn_clean.startswith("GRM"):
        return False

    # Explicit Automotive series prefixes
    if re.match(r"^(?:AC|GCM|CGA|AT|AF)\d", pn_clean):
        return True

    text_lower = text.casefold()

    # Check title / first 1500 chars for explicit automotive grade declaration
    header_snippet = text_lower[:1500]
    auto_title_patterns = (
        "automotive grade", "automotive chip resistor", "automotive mlcc",
        "aec-q200 qualified chip", "for automotive applications"
    )
    if any(p in header_snippet for p in auto_title_patterns):
        if "not for automotive" in header_snippet or "refer to ac series" in header_snippet:
            return False
        return True

    return False


def extract_part_numbers_from_text(text: str) -> list[str]:
    """Scan text for genuine manufacturer part numbers (MPNs)."""
    if not text:
        return []
    found: list[str] = []
    seen = set()

    patterns = [
        # Yageo Resistors: RC0805JR-07300RL, RC0402FR-0710KL, AC0805JR-07300RL, etc.
        r"\b(?:RC|AC)\d{4}[A-Z][A-Z]-[0-9A-Z]{2}[0-9A-Z]+L?\b",
        # Yageo Capacitors: CC0805KRX7R9BB104, AC0603KRX7R9BB104
        r"\b(?:CC|AC)\d{4}[A-Z]{2}[A-Z0-9]+BB\d{3}\b",
        # Samsung MLCC: CL21B104KBCNNNC, CL10B103KB8NNNC, etc.
        r"\bCL\d{2}[A-Z]\d{3}[A-Z]{2}[A-Z0-9]{4,6}\b",
        # Murata MLCC: GRM188R71H104KA93D, GRM21BR71E105KA99L
        r"\bGRM\d{2}[A-Z0-9]{6,14}\b",
        # Royal Ohm Chip Resistors: 0805W8F1001T5E, 0603WAF1002T5E
        r"\b\d{4}(?:WG|WA|W8|W4|W2|1W|2W|S4|S8|SA)[A-Z][0-9A-Z]{4}[A-Z0-9]*\b",
    ]

    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            code = m.group(0).upper()
            if code not in seen:
                seen.add(code)
                found.append(code)

    return found


def classify_material_type(text: str, part_number: str = "") -> tuple[str, float, str]:
    """Determine component material type (CAP, RES, DIO, etc.) from PN and datasheet content."""
    pn = part_number.strip().upper()
    text_lower = text.casefold()

    # 1. High confidence Part Number prefixes
    if re.match(r"^CL\d{2}[A-Z]\d{3}", pn) or re.match(r"^GRM\d{2}", pn) or re.match(r"^CC\d{4}", pn) or re.match(r"^AC\d{4}[A-Z]", pn):
        return "CAP", 0.98, "Part Number standart seramik kapasitör kodlamasına uyuyor."
    if re.match(r"^(?:RC\d{4}|CRCW\d{4}|ERJ\d|WR\d{2}|PYU-RC)", pn):
        return "RES", 0.98, "Part Number standart çip direnç kodlamasına uyuyor."
    if re.match(r"^(?:BAT|BZX|1N4148|1N400|1N58|SS\d{2}|MBR|SM712|SMAJ|SMBJ)", pn):
        return "DIO", 0.98, "Part Number standart diyot serisine uyuyor."
    if re.match(r"^(?:2N7002|BSS138|BSS84|SI23\d{2}|AO34\d{2}|IRF)", pn):
        return "FET", 0.95, "Part Number MOSFET kodlamasına uyuyor."
    if re.match(r"^(?:BC84|BC85|BC807|BC817|2N3904|2N3906|MMBT)", pn):
        return "TRA", 0.95, "Part Number BJT transistör kodlamasına uyuyor."
    if re.match(r"^(?:LM1117|AMS1117|LM78|TL431|AP1117|LP29)", pn):
        return "ICR", 0.95, "Part Number regülatör/LDO kodlamasına uyuyor."
    if re.match(r"^(?:VIPER|LNK|TNY|TOP\d+|FSL)", pn):
        return "ICS", 0.95, "Part Number SMPS güç entegresi kodlamasına uyuyor."
    if re.match(r"^(?:STM32|PIC\d|ATMEGA|ESP32|ATTINY)", pn):
        return "MCU", 0.98, "Part Number mikrodenetleyici kodlamasına uyuyor."
    if re.match(r"^(?:PC817|EL817|LTV-817|MOC30)", pn):
        return "OPT", 0.98, "Part Number optokuplör kodlamasına uyuyor."
    if re.match(r"^(?:HF\d|G5Q|SRD|G2R)", pn):
        return "REL", 0.95, "Part Number röle kodlamasına uyuyor."
    if re.match(r"^(?:22-01|0022|2695|5557|JST-|PH-|XH-|VH-|MOLEX|200G|254G|200W|254W|200H|254H|200D|254D|\d{3}[A-Z]\d?-\d{2})", pn):
        return "CON", 0.98, "Part Number standart konnektör/klemens serisine uyuyor."

    # 2. Text Keyword Scoring
    scores: dict[str, float] = {
        "CAP": 0.0, "RES": 0.0, "DIO": 0.0, "FET": 0.0, "TRA": 0.0,
        "ICR": 0.0, "ICS": 0.0, "MCU": 0.0, "IND": 0.0, "LED": 0.0,
        "CON": 0.0, "OPT": 0.0, "REL": 0.0, "FUS": 0.0, "VAR": 0.0,
    }

    # Royal Ohm check (exclusive resistor manufacturer)
    if any(k in text_lower or k in pn.casefold() for k in ("royal ohm", "royalohm", "royal-ohm")):
        scores["RES"] += 30.0

    # Capacitor triggers
    if "capacitor" in text_lower or "capacitance" in text_lower or "mlcc" in text_lower:
        scores["CAP"] += 12.0
    if "dielectric" in text_lower or "x7r" in text_lower or "c0g" in text_lower or "np0" in text_lower:
        scores["CAP"] += 6.0
    if "picofarad" in text_lower or "nanofarad" in text_lower or "microfarad" in text_lower:
        scores["CAP"] += 6.0

    # Resistor triggers
    if any(w in text_lower for w in ("resistor", "resistors", "resistance", "thick film chip", "chip resistor", "metal film")):
        scores["RES"] += 15.0
    if "power rating" in text_lower or "rated power" in text_lower:
        scores["RES"] += 6.0
    if any(w in text_lower for w in ("kohm", "mohm", "ohm", "kω", "mω", "ω")):
        scores["RES"] += 8.0
    if re.search(r"\b\d+(?:\.\d+)?\s*(?:kohm|mohm|ohm|kω|mω|ω)\b", text_lower):
        scores["RES"] += 10.0
    if any(w in text_lower for w in ("1/16w", "1/10w", "1/8w", "1/4w", "1/2w", "1w", "2w")):
        scores["RES"] += 8.0

    # Connector triggers
    if any(w in text_lower for w in ("connector", "header", "housing", "crimp", "terminal block", "pin header", "wire-to-board", "board-to-board", "receptacle housing", "wafer", "crimp terminal")):
        scores["CON"] += 20.0
    if any(w in text_lower for w in ("molex", "jst", "amphenol", "te connectivity", "phoenix contact", "wago", "harting", "hirose", "samtec", "ttaf", "amtek")):
        scores["CON"] += 15.0
    if "circuits" in text_lower and any(w in text_lower for w in ("pitch", "housing", "crimp", "kk")):
        scores["CON"] += 12.0
    if re.search(r"\b\d+\s*(?:circuits|positions|ways)\b", text_lower):
        scores["CON"] += 10.0

    # Diode triggers
    if "schottky" in text_lower or "zener" in text_lower or "rectifier" in text_lower:
        scores["DIO"] += 12.0
    if "forward voltage" in text_lower and "reverse leakage" in text_lower:
        scores["DIO"] += 6.0
    if "diode" in text_lower:
        scores["DIO"] += 6.0

    # FET triggers
    if "mosfet" in text_lower or "n-channel" in text_lower or "p-channel" in text_lower or "drain-source" in text_lower:
        scores["FET"] += 12.0

    # Transistor triggers
    if "bipolar transistor" in text_lower or "npn" in text_lower or "pnp" in text_lower or "collector-emitter" in text_lower:
        scores["TRA"] += 12.0

    # Regulator triggers
    if "voltage regulator" in text_lower or "ldo" in text_lower or "low dropout" in text_lower or "linear regulator" in text_lower:
        scores["ICR"] += 12.0

    # Microcontroller triggers
    if "microcontroller" in text_lower or "arm cortex" in text_lower or "flash mcu" in text_lower:
        scores["MCU"] += 12.0

    # Inductor triggers
    if "inductor" in text_lower or "inductance" in text_lower or "choke" in text_lower:
        scores["IND"] += 12.0

    # Optocoupler triggers
    if "optocoupler" in text_lower or "optoisolator" in text_lower or "phototransistor" in text_lower:
        scores["OPT"] += 12.0

    # Relay triggers
    if "relay" in text_lower or "contact rating" in text_lower or "coil voltage" in text_lower:
        scores["REL"] += 12.0

    best_type = max(scores, key=scores.get)
    best_score = scores[best_type]

    if best_score >= 4.0:
        conf = min(0.98, 0.5 + (best_score / 30.0))
        return best_type, conf, f"Datasheet metnindeki teknik anahtar kelimelerden tespit edildi ({best_type})."

    # Contextual fallback based on part number or filename words
    pn_lower = pn.casefold()
    if any(k in pn_lower for k in ("res", "ohm", "royal")):
        return "RES", 0.70, "Dosya adı veya üretici bilgisinden direnç olarak tespit edildi."
    if any(k in pn_lower for k in ("cap", "uf", "nf", "pf", "mlcc")):
        return "CAP", 0.70, "Dosya adı veya birim bilgisinden kapasitör olarak tespit edildi."
    if any(k in pn_lower for k in ("dio", "bat", "bzx", "1n4")):
        return "DIO", 0.70, "Dosya adından diyot olarak tespit edildi."
    if any(k in pn_lower for k in ("con", "molex", "jst", "hdr", "header", "term", "crimp")):
        return "CON", 0.70, "Dosya adı veya marka bilgisinden konnektör olarak tespit edildi."

    # Look for known manufacturer in text
    if any(k in text_lower for k in ("molex", "jst", "ttaf", "amtek", "phoenix contact")):
        return "CON", 0.75, "Datasheet metnindeki konnektör üreticisinden tespit edildi (CON)."
    if any(k in text_lower for k in ("royal ohm", "royalohm")):
        return "RES", 0.85, "Datasheet metnindeki direnç üreticisinden tespit edildi (RES)."

    return "RES" if "ohm" in text_lower else ("CON" if any(w in text_lower for w in ("housing", "connector", "circuits", "molex")) else "CAP"), 0.35, "Malzeme tipi kesin belirlenemedi."


def decode_manufacturer_pn(part_number: str) -> dict[str, Any]:
    """Decode technical properties directly from well-known manufacturer part numbers."""
    pn = part_number.strip().upper()
    specs: dict[str, Any] = {}

    # 1. Samsung MLCC: e.g. CL21B104KBCNNNC
    samsung_match = re.match(r"^CL(\d{2})([A-Z])(\d{3})([A-Z])([A-Z])", pn)
    if samsung_match:
        size_code, dielec_code, cap_code, tol_code, volt_code = samsung_match.groups()
        size_map = {"03": "0201", "05": "0402", "10": "0603", "21": "0805", "31": "1206", "32": "1210"}
        dielec_map = {"A": "C0G", "B": "X7R", "X": "X5R", "F": "Y5V", "S": "X7S"}
        tol_map = {"F": "1%", "G": "2%", "J": "5%", "K": "10%", "M": "20%", "C": "±0.25pF", "D": "±0.5pF"}
        volt_map = {"R": "4V", "Q": "6.3V", "P": "10V", "O": "16V", "A": "25V", "L": "35V", "B": "50V", "E": "100V", "G": "500V"}

        specs["manufacturer"] = "Samsung"
        specs["mounting"] = "SMD"
        if size_code in size_map:
            specs["package"] = size_map[size_code]
        if dielec_code in dielec_map:
            specs["special_def"] = dielec_map[dielec_code]
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]
        if volt_code in volt_map:
            specs["voltage"] = volt_map[volt_code]

        # Calculate cap: 104 -> 10 * 10^4 pF = 100,000 pF = 100 nF
        try:
            val = int(cap_code[:2]) * (10 ** int(cap_code[2]))
            specs["capacitance_pf"] = val
            if val >= 1_000_000:
                specs["capacitance_str"] = f"{val // 1_000_000}uF"
            elif val >= 1_000:
                specs["capacitance_str"] = f"{val // 1_000}nF"
            else:
                specs["capacitance_str"] = f"{val}pF"
        except Exception:
            pass
        return specs

    # 2. Yageo MLCC: e.g. CC0805KRX7R9BB104 or AC0603KRX7R9BB104
    yageo_mlcc = re.match(r"^(?:CC|AC)(\d{4})([A-Z])([A-Z0-9]+?)([0-9A-Z])BB(\d{3})", pn)
    if yageo_mlcc:
        pkg, tol_code, dielec_raw, volt_code, cap_code = yageo_mlcc.groups()
        tol_map = {"F": "1%", "G": "2%", "J": "5%", "K": "10%", "M": "20%"}
        volt_map = {"5": "6.3V", "6": "10V", "7": "16V", "8": "25V", "9": "50V", "0": "100V", "A": "200V"}

        specs["manufacturer"] = "Yageo"
        specs["mounting"] = "SMD"
        specs["package"] = pkg
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]
        if volt_code in volt_map:
            specs["voltage"] = volt_map[volt_code]

        dielec = dielec_raw
        for d in ("X7R", "X5R", "NPO", "C0G", "Y5V", "X7S"):
            if d in dielec_raw:
                dielec = d
                break
        specs["special_def"] = dielec

        try:
            val = int(cap_code[:2]) * (10 ** int(cap_code[2]))
            specs["capacitance_pf"] = val
            if val >= 1_000_000:
                specs["capacitance_str"] = f"{val // 1_000_000}uF"
            elif val >= 1_000:
                specs["capacitance_str"] = f"{val // 1_000}nF"
            else:
                specs["capacitance_str"] = f"{val}pF"
        except Exception:
            pass
        return specs

    # 3. Murata MLCC: e.g. GRM188R71H104KA93D
    murata_match = re.match(r"^GRM(\d{2})\d([A-Z0-9]{2})([0-9A-Z]{2})(\d{3})([A-Z])", pn)
    if murata_match:
        size_code, dielec, volt_code, cap_code, tol_code = murata_match.groups()
        size_map = {"03": "0201", "15": "0402", "18": "0603", "21": "0805", "31": "1206", "32": "1210"}
        volt_map = {"0J": "6.3V", "1A": "10V", "1C": "16V", "1E": "25V", "1H": "50V", "2A": "100V"}
        tol_map = {"F": "1%", "G": "2%", "J": "5%", "K": "10%", "M": "20%"}

        specs["manufacturer"] = "Murata"
        specs["mounting"] = "SMD"
        if size_code in size_map:
            specs["package"] = size_map[size_code]
        if volt_code in volt_map:
            specs["voltage"] = volt_map[volt_code]
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]
        if "R7" in dielec:
            specs["special_def"] = "X7R"
        elif "5C" in dielec:
            specs["special_def"] = "C0G"

        try:
            val = int(cap_code[:2]) * (10 ** int(cap_code[2]))
            specs["capacitance_pf"] = val
            if val >= 1_000_000:
                specs["capacitance_str"] = f"{val // 1_000_000}uF"
            elif val >= 1_000:
                specs["capacitance_str"] = f"{val // 1_000}nF"
            else:
                specs["capacitance_str"] = f"{val}pF"
        except Exception:
            pass
        return specs

    # 4. Yageo Chip Resistor: e.g. RC0805JR-07300RL, RC0805FR-0710KL, AC0805JR-07300RL, RC0603JR-074K7L
    yageo_res = re.match(r"^(RC|AC)(\d{4})([A-Z])([A-Z])-([0-9A-Z]{2})([0-9A-Z]+)$", pn)
    if yageo_res:
        series, pkg, tol_code, pack_code, reel_code, val_code = yageo_res.groups()
        tol_map = {
            "B": "0.1%",
            "C": "0.25%",
            "D": "0.5%",
            "F": "1%",
            "G": "2%",
            "J": "5%",
            "Z": "Jumper",
        }
        specs["manufacturer"] = "Yageo"
        specs["mounting"] = "SMD"
        specs["package"] = pkg
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]

        # Series awareness: RC = strictly General; AC = Automotive AEC-Q200
        if series == "AC":
            specs["special_def"] = "Automotive AEC-Q200"
        else:
            specs["special_def"] = "General"

        specs["composition"] = "Thick Film"

        # Standard wattage by package size
        yageo_wattage_by_pkg = {
            "0075": "1/50W",
            "0100": "1/32W",
            "0201": "1/20W",
            "0402": "1/16W",
            "0603": "1/10W",
            "0805": "1/8W",
            "1206": "1/4W",
            "1210": "1/2W",
            "2010": "3/4W",
            "2512": "1W",
        }
        if pkg in yageo_wattage_by_pkg:
            specs["wattage"] = yageo_wattage_by_pkg[pkg]

        # Decode resistance value: e.g. 300RL -> 300R, 10KL -> 10k, 4K7L -> 4.7k, 0RL -> 0R
        clean_val = val_code.rstrip("L")
        if re.fullmatch(r"\d+K", clean_val):
            clean_val = clean_val[:-1] + "k"
        elif m := re.fullmatch(r"(\d+)K(\d+)", clean_val):
            clean_val = f"{m.group(1)}.{m.group(2)}k"
        elif m := re.fullmatch(r"(\d+)M(\d+)", clean_val):
            clean_val = f"{m.group(1)}.{m.group(2)}M"
        elif m := re.fullmatch(r"(\d+)R(\d+)", clean_val):
            clean_val = f"{m.group(1)}.{m.group(2)}R"
        elif re.fullmatch(r"\d+R", clean_val):
            clean_val = clean_val
        elif re.fullmatch(r"\d+", clean_val):
            clean_val = f"{clean_val}R"

        specs["resistance_str"] = clean_val
        return specs

    # Yageo catalog / document code e.g. PYU-RC_51_ROHS_P, PYU-RC_GROUP_51_ROHS_L
    if pn.startswith("PYU-RC"):
        specs["manufacturer"] = "Yageo"
        specs["mounting"] = "SMD"
        specs["composition"] = "Thick Film"
        specs["special_def"] = "General"
        return specs

    # 5. Royal Ohm Chip Resistor: e.g. 0805W8F1001T5E, 0603WAF1002T5E, 0603(S/W)AJ0221T5E, 1206W4F1001T5E
    clean_royal_pn = re.sub(r"\(S/W\)", "W", pn)
    royal_match = re.match(r"^(\d{4})([A-Z0-9]{2})([A-Z])([0-9A-Z]{4})([A-Z0-9]+)?$", clean_royal_pn)
    if royal_match and any(pkg in royal_match.group(1) for pkg in ("0201", "0402", "0603", "0805", "1206", "1210", "1812", "2010", "2512")):
        pkg, watt_code, tol_code, val_code, _ = royal_match.groups()
        watt_map = {
            "WG": "1/16W", "WA": "1/10W", "W8": "1/8W", "W4": "1/4W", "W2": "1/2W",
            "1W": "1W", "2W": "2W", "S4": "1/4W", "S8": "1/8W", "SA": "1/10W",
        }
        tol_map = {"F": "1%", "J": "5%", "D": "0.5%", "B": "0.1%", "G": "2%"}
        specs["manufacturer"] = "Royal Ohm"
        specs["mounting"] = "SMD"
        specs["package"] = pkg
        if watt_code in watt_map:
            specs["wattage"] = watt_map[watt_code]
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]
        specs["composition"] = "Thick Film"
        specs["special_def"] = "General"
        if val_code == "0000":
            specs["resistance_str"] = "0R"
        elif len(val_code) == 4 and val_code.isdigit():
            base = int(val_code[:3])
            multiplier = int(val_code[3])
            r_val = base * (10 ** multiplier)
            if r_val >= 1_000_000:
                specs["resistance_str"] = f"{r_val / 1_000_000:g}M"
            elif r_val >= 1_000:
                specs["resistance_str"] = f"{r_val / 1_000:g}k"
            else:
                specs["resistance_str"] = f"{r_val}R"
        elif "R" in val_code:
            specs["resistance_str"] = val_code
        return specs

    # 6. TTAF Connectors: e.g. 200G1-04R, 254G1-04R, 200W-04, 254H-04, 200D1-04
    ttaf_m = re.match(r"^(?:TTAF[-_])?(\d{3})([A-Z0-9]+)-(\d{2})([A-Z0-9]*)$", pn)
    if ttaf_m:
        p_code, s_code, pin_code, ang_code = ttaf_m.groups()
        pitch_map = {
            "100": "1mm", "125": "1.25mm", "127": "1.27mm", "150": "1.5mm",
            "200": "2mm", "250": "2.5mm", "254": "2.54mm", "300": "3mm",
            "381": "3.81mm", "396": "3.96mm", "420": "4.2mm", "500": "5mm", "508": "5.08mm"
        }
        specs["manufacturer"] = "Ttaf"
        if p_code in pitch_map:
            specs["pitch"] = pitch_map[p_code]
        try:
            specs["pin_number"] = f"{int(pin_code)} pin"
        except ValueError:
            pass

        if any(c in s_code for c in ("G", "W", "D", "P")):
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Erkek"
            specs["mounting"] = "TH"
        elif "H" in s_code:
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Dişi"
            specs["mounting"] = "Cable"
        else:
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Erkek"
            specs["mounting"] = "TH"

        if "R" in ang_code:
            specs["angle"] = "90°(Yatık)"
        else:
            specs["angle"] = "180°(Dik)"
        return specs

    # 7. Molex Connectors: e.g. 22-01-2045, 22-23-2041, 22-05-2041
    molex_m = re.match(r"^(?:00)?(22)-(\d{2})-(2\d{2}\d)$", pn)
    if molex_m:
        prefix_code, series_code, item_code = molex_m.groups()
        specs["manufacturer"] = "Molex"
        specs["pitch"] = "2.54mm"
        if len(item_code) == 4 and item_code[1:3].isdigit():
            specs["pin_number"] = f"{int(item_code[1:3])} pin"
        if series_code == "01":
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Dişi"
            specs["mounting"] = "Cable"
        elif series_code == "23":
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Erkek"
            specs["mounting"] = "TH"
            specs["angle"] = "180°(Dik)"
        elif series_code == "05":
            specs["type_connector"] = "Wire to Board Konnektör"
            specs["pin_type"] = "Erkek"
            specs["mounting"] = "TH"
            specs["angle"] = "90°(Yatık)"
        return specs

    # 8. JST Connectors: e.g. B4B-XH-A, S4B-XH-A, XHP-4, B4B-PH-K-S, S4B-PH-K-S, PHR-4
    jst_hdr = re.match(r"^([BS])(\d+)B-(XH|PH|VH|EH)", pn)
    if jst_hdr:
        ang_char, pin_count, fam = jst_hdr.groups()
        specs["manufacturer"] = "JST"
        specs["type_connector"] = "Wire to Board Konnektör"
        specs["pin_type"] = "Erkek"
        specs["mounting"] = "TH"
        specs["pin_number"] = f"{int(pin_count)} pin"
        specs["angle"] = "90°(Yatık)" if ang_char == "S" else "180°(Dik)"
        jst_pitches = {"PH": "2mm", "XH": "2.5mm", "EH": "2.5mm", "VH": "3.96mm"}
        if fam in jst_pitches:
            specs["pitch"] = jst_pitches[fam]
        return specs

    jst_hsg = re.match(r"^(XH|PH|VH|EH)P?R?-(\d+)$", pn)
    if jst_hsg:
        fam, pin_count = jst_hsg.groups()
        specs["manufacturer"] = "JST"
        specs["type_connector"] = "Wire to Board Konnektör"
        specs["pin_type"] = "Dişi"
        specs["mounting"] = "Cable"
        specs["pin_number"] = f"{int(pin_count)} pin"
        jst_pitches = {"PH": "2mm", "XH": "2.5mm", "EH": "2.5mm", "VH": "3.96mm"}
        if fam in jst_pitches:
            specs["pitch"] = jst_pitches[fam]
        return specs

    return specs


def extract_component_specs(material_type: str, text: str, part_number: str = "", is_catalog: bool = False) -> dict[str, Any]:
    """Extract physical and electrical specifications from decoded PN + full text."""
    specs = decode_manufacturer_pn(part_number)
    text_lower = text.casefold()
    is_cat = is_catalog or is_catalog_document_name(part_number)

    # Mounting
    if "mounting" not in specs:
        if material_type == "CON":
            if any(k in text_lower for k in ("smd", "smt", "surface mount")):
                specs["mounting"] = "SMD"
            else:
                specs["mounting"] = "TH"
        elif "smd" in text_lower or "smt" in text_lower or "surface mount" in text_lower or "chip" in text_lower:
            specs["mounting"] = "SMD"
        elif "through hole" in text_lower or "tht" in text_lower or "dip" in text_lower or "radial" in text_lower:
            specs["mounting"] = "TH"
        else:
            specs["mounting"] = "SMD"  # standard default for modern electronic parts

    # Package (Avoid greedy first match from table if catalog series)
    if "package" not in specs and not is_cat:
        pkg = _extract_package(text)
        if pkg:
            specs["package"] = pkg
        else:
            # Check metric or other footprints
            if "2012" in text and "0805" in text:
                specs["package"] = "0805"
            elif "1608" in text and "0603" in text:
                specs["package"] = "0603"
            elif "1005" in text and "0402" in text:
                specs["package"] = "0402"
            elif "sot-23" in text_lower or "sot23" in text_lower:
                specs["package"] = "SOT-23"
            elif "sod-123" in text_lower or "sod123" in text_lower:
                specs["package"] = "SOD-123"
            elif "sma" in text_lower or "do-214ac" in text_lower:
                specs["package"] = "SMA"
            elif "smb" in text_lower or "do-214aa" in text_lower:
                specs["package"] = "SMB"
            elif "to-220" in text_lower:
                specs["package"] = "TO-220"
            elif "to-92" in text_lower:
                specs["package"] = "TO-92"

    # Manufacturer
    if "manufacturer" not in specs:
        for kw, name, code in KNOWN_MANUFACTURERS:
            if kw in text_lower:
                specs["manufacturer"] = name
                specs["manufacturer_code"] = code
                break

    # Material-specific attributes
    if material_type == "CAP":
        # Capacitance
        if "capacitance_str" not in specs and not is_cat:
            cap_dec = _extract_capacitance(text_lower)
            if cap_dec is not None:
                specs["capacitance_str"] = _format_cap(cap_dec)

        # Voltage
        if "voltage" not in specs and not is_cat:
            volt_dec = _extract_voltage(text_lower)
            if volt_dec is not None:
                specs["voltage"] = f"{int(volt_dec) if volt_dec == int(volt_dec) else volt_dec}V"

        # Tolerance
        if "tolerance" not in specs and not is_cat:
            tol_dec = _extract_tolerance(text_lower)
            if tol_dec is not None:
                specs["tolerance"] = f"%{int(tol_dec)}"

        # Dielectric / Special
        if "special_def" not in specs:
            for d in ["X7R", "C0G", "NP0", "X5R", "Y5V", "X7S"]:
                if d.casefold() in text_lower:
                    specs["special_def"] = d
                    break

        # Temperature
        if "-55" in text and "125" in text:
            specs["temperature"] = "-55°C~+125°C"
        elif "-40" in text and "105" in text:
            specs["temperature"] = "-40°C~+105°C"
        elif "-40" in text and "85" in text:
            specs["temperature"] = "-40°C~+85°C"

    elif material_type == "RES":
        # Resistance (do not guess for catalog)
        if "resistance_str" not in specs and not is_cat:
            res_dec = _extract_resistance(text_lower)
            if res_dec is not None:
                specs["resistance_str"] = _format_res(res_dec)

        # Tolerance (do not guess for catalog)
        if "tolerance" not in specs and not is_cat:
            tol_dec = _extract_tolerance(text_lower)
            if tol_dec is not None:
                specs["tolerance"] = f"%{int(tol_dec)}"

        # Wattage (do not guess for catalog)
        if "wattage" not in specs and not is_cat:
            for w in ["1/20W", "1/16W", "1/10W", "1/8W", "1/4W", "1/2W", "1W", "2W", "3W", "5W"]:
                if w.casefold() in text_lower:
                    specs["wattage"] = w
                    break

        # Composition - Resistors (SORTNO 60)
        if "composition" not in specs:
            if "thick film" in text_lower:
                specs["composition"] = "Thick Film"
            elif "thin film" in text_lower:
                specs["composition"] = "Thin Film"
            elif "metal film" in text_lower:
                specs["composition"] = "Metal Film"
            elif "carbon film" in text_lower:
                specs["composition"] = "Carbon Film"
            elif "wirewound" in text_lower:
                specs["composition"] = "Wirewound"
            elif "resistor array" in text_lower or "array" in text_lower:
                specs["composition"] = "Resistor Array"
            else:
                specs["composition"] = "Thick Film" if specs.get("mounting") == "SMD" else "General"

        # Special Definition - Resistors (SORTNO 70) [Automotive AEC-Q200 vs General]
        # Uses series-aware detection to prevent false positives from footnotes/reference tables
        if "special_def" not in specs:
            if is_automotive_qualified(part_number, text):
                specs["special_def"] = "Automotive AEC-Q200"
            elif "anti surge" in text_lower or "anti-surge" in text_lower:
                specs["special_def"] = "Anti Surge"
            elif "flame proof" in text_lower or "flameproof" in text_lower:
                specs["special_def"] = "Flame Proof"
            elif "moisture resistant" in text_lower:
                specs["special_def"] = "Moisture Resistant"
            else:
                specs["special_def"] = "General"

    elif material_type == "DIO":
        # Diode Type
        if "type_diode" not in specs:
            if "schottky" in text_lower:
                specs["type_diode"] = "Schottky"
            elif "zener" in text_lower:
                specs["type_diode"] = "Zener"
            elif "fast switching" in text_lower or "switching diode" in text_lower:
                specs["type_diode"] = "Fast Switching"
            elif "tvs" in text_lower or "transient voltage" in text_lower:
                specs["type_diode"] = "Tvs"
            elif "bridge" in text_lower:
                specs["type_diode"] = "Bridge"

        # Diode Value / PN
        specs["value_diode"] = part_number.strip()

    elif material_type in ("FET", "TRA"):
        if material_type == "FET":
            if "n-channel" in text_lower or "n channel" in text_lower:
                specs["fet_channel"] = "N Channel"
            elif "p-channel" in text_lower or "p channel" in text_lower:
                specs["fet_channel"] = "P Channel"
            specs["value_fet"] = part_number.strip()
        else:
            if "npn" in text_lower:
                specs["tra_type"] = "NPN"
            elif "pnp" in text_lower:
                specs["tra_type"] = "PNP"
            specs["value_tra"] = part_number.strip()

    elif material_type == "ICR":
        if "ldo" in text_lower or "linear" in text_lower:
            specs["type_regulator"] = "LDO"
        elif "reference" in text_lower:
            specs["type_regulator"] = "Voltage Reference"
        volt_dec = _extract_voltage(text_lower)
        if volt_dec is not None:
            specs["regulator_voltage"] = f"{volt_dec}V"

    elif material_type == "CON":
        # Connector Type
        if "type_connector" not in specs:
            if any(k in text_lower for k in ("wire to board", "wire-to-board", "crimp housing", "housing")):
                specs["type_connector"] = "Wire to Board Konnektör"
            elif any(k in text_lower for k in ("pin header", "header")):
                specs["type_connector"] = "Pin Header"
            elif "terminal block" in text_lower or "klemens" in text_lower:
                specs["type_connector"] = "PCB Klemens"

        # Pin Type
        if "pin_type" not in specs:
            if any(k in text_lower for k in ("receptacle", "housing", "female", "dişi")):
                specs["pin_type"] = "Dişi"
            elif any(k in text_lower for k in ("header", "plug", "male", "erkek", "wafer", "pin")):
                specs["pin_type"] = "Erkek"

        # Pitch
        if "pitch" not in specs:
            pitch_m = re.search(r"(?:pitch|adım|spacing|grid)[:\s=]+(\d+(?:\.\d+)?)\s*mm\b", text_lower)
            if not pitch_m:
                pitch_m = re.search(r"\b(1\.00?|1\.25|1\.27|1\.50?|2\.00?|2\.50?|2\.54|3\.00?|3\.81|3\.96|4\.20?|5\.00?|5\.08)\s*mm\b", text_lower)
            if not pitch_m:
                pitch_m = re.search(r"(\d+(?:\.\d+)?)\s*mm\b", text_lower)
            if pitch_m:
                val = float(pitch_m.group(1))
                specs["pitch"] = f"{int(val) if val == int(val) else val}mm"

        # Pin Number
        if "pin_number" not in specs:
            pins_m = re.search(r"(?:circuits?|pins?|positions?|poles?|kutup|pin sayısı)[:\s=]+(\d+)\b", text_lower)
            if not pins_m:
                pins_m = re.search(r"\b(\d+)\s*(?:circuits?|pins?|pos|poles?|ways?|kutup)\b", text_lower)
            if pins_m:
                specs["pin_number"] = f"{pins_m.group(1)} pin"

        # Angle
        if "angle" not in specs:
            if any(k in text_lower for k in ("right angle", "90 deg", "90°", "horizontal", "yatık")):
                specs["angle"] = "90°(Yatık)"
            elif any(k in text_lower for k in ("vertical", "straight", "180 deg", "180°", "dik")):
                specs["angle"] = "180°(Dik)"

    return specs


CORE_ATTRIBUTES_BY_TYPE: dict[str, set[str]] = {
    "RES": {"mounting", "package", "value", "tolerance", "wattage", "compositionresistors", "specialdefinitionresistors", "q"},
    "CAP": {"mounting", "package", "value", "voltage", "tolerance", "temperature", "specialdefinitioncapacitors", "q"},
    "DIO": {"mounting", "package", "typediode", "value", "q"},
    "CON": {"mounting", "typeconnector", "pintypeofconnector", "pitchconnector", "pitch", "pinnumber", "angle", "q"},
    "IND": {"mounting", "package", "value", "q"},
    "FET": {"mounting", "package", "value", "fetchanneltype", "q"},
    "TRA": {"mounting", "package", "value", "typetransistor", "q"},
    "ICR": {"mounting", "package", "typeregulator", "regulatoroutputvoltage", "voltage", "q"},
    "ICS": {"mounting", "package", "value", "q"},
    "MCU": {"mounting", "package", "value", "q"},
    "LED": {"mounting", "package", "ledcolor", "q"},
    "OPT": {"mounting", "package", "value", "q"},
    "REL": {"mounting", "relaycoilvoltage", "relaycontactvoltage", "q"},
    "FUS": {"mounting", "package", "typefuse", "fusecurrent", "q"},
    "VAR": {"mounting", "package", "varistorsize", "voltageac", "q"},
}


def map_specs_to_erp(rulebook: RuleBook, material_type: str, specs: dict[str, Any]) -> tuple[list[PropertyMapping], str, str]:
    """Map extracted specifications to the exact ERP rules for the material type."""
    type_key = _normalize_label(material_type)
    typed_rules = rulebook.by_type_attribute.get(type_key, {})
    if not typed_rules:
        # Fallback if material type not found in rulebook
        return [], "", ""

    # Sort attributes by their official SORTNO
    attribute_sorts: list[tuple[float, str, tuple[RuleEntry, ...]]] = []
    for attr_norm, entries in typed_rules.items():
        if not entries:
            continue
        try:
            s_order = float(entries[0].sort_order)
        except (ValueError, TypeError):
            s_order = 999.0
        attribute_sorts.append((s_order, attr_norm, entries))

    attribute_sorts.sort(key=lambda x: x[0])

    mappings: list[PropertyMapping] = []
    description_tokens: list[str] = []
    code_tokens: list[str] = []

    is_ceramic_mlcc = (
        material_type == "CAP" and (
            specs.get("package") in ("0201", "0402", "0603", "0805", "1206", "1210", "1812", "2220") or
            specs.get("special_def") in ("X7R", "X5R", "C0G", "NP0", "Y5V", "X7S", "NPO") or
            "mlcc" in str(specs).casefold() or
            (specs.get("mounting") == "SMD" and not any(k in str(specs).casefold() for k in ("electrolytic", "elec", "aluminium", "aluminum", "tantalum")))
        )
    )
    is_electrolytic = (
        material_type == "CAP" and (
            any(k in str(specs).casefold() for k in ("electrolytic", "elec", "aluminium", "aluminum")) or
            bool(specs.get("load_life")) or bool(specs.get("size_cap"))
        )
    )

    for sort_no, attr_norm, entries in attribute_sorts:
        attr_display_name = entries[0].attribute_name
        matched_entry: RuleEntry | None = None
        extracted_val_str = ""
        status = "BULUNAMADI"

        # Sub-group conditional attributes: skip attributes that do not apply to this component subtype
        if is_ceramic_mlcc and any(k in attr_norm for k in ("loadlife", "pitch", "sizecapacitor")):
            continue

        if is_electrolytic and "specialdefinition" in attr_norm:
            continue

        if material_type == "CON" and "serial" in attr_norm:
            continue

        # 1. MOUNTING
        if attr_norm == "mounting":
            target = specs.get("mounting", "SMD").upper()
            extracted_val_str = target
            for e in entries:
                if e.value == "S" and target == "SMD":
                    matched_entry = e
                    break
                elif e.value == "T" and target == "TH":
                    matched_entry = e
                    break
                elif e.value_description and target in e.value_description.upper():
                    matched_entry = e
                    break
            if matched_entry:
                status = "TAM EŞLEŞTİ"

        # 2. PACKAGE
        elif attr_norm == "package":
            pkg = specs.get("package", "")
            extracted_val_str = pkg
            if pkg:
                for e in entries:
                    desc_clean = re.sub(r"[^A-Za-z0-9]", "", e.value_description or "").upper()
                    pkg_clean = re.sub(r"[^A-Za-z0-9]", "", pkg).upper()
                    if pkg_clean and (pkg_clean in desc_clean or desc_clean == pkg_clean):
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 3. VALUE
        elif attr_norm == "value":
            if material_type == "CAP":
                target_val = specs.get("capacitance_str", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        # Compare normalized capacitance
                        c_entry = _extract_capacitance(e.value_description.casefold())
                        c_target = _extract_capacitance(target_val.casefold())
                        if c_entry is not None and c_target is not None and c_entry == c_target:
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                        elif target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "YAKIN EŞLEŞTİ"
                            break

            elif material_type == "RES":
                target_val = specs.get("resistance_str", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        r_entry = _extract_resistance(e.value_description.casefold())
                        r_target = _extract_resistance(target_val.casefold())
                        if r_entry is not None and r_target is not None and r_entry == r_target:
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                        elif target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "YAKIN EŞLEŞTİ"
                            break

            elif material_type == "DIO":
                target_val = specs.get("value_diode", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        if target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break

            elif material_type in ("FET", "TRA"):
                target_val = specs.get("value_fet" if material_type == "FET" else "value_tra", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        if target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break

        # 4. VOLTAGE
        elif attr_norm == "voltage":
            target_volt = specs.get("voltage", "")
            extracted_val_str = target_volt
            if target_volt:
                for e in entries:
                    v_entry = _extract_voltage(e.value_description.casefold())
                    v_target = _extract_voltage(target_volt.casefold())
                    if v_entry is not None and v_target is not None and v_entry == v_target:
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 5. TOLERANCE
        elif attr_norm == "tolerance":
            target_tol = specs.get("tolerance", "")
            extracted_val_str = target_tol
            if target_tol:
                for e in entries:
                    t_entry = _extract_tolerance(e.value_description.casefold())
                    t_target = _extract_tolerance(target_tol.casefold())
                    if t_entry is not None and t_target is not None and t_entry == t_target:
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
                    # Check letter code or decimal representation (e.g. 0.01 for 1%, 0.1 for 10%)
                    tol_num = re.search(r"(\d+)", target_tol)
                    if tol_num and e.value_description:
                        try:
                            dec = float(e.value_description.strip(" ,%"))
                            if dec < 1.0 and abs(dec * 100 - float(tol_num.group(1))) < 0.01:
                                matched_entry = e
                                status = "TAM EŞLEŞTİ"
                                break
                        except ValueError:
                            pass

        # 6. COMPOSITION (e.g. Composition - Resistors)
        elif "composition" in attr_norm:
            target_comp = specs.get("composition", "")
            extracted_val_str = target_comp
            if target_comp:
                for e in entries:
                    if target_comp.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 7. SPECIAL DEFINITION / DIELECTRIC
        elif "specialdefinition" in attr_norm:
            target_sp = specs.get("special_def", "")
            extracted_val_str = target_sp
            if target_sp:
                for e in entries:
                    if target_sp.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 8. WATTAGE
        elif attr_norm == "wattage":
            target_w = specs.get("wattage", "")
            extracted_val_str = target_w
            if target_w:
                for e in entries:
                    if target_w.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 9. TEMPERATURE
        elif attr_norm == "temperature":
            target_temp = specs.get("temperature", "")
            extracted_val_str = target_temp
            if target_temp:
                for e in entries:
                    if target_temp.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 10. MANUFACTURER (Q)
        elif attr_norm == "q":
            target_mfr = specs.get("manufacturer", "")
            extracted_val_str = target_mfr
            if target_mfr:
                for e in entries:
                    if target_mfr.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 11. CONNECTOR ATTRIBUTES
        elif "typeconnector" in attr_norm:
            tc = specs.get("type_connector", "")
            extracted_val_str = tc
            if tc:
                for e in entries:
                    if tc.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "pintype" in attr_norm:
            pt = specs.get("pin_type", "")
            extracted_val_str = pt
            if pt:
                for e in entries:
                    if pt.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "pitch" in attr_norm:
            p_val = specs.get("pitch", "")
            extracted_val_str = p_val
            if p_val:
                target_pitch_m = re.search(r"(\d+(?:\.\d+)?)", p_val)
                target_pitch_num = float(target_pitch_m.group(1)) if target_pitch_m else None
                for e in entries:
                    em = re.search(r"(\d+(?:\.\d+)?)\s*mm", e.value_description.casefold())
                    if em and target_pitch_num is not None:
                        if abs(float(em.group(1)) - target_pitch_num) < 0.001:
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                    elif p_val.casefold() == e.value_description.strip().rstrip(",").casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "pinnumber" in attr_norm:
            pn_val = specs.get("pin_number", "")
            extracted_val_str = pn_val
            if pn_val:
                t_m = re.search(r"(\d+)", pn_val)
                target_pin_count = int(t_m.group(1)) if t_m else None
                for e in entries:
                    desc_clean = e.value_description.strip().rstrip(",").casefold()
                    if desc_clean == f"{target_pin_count} pin" or desc_clean == pn_val.strip().casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
                if not matched_entry and target_pin_count is not None:
                    for e in entries:
                        em = re.search(r"^(\d+)\s*pin", e.value_description.strip().casefold())
                        if em and int(em.group(1)) == target_pin_count:
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
        elif "angle" in attr_norm:
            ang = specs.get("angle", "")
            extracted_val_str = ang
            if ang:
                for e in entries:
                    if ang.casefold() in e.value_description.casefold() or (
                        "90" in ang and "90" in e.value_description
                    ) or (
                        "180" in ang and "180" in e.value_description
                    ):
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        if matched_entry:
            mappings.append(PropertyMapping(
                attribute_name=attr_display_name,
                sort_no=sort_no,
                extracted_value=extracted_val_str,
                matched_erp_value=matched_entry.value_description,
                erp_code=matched_entry.value,
                prefix=matched_entry.prefix or "",
                status=status,
            ))
            desc_val = matched_entry.value_description.strip()
            if desc_val:
                description_tokens.append(desc_val)
            code_val = (matched_entry.prefix or "") + (matched_entry.value or "")
            if code_val:
                code_tokens.append(code_val)
        elif extracted_val_str:
            mappings.append(PropertyMapping(
                attribute_name=attr_display_name,
                sort_no=sort_no,
                extracted_value=extracted_val_str,
                matched_erp_value="—",
                erp_code="—",
                prefix="",
                status="SÖZLÜKTE YOK",
            ))
        else:
            # Attribute is in the ERP dictionary schema for this material type, but not specified in datasheet.
            # Only include if it is a CORE essential attribute for this material type (e.g. resistance, voltage, package)
            core_attrs = CORE_ATTRIBUTES_BY_TYPE.get(material_type, {"mounting", "package", "value", "q"})
            if attr_norm in core_attrs:
                mappings.append(PropertyMapping(
                    attribute_name=attr_display_name,
                    sort_no=sort_no,
                    extracted_value="Belirtilmemiş",
                    matched_erp_value="—",
                    erp_code="—",
                    prefix="",
                    status="BULUNAMADI",
                ))

    suggested_desc = " ".join(description_tokens).strip()
    if specs.get("is_catalog"):
        suggested_desc = f"{suggested_desc} [Katalog Dokümanı - Tam MPN Belirtiniz]".strip()

    first_token = code_tokens[0] if code_tokens else ""
    if first_token.startswith(material_type):
        suggested_code = "".join(code_tokens).strip()
    else:
        suggested_code = (material_type + "".join(code_tokens)).strip() if code_tokens else ""

    return mappings, suggested_desc, suggested_code


def analyze_single_datasheet(
    filename: str,
    file_bytes: bytes,
    rulebook: RuleBook,
    part_number_override: str | None = None,
    material_type_override: str | None = None,
) -> DatasheetAnalysisResult:
    """Analyze a single datasheet PDF and map its properties to DEKA ERP format."""
    raw_pn = (part_number_override or Path(filename).stem).strip()
    raw_pn = re.sub(r"\.pdf$", "", raw_pn, flags=re.I).strip()
    pn = raw_pn

    try:
        text, page_count, analyzed_pages = extract_pdf_data(file_bytes, part_number=pn)
    except Exception as exc:
        return DatasheetAnalysisResult(
            filename=filename,
            part_number=pn,
            material_type="BİLİNMİYOR",
            material_type_name="Bilinmiyor",
            inferred_mounting="",
            confidence=0.0,
            suggested_description="",
            suggested_erp_code="",
            page_count=0,
            error=f"PDF okunamadı: {str(exc)}",
        )

    # 1. Check whether input is a catalog/document code
    is_catalog = is_catalog_document_name(pn)
    detected_mpns = extract_part_numbers_from_text(text)

    # If user didn't enter a manual override, and text has a single unique genuine MPN
    if is_catalog and not part_number_override and len(detected_mpns) == 1:
        pn = detected_mpns[0]
        is_catalog = False

    if material_type_override and material_type_override.upper() != "AUTO":
        mat_type = material_type_override.upper()
        conf = 1.0
        reason = f"Kullanıcı tarafından seçildi ({mat_type})."
    else:
        mat_type, conf, reason = classify_material_type(text, pn)

    mat_type_name = MATERIAL_TYPE_NAMES.get(mat_type, mat_type)

    specs = extract_component_specs(mat_type, text, pn, is_catalog=is_catalog)
    specs["is_catalog"] = is_catalog
    mounting = specs.get("mounting", "SMD")

    mappings, suggested_desc, suggested_code = map_specs_to_erp(rulebook, mat_type, specs)

    if is_catalog:
        conf = min(conf, 0.65)
        mfr = specs.get("manufacturer", "Üretici")
        summary = (
            f"⚠️ Katalog/Seri Dokümanı Tespit Edildi ({raw_pn}). "
            f"Bu doküman {mfr} genel ürün serisi kataloğudur (birden fazla kılıf ve değer içerir). "
            f"Kesin kılıf, direnç, tolerans ve güç eşleşmesi için lütfen tam parça kodunu (örn: RC0805JR-07300RL) giriniz."
        )
    else:
        summary = f"{mat_type_name} olarak tespit edildi ({int(conf * 100)}% güven). {reason}"

    return DatasheetAnalysisResult(
        filename=filename,
        part_number=pn,
        material_type=mat_type,
        material_type_name=mat_type_name,
        inferred_mounting=mounting,
        confidence=conf,
        suggested_description=suggested_desc,
        suggested_erp_code=suggested_code,
        properties=mappings,
        raw_specs=specs,
        page_count=page_count,
        analyzed_pages=analyzed_pages,
        summary_text=summary,
        is_catalog=is_catalog,
        detected_part_numbers=detected_mpns,
    )


def generate_datasheet_excel_report(analyses: list[DatasheetAnalysisResult], output_path: Path) -> None:
    """Generate a clean, styled Excel report with DEKA colors for datasheet analyses."""
    wb = Workbook()
    
    # ── Sheet 1: Genel Özet (Summary) ──
    ws_summary = wb.active
    ws_summary.title = "Datasheet ERP Eşleme Özeti"
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="CA130D", end_color="CA130D", fill_type="solid")  # DEKA Red
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    headers_summary = [
        "No",
        "Dosya Adı",
        "Part Number",
        "Malzeme Tipi",
        "Montaj (TH/SMD)",
        "Önerilen ERP Açıklaması",
        "Güven Skoru",
        "Sayfa Sayısı",
        "Eşleşen Özellik Sayısı",
    ]
    ws_summary.append(headers_summary)
    for col_idx in range(1, len(headers_summary) + 1):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border

    # ── Sheet 2: Detaylı Özellikler (Properties Breakdown) ──
    ws_details = wb.create_sheet(title="Özellik Eşleme Detayları")
    headers_details = [
        "Part Number",
        "Malzeme Tipi",
        "ERP Özellik Adı",
        "SORTNO",
        "Datasheet Değeri",
        "Eşleşen ERP Açıklaması (VALUE AÇIKLAMASI)",
        "ERP Değer Kodu (VALUE)",
        "ERP Prefix",
        "Eşleşme Durumu",
    ]
    ws_details.append(headers_details)
    for col_idx in range(1, len(headers_details) + 1):
        cell = ws_details.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border

    # Fill data
    for idx, item in enumerate(analyses, start=1):
        # Summary row
        ws_summary.append([
            idx,
            item.filename,
            item.part_number,
            item.material_type_name,
            item.inferred_mounting,
            item.suggested_description,
            f"{int(item.confidence * 100)}%",
            item.page_count,
            len([p for p in item.properties if p.status == "TAM EŞLEŞTİ"]),
        ])
        s_row = ws_summary.max_row
        for col_idx in range(1, len(headers_summary) + 1):
            cell = ws_summary.cell(row=s_row, column=col_idx)
            cell.border = thin_border
            cell.alignment = align_center if col_idx in (1, 4, 5, 8, 9, 10) else align_left

        # Details rows
        for p in item.properties:
            ws_details.append([
                item.part_number,
                item.material_type,
                p.attribute_name,
                p.sort_no,
                p.extracted_value,
                p.matched_erp_value,
                p.erp_code,
                p.prefix,
                p.status,
            ])
            d_row = ws_details.max_row
            for col_idx in range(1, len(headers_details) + 1):
                cell = ws_details.cell(row=d_row, column=col_idx)
                cell.border = thin_border
                cell.alignment = align_center if col_idx in (2, 4, 7, 8, 9) else align_left

    # Auto column width
    for ws in (ws_summary, ws_details):
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
