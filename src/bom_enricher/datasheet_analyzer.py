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


def classify_material_type(text: str, part_number: str = "") -> tuple[str, float, str]:
    """Determine component material type (CAP, RES, DIO, etc.) from PN and datasheet content."""
    pn = part_number.strip().upper()
    text_lower = text.casefold()

    # 1. High confidence Part Number prefixes
    if re.match(r"^CL\d{2}[A-Z]\d{3}", pn) or re.match(r"^GRM\d{2}", pn) or re.match(r"^CC\d{4}", pn) or re.match(r"^AC\d{4}[A-Z]", pn):
        return "CAP", 0.98, "Part Number standart seramik kapasitör kodlamasına uyuyor."
    if re.match(r"^RC\d{4}", pn) or re.match(r"^CRCW\d{4}", pn) or re.match(r"^ERJ\d", pn) or re.match(r"^WR\d{2}", pn):
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
    if re.match(r"^(?:22-01|0022|2695|5557|JST-|PH-|XH-|VH-|MOLEX)", pn):
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

    # 4. Yageo Chip Resistor: e.g. RC0805FR-0710KL or RC0603JR-07100RL
    yageo_res = re.match(r"^RC(\d{4})([A-Z])R-07([0-9A-Z]+)", pn)
    if yageo_res:
        pkg, tol_code, val_code = yageo_res.groups()
        tol_map = {"F": "1%", "J": "5%", "D": "0.5%"}
        specs["manufacturer"] = "Yageo"
        specs["mounting"] = "SMD"
        specs["package"] = pkg
        if tol_code in tol_map:
            specs["tolerance"] = tol_map[tol_code]
        specs["special_def"] = "General"
        specs["composition"] = "Thick Film"

        # 10KL -> 10k, 100RL -> 100R, 0RL -> 0R
        clean_val = val_code.rstrip("L")
        if clean_val.endswith("K"):
            clean_val = clean_val[:-1] + "k"
        specs["resistance_str"] = clean_val
        return specs

    # 5. Royal Ohm Chip Resistor: e.g. 0805W8F1001T5E, 0603WAF1002T5E, 1206W4F1001T5E, 0402WGJ0102TCE
    royal_match = re.match(r"^(\d{4})([A-Z0-9]{2})([A-Z])([0-9A-Z]{4})([A-Z0-9]+)?$", pn)
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

    return specs


def extract_component_specs(material_type: str, text: str, part_number: str = "") -> dict[str, Any]:
    """Extract physical and electrical specifications from decoded PN + full text."""
    specs = decode_manufacturer_pn(part_number)
    text_lower = text.casefold()

    # Mounting
    if "mounting" not in specs:
        if "smd" in text_lower or "smt" in text_lower or "surface mount" in text_lower or "chip" in text_lower:
            specs["mounting"] = "SMD"
        elif "through hole" in text_lower or "tht" in text_lower or "dip" in text_lower or "radial" in text_lower:
            specs["mounting"] = "TH"
        else:
            specs["mounting"] = "SMD"  # standard default for modern electronic parts

    # Package
    if "package" not in specs:
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
        if "capacitance_str" not in specs:
            cap_dec = _extract_capacitance(text_lower)
            if cap_dec is not None:
                specs["capacitance_str"] = _format_cap(cap_dec)

        # Voltage
        if "voltage" not in specs:
            volt_dec = _extract_voltage(text_lower)
            if volt_dec is not None:
                specs["voltage"] = f"{int(volt_dec) if volt_dec == int(volt_dec) else volt_dec}V"

        # Tolerance
        if "tolerance" not in specs:
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
        # Resistance
        if "resistance_str" not in specs:
            res_dec = _extract_resistance(text_lower)
            if res_dec is not None:
                specs["resistance_str"] = _format_res(res_dec)

        # Tolerance
        if "tolerance" not in specs:
            tol_dec = _extract_tolerance(text_lower)
            if tol_dec is not None:
                specs["tolerance"] = f"%{int(tol_dec)}"

        # Wattage
        if "wattage" not in specs:
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
        if "special_def" not in specs:
            if any(k in text_lower for k in ("aec-q200", "aec q200", "automotive", "otomotiv")):
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
        if any(k in text_lower for k in ("wire to board", "wire-to-board", "crimp housing", "housing")):
            specs["type_connector"] = "Wire to Board Konnektör"
        elif any(k in text_lower for k in ("pin header", "header")):
            specs["type_connector"] = "Pin Header"
        elif "terminal block" in text_lower or "klemens" in text_lower:
            specs["type_connector"] = "PCB Klemens"

        # Pin Type
        if any(k in text_lower for k in ("receptacle", "housing", "female", "dişi")):
            specs["pin_type"] = "Dişi"
        elif any(k in text_lower for k in ("header", "plug", "male", "erkek")):
            specs["pin_type"] = "Erkek"

        # Pitch
        pitch_m = re.search(r"(\d+(?:\.\d+)?)\s*mm\b", text_lower)
        if pitch_m:
            specs["pitch"] = f"{pitch_m.group(1)}mm"

        # Pin Number
        pins_m = re.search(r"(\d+)\s*(?:circuits|circ|pin|pos|poles|ways)\b", text_lower)
        if pins_m:
            specs["pin_number"] = f"{pins_m.group(1)} pin"

        # Angle
        if any(k in text_lower for k in ("right angle", "90 deg", "90°", "horizontal")):
            specs["angle"] = "90°(Yatık)"
        elif any(k in text_lower for k in ("vertical", "straight", "180 deg", "180°")):
            specs["angle"] = "180°(Dik)"

    return specs


CORE_ATTRIBUTES_BY_TYPE: dict[str, set[str]] = {
    "RES": {"mounting", "package", "value", "tolerance", "wattage", "q"},
    "CAP": {"mounting", "package", "value", "voltage", "tolerance", "q"},
    "DIO": {"mounting", "package", "typediode", "value", "q"},
    "CON": {"mounting", "typeconnector", "pitchconnector", "pitch", "pinnumber", "q"},
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

        if material_type == "CON" and "serialconnector" in attr_norm:
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
                for e in entries:
                    if p_val.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "pinnumber" in attr_norm:
            pn_val = specs.get("pin_number", "")
            extracted_val_str = pn_val
            if pn_val:
                for e in entries:
                    if pn_val.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "angle" in attr_norm:
            ang = specs.get("angle", "")
            extracted_val_str = ang
            if ang:
                for e in entries:
                    if ang.casefold() in e.value_description.casefold():
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
    suggested_code = "".join(code_tokens).strip()

    return mappings, suggested_desc, suggested_code


def analyze_single_datasheet(
    filename: str,
    file_bytes: bytes,
    rulebook: RuleBook,
    part_number_override: str | None = None,
    material_type_override: str | None = None,
) -> DatasheetAnalysisResult:
    """Analyze a single datasheet PDF and map its properties to DEKA ERP format."""
    pn = (part_number_override or Path(filename).stem).strip()
    pn = re.sub(r"\.pdf$", "", pn, flags=re.I).strip()

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

    if material_type_override and material_type_override.upper() != "AUTO":
        mat_type = material_type_override.upper()
        conf = 1.0
        reason = f"Kullanıcı tarafından seçildi ({mat_type})."
    else:
        mat_type, conf, reason = classify_material_type(text, pn)

    mat_type_name = MATERIAL_TYPE_NAMES.get(mat_type, mat_type)

    specs = extract_component_specs(mat_type, text, pn)
    mounting = specs.get("mounting", "SMD")

    mappings, suggested_desc, suggested_code = map_specs_to_erp(rulebook, mat_type, specs)

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
