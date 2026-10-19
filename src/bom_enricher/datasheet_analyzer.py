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

from bom_enricher.comparison import (
    RuleBook,
    RuleEntry,
    _normalize_label,
    _normalize_attribute,
    _entry_text,
    _contains_phrase,
    load_rulebook,
)
from bom_enricher.pdf_renderer import select_and_render_pages
from bom_enricher.vision_client import VisionLLMClient
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
    _format_ind,
)

_DEFAULT_RULEBOOK: RuleBook | None = None


def get_default_rulebook() -> RuleBook | None:
    """Load and cache default rulebook from data/rules.xlsx if available."""
    global _DEFAULT_RULEBOOK
    if _DEFAULT_RULEBOOK is None:
        rules_path = Path(__file__).resolve().parents[2] / "data" / "rules.xlsx"
        if rules_path.exists():
            try:
                _DEFAULT_RULEBOOK = load_rulebook(rules_path)
            except Exception:
                _DEFAULT_RULEBOOK = None
    return _DEFAULT_RULEBOOK


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
    rich_description: str = ""          # e.g. "RF Board Antenna, WiFi, 50mm MHF Cable, TE Connectivity"
    suggested_erp_code: str = ""        # e.g. "CAPSS31003KD4X7R-SA"
    properties: list[PropertyMapping] = field(default_factory=list)
    raw_specs: dict[str, Any] = field(default_factory=dict)
    page_count: int = 1
    analyzed_pages: list[int] = field(default_factory=list)
    summary_text: str = ""
    error: str | None = None
    is_catalog: bool = False
    detected_part_numbers: list[str] = field(default_factory=list)
    analysis_method: str = "deterministic_text"

    def to_dict(self) -> dict[str, Any]:
        return {
            "filename": self.filename,
            "part_number": self.part_number,
            "material_type": self.material_type,
            "material_type_name": self.material_type_name,
            "inferred_mounting": self.inferred_mounting,
            "confidence": round(self.confidence, 2),
            "suggested_description": self.suggested_description,
            "rich_description": self.rich_description or self.suggested_description,
            "suggested_erp_code": self.suggested_erp_code,
            "erp_code": self.suggested_erp_code,
            "page_count": self.page_count,
            "analyzed_pages": self.analyzed_pages,
            "summary_text": self.summary_text,
            "error": self.error,
            "is_catalog": self.is_catalog,
            "detected_part_numbers": self.detected_part_numbers,
            "analysis_method": self.analysis_method,
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
            "specs": {
                k: str(v) if isinstance(v, Decimal) else v
                for k, v in self.raw_specs.items()
            },
        }


MATERIAL_TYPE_NAMES = {
    # Pasif Bileşenler
    "RES": "Direnç (RES)",
    "CAP": "Kapasitör (CAP)",
    "IND": "Bobin / İndüktör (IND)",
    "POT": "Potansiyometre (POT)",
    "FER": "Ferrit Boncuk (FER)",
    "TER": "Termistör (TER)",
    "VAR": "Varistör (VAR)",
    "FUS": "Sigorta (FUS)",
    "CRY": "Kristal / Osilatör (CRY)",
    "FLT": "Filtre (FLT)",
    # Yarı İletkenler & Aktif Bileşenler
    "DIO": "Diyot (DIO)",
    "FET": "MOSFET / FET",
    "TRA": "Transistör (TRA)",
    "TRI": "Triyak (TRI)",
    "SCR": "Tristör (SCR)",
    "LED": "LED",
    "OPT": "Optokuplör (OPT)",
    # Entegre Devreler (IC)
    "ICR": "Voltaj Regülatörü / LDO (ICR)",
    "ICS": "SMPS / Güç Entegresi (ICS)",
    "MCU": "Mikrodenetleyici (MCU)",
    "AMP": "Amplifikatör / Op-Amp (AMP)",
    "ICL": "Mantık Entegresi / Logic (ICL)",
    "ICM": "Bellek Entegresi / Memory (ICM)",
    "ICC": "Haberleşme Entegresi (ICC)",
    "ICD": "Sürücü Entegre (ICD)",
    "ICB": "Batarya Yönetim Entegresi (ICB)",
    # Elektromekanik & Bağlantı
    "CON": "Konnektör (CON)",
    "TRM": "Terminal / Klemens (TRM)",
    "SWT": "Anahtar / Switch (SWT)",
    "REL": "Röle (REL)",
    "BUZ": "Buzzer (BUZ)",
    "SPK": "Hoparlör (SPK)",
    "MIC": "Mikrofon (MIC)",
    "TRF": "Trafo (TRF)",
    "DSP": "Ekran / Display (DSP)",
    "ENC": "Enkoder (ENC)",
    "LDR": "Foto Direnç (LDR)",
    "BAH": "Pil Yuvası (BAH)",
    "CBL": "Kablo (CBL)",
    "PCB": "Baskı Devre Kartı (PCB)",
    "SGT": "Soğutucu (SGT)",
    "ANT": "Anten (ANT)",
    "JMP": "Jumper (JMP)",
    "SHI": "Shield / Koruma (SHI)",
    "SPR": "Yay / Spring (SPR)",
    "ADP": "Adaptör (ADP)",
    "FLM": "Flash Bellek (FLM)",
    "BİLİNMİYOR": "Bilinmiyor",
}


NORMALIZE_VLM_TYPE = {
    "DIODE": "DIO",
    "TRANSISTOR": "TRA",
    "TRANS": "TRA",
    "CONNECTOR": "CON",
    "INDUCTOR": "IND",
    "RESISTOR": "RES",
    "CAPACITOR": "CAP",
    "BATTERY": "BAH",
    "BUZZER": "BUZ",
    "VARISTOR": "VAR",
    "ANTENNA": "ANT",
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
    ("yixinzy", "YIXINZY", "YX"),
    ("epcos", "Epcos", "EP"),
    ("multicomp pro", "Multicomp Pro", "MCP"),
    ("multicomp", "Multicomp Pro", "MCP"),
    ("amtek", "Amtek", "AMT"),
    ("wago", "Wago", "WAG"),
    ("harting", "Harting", "HAR"),
    ("hirose", "Hirose", "HIR"),
    ("samtec", "Samtec", "SAM"),
    ("raltron", "Raltron", "RA"),
]


def extract_pdf_data(stream_or_path: BinaryIO | bytes | Path, part_number: str | None = None) -> tuple[str, int, list[int]]:
    """Extract informative text from a PDF, intelligently filtering long documents."""
    if isinstance(stream_or_path, (bytes, bytearray)):
        if not stream_or_path:
            return "", 0, []
        stream = io.BytesIO(stream_or_path)
    elif isinstance(stream_or_path, Path):
        stream = open(stream_or_path, "rb")
    else:
        stream = stream_or_path

    reader = PdfReader(stream)
    total_pages = len(reader.pages)
    if total_pages == 0:
        return "", 0, []

    # If small PDF (<= 10 pages), read all pages
    if total_pages <= 10:
        pages_to_read = list(range(total_pages))
    else:
        # Prioritize key pages: always pages 0 and 1 (cover & main specifications)
        pages_set = {0, 1}
        pn_clean = re.sub(r"[^A-Za-z0-9]", "", part_number or "").casefold()
        keywords = (
            "ordering", "part number", "specifications", "electrical", "ratings",
            "absolute maximum", "dimension", "package", "characteristics", "marking",
            "features", "description", "table", "typical", "standard", "selection guide",
            "pin configuration", "terminal", "connection"
        )
        for i in range(total_pages):
            if len(pages_set) >= 16:
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

        pages_to_read = sorted(pages_set)[:16]

    extracted_parts = []
    analyzed_pages = []
    for idx in pages_to_read:
        try:
            page_txt = reader.pages[idx].extract_text()
            if page_txt:
                # Normalize Unicode characters (e.g. non-breaking spaces, symbols) and line breaks
                clean_txt = unicodedata.normalize("NFKC", page_txt).replace("\r\n", "\n").replace("\r", "\n")
                extracted_parts.append(f"--- PAGE {idx + 1} ---\n" + clean_txt)
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


COVER_TITLE_PATTERNS: list[tuple[str, str, str]] = [
    # RESISTOR
    ("RES", r"\b(?:thick\s+film\s+chip\s+resistors?|chip\s+resistors?|metal\s+film\s+resistors?|wirewound\s+resistors?|precision\s+(?:thin\s+film\s+)?resistors?|current\s+sens(?:ing|e)\s+resistors?|carbon\s+film\s+resistors?|resistor\s+networks?|fixed\s+resistors?)\b", "Direnç (RES)"),
    # CAPACITOR
    ("CAP", r"\b(?:ceramic\s+chip\s+capacitors?|multilayer\s+ceramic\s+capacitors?|surface\s+mount\s+ceramic\s+capacitors?|mlcc|electrolytic\s+capacitors?|aluminum\s+electrolytic|tantalum\s+capacitors?|film\s+capacitors?|supercapacitors?)\b", "Kapasitör (CAP)"),
    # DIODE
    ("DIO", r"\b(?:schottky\s+barrier\s+rectifier|schottky\s+diode|rectifier\s+diode|switching\s+diode|zener\s+diode|fast\s+recovery\s+rectifier|bridge\s+rectifier|transient\s+voltage\s+suppressor|tvs\s+diode|esd\s+protection\s+diode|plastic-encapsulate\s+diodes?|surface\s+mount\s+rectifiers?|general\s+purpose\s+rectifiers?)\b", "Diyot (DIO)"),
    # MOSFET
    ("FET", r"\b(?:power\s+mosfet|n-channel\s+(?:enhancement\s+mode\s+power\s+)?mosfet|p-channel\s+(?:enhancement\s+mode\s+power\s+)?mosfet|field\s+effect\s+transistor|dual\s+mosfet)\b", "MOSFET / FET"),
    # TRANSISTOR (BJT)
    ("TRA", r"\b(?:bipolar\s+(?:junction\s+)?transistor|npn\s+(?:general\s+purpose\s+)?transistor|pnp\s+(?:general\s+purpose\s+)?transistor|digital\s+transistor|darlington\s+transistor|small\s+signal\s+transistor|general\s+purpose\s+transistor)\b", "Transistör (TRA)"),
    # INDUCTOR
    ("IND", r"\b(?:power\s+choke|choke\s+coil|chip\s+inductor|power\s+inductor|smd\s+inductor|common\s+mode\s+choke|wire-wound\s+inductor|ferrite\s+inductor|shielded\s+power\s+inductor|molded\s+power\s+inductor|inductors?|chokes?)\b", "İndüktör / Bobin (IND)"),
    # VOLTAGE REGULATOR
    ("ICR", r"\b(?:voltage\s+regulator|low\s+dropout\s+(?:linear\s+)?regulator|ldo\s+regulator|linear\s+regulator|voltage\s+reference|shunt\s+regulator)\b", "Voltaj Regülatörü (ICR)"),
    # SMPS / POWER IC
    ("ICS", r"\b(?:step-down\s+(?:converter|regulator|switcher)|step-up\s+(?:converter|regulator)|buck\s+converter|boost\s+converter|pwm\s+controller|switching\s+regulator|smps\s+controller|power\s+switch|off-line\s+switcher)\b", "SMPS Güç Entegresi (ICS)"),
    # MICROCONTROLLER
    ("MCU", r"\b(?:microcontroller|32-bit\s+(?:arm\s+)?mcu|8-bit\s+mcu|flash\s+microcontroller|arm\s+cortex-m\d*|pic\s+microcontroller)\b", "Mikrodenetleyici (MCU)"),
    # OPTOCOUPLER
    ("OPT", r"\b(?:photocoupler|optocoupler|optoisolator|phototransistor\s+optocoupler)\b", "Optokuplör (OPT)"),
    # RELAY
    ("REL", r"\b(?:power\s+relay|miniature\s+relay|signal\s+relay|solid\s+state\s+relay|automotive\s+relay|pcb\s+relay|relays?)\b", "Röle (REL)"),
    # FUSE
    ("FUS", r"\b(?:resettable\s+fuse|ptc\s+fuse|smd\s+fuse|miniature\s+fuse|thermal\s+fuse|cartridge\s+fuse)\b", "Sigorta (FUS)"),
    # VARISTOR
    ("VAR", r"\b(?:metal\s+oxide\s+varistors?|siov\s+metal\s+oxide\s+varistors?|leaded\s+varistors?|varistors?|surge\s+absorbers?)\b", "Varistör (VAR)"),
    # THERMISTOR
    ("TER", r"\b(?:ntc\s+thermistor|ptc\s+thermistor|thermistor|temperature\s+sensor)\b", "Termistör (TER)"),
    # CRYSTAL / OSCILLATOR
    ("CRY", r"\b(?:(?:surface\s+mount\s+)?(?:microprocessor\s+)?crystals?|quartz\s+crystals?|crystal\s+units?|crystal\s+oscillators?|ceramic\s+resonators?|tcxo|vcxo|ocxo)\b", "Kristal / Osilatör (CRY)"),
    # SWITCH
    ("SWT", r"\b(?:tactile\s+switch|tact\s+switch|push\s+button\s+switch|dip\s+switch|slide\s+switch|rocker\s+switch|toggle\s+switch|rotary\s+switch|switches?)\b", "Anahtar / Switch (SWT)"),
    # FERRITE BEAD
    ("FER", r"\b(?:ferrite\s+bead|ferrite\s+chip\s+bead|chip\s+bead)\b", "Ferrit Boncuk (FER)"),
    # POTENTIOMETER
    ("POT", r"\b(?:potentiometer|trimmer\s+potentiometer|trimpot)\b", "Potansiyometre (POT)"),
    # BUZZER
    ("BUZ", r"\b(?:piezo\s+buzzer|magnetic\s+buzzer|piezo\s+sounder|active\s+buzzer|electromagnetic\s+(?:active\s+)?buzzer|acoustic\s+components?|buzzers?|蜂鸣器)\b", "Buzzer (BUZ)"),
    # BATTERY HOLDER / CLIP
    ("BAH", r"\b(?:battery\s+(?:holder|clip|retainer|contact)|coin\s+cell\s+(?:holder|retainer)|pil\s+(?:yuvası|tutucu|klipsi))\b", "Pil Yuvası (BAH)"),
    # TRANSFORMER
    ("TRF", r"\b(?:current\s+transformer|pulse\s+transformer|isolation\s+transformer|flyback\s+transformer|transformers?)\b", "Trafo (TRF)"),
    # AMPLIFIER
    ("AMP", r"\b(?:operational\s+amplifier|op\s+amp|audio\s+power\s+amplifier|instrumentation\s+amplifier)\b", "Amplifikatör (AMP)"),
    # LED
    ("LED", r"\b(?:surface\s+mount\s+led|light\s+emitting\s+diode|chip\s+led|smd\s+led)\b", "LED"),
    # CONNECTOR
    ("CON", r"\b(?:wire-to-board|board-to-board|pin\s+header|terminal\s+block|receptacle\s+housing|pcb\s+connector|crimp\s+housing|wafer\s+connector|box\s+header)\b", "Konnektör (CON)"),
    # ANTENNA
    ("ANT", r"\b(?:triple\s+band\s+(?:embedded\s+)?antenna|embedded\s+antenna|chip\s+antenna|pcb\s+antenna|rf\s+antenna|wifi\s+antenna|patch\s+antenna|dipole\s+antenna|antennas?)\b", "Anten (ANT)"),
]


def classify_material_type(text: str, part_number: str = "", filename: str = "") -> tuple[str, float, str]:
    """Determine component material type from Cover Page Title, PN, and weighted document content."""
    pn = part_number.strip().upper()
    text_lower = text.casefold()

    # ── 1. COVER PAGE / TITLE BLOCK MATCHING (Highest Authority) ──
    # Manufacturers prominently declare the component category on Page 1.
    cover_text = ""
    if "--- PAGE 1 ---" in text:
        p1 = text.split("--- PAGE 1 ---", 1)[1]
        cover_text = p1.split("--- PAGE 2 ---", 1)[0] if "--- PAGE 2 ---" in p1 else p1[:2500]
    else:
        cover_text = text[:2500]
    cover_lower = cover_text.casefold()

    for mat_code, pat, display_title in COVER_TITLE_PATTERNS:
        if re.search(pat, cover_lower):
            return mat_code, 0.99, f"Datasheet kapağındaki ürün başlığından tespit edildi ({display_title})."

    # ── 2. HIGH CONFIDENCE PART NUMBER PREFIXES ──
    # Resistors: Yageo RC/AC, Vishay CRCW, Panasonic ERJ, Walsin WR
    if re.match(r"^(?:RC|AC)\d{4}[A-Z]{2}-", pn) or re.match(r"^(?:RC\d{4}|CRCW\d{4}|ERJ\d|WR\d{2}|PYU-RC)", pn) or re.search(r"\bMC\s+0\.\d+W\b", text):
        return "RES", 0.98, "Part Number standart çip direnç kodlamasına uyuyor."
    # Capacitors: Samsung CL, Murata GRM, Yageo CC, Yageo AC automotive MLCC
    if re.match(r"^CL\d{2}[A-Z]\d{3}", pn) or re.match(r"^GRM\d{2}", pn) or re.match(r"^CC\d{4}", pn) or re.match(r"^AC\d{4}[A-Z0-9]*(?:BB|BN)\d+", pn):
        return "CAP", 0.98, "Part Number standart seramik kapasitör kodlamasına uyuyor."
    # Diodes: M1-M7 surface mount rectifiers, BAT, BZX, etc.
    if re.match(r"^(?:M[1-7]$|M[1-7]-|M[1-7]_|BAT|BZX|1N4148|1N400|1N58|SS\d{2}|MBR|SM712|SMAJ|SMBJ|US1|RS1)", pn) or (len(pn) == 2 and re.match(r"^M[1-7]$", pn)):
        return "DIO", 0.98, "Part Number standart diyot serisine uyuyor."
    # Buzzer
    if re.match(r"^(?:FUET|HNB|FST|CPI-|KPI-)", pn):
        return "BUZ", 0.98, "Part Number standart buzzer serisine uyuyor."
    # Battery Holder / Clip
    if re.match(r"^(?:L-KLS5-SP|BK-|BH-|BS-)", pn):
        return "BAH", 0.98, "Part Number pil yuvası/klipsi serisine uyuyor."
    # MOSFET
    if re.match(r"^(?:2N7002|BSS138|BSS84|SI23\d{2}|AO34\d{2}|IRF|DMN|DMP)", pn):
        return "FET", 0.95, "Part Number MOSFET kodlamasına uyuyor."
    # Transistor (BJT)
    if re.match(r"^(?:BC84|BC85|BC807|BC817|2N3904|2N3906|MMBT)", pn):
        return "TRA", 0.95, "Part Number BJT transistör kodlamasına uyuyor."
    # Voltage Regulator
    if re.match(r"^(?:LM1117|AMS1117|LM78|LM79|TL431|AP1117|LP29|TPS7)", pn):
        return "ICR", 0.95, "Part Number regülatör/LDO kodlamasına uyuyor."
    # SMPS IC
    if re.match(r"^(?:VIPER|LNK|TNY|TOP\d+|FSL|MP\d{4}|LM2596|LM2576)", pn):
        return "ICS", 0.95, "Part Number SMPS güç entegresi kodlamasına uyuyor."
    # Microcontroller
    if re.match(r"^(?:STM32|PIC\d|ATMEGA|ESP32|ATTINY)", pn):
        return "MCU", 0.98, "Part Number mikrodenetleyici kodlamasına uyuyor."
    # Optocoupler
    if re.match(r"^(?:PC817|EL817|LTV-817|MOC30|6N13|FOD817)", pn):
        return "OPT", 0.98, "Part Number optokuplör kodlamasına uyuyor."
    # Relay
    if re.match(r"^(?:HF\d|G5Q|SRD|G2R)", pn):
        return "REL", 0.95, "Part Number röle kodlamasına uyuyor."
    # Connector
    if re.match(r"^(?:22-01|0022|2695|5557|JST-|PH-|XH-|VH-|MOLEX|200G|254G|200W|254W|200H|254H|200D|254D|\d{3}[A-Z]\d?-\d{2})", pn):
        return "CON", 0.98, "Part Number standart konnektör/klemens serisine uyuyor."
    if re.match(r"^DG\d+.*", pn):
        return "CON", 0.98, "Part Number Degson klemens serisine uyuyor."
    # Fuse
    if re.match(r"^(?:MF-MSMF|MF-R|0451|0452|SMD\d{4}P)", pn):
        return "FUS", 0.95, "Part Number sigorta serisine uyuyor."
    # Varistor (Epcos B722*, etc.)
    if re.match(r"^(?:B722|\d{2}D\d{3}K|V\d{3}LA|S\d{2}K)", pn):
        return "VAR", 0.95, "Part Number varistör serisine uyuyor."
    # Thermistor
    if re.match(r"^(?:NTC|PTC|MF52|MF58|\d{1,2}D-[579])", pn):
        return "TER", 0.95, "Part Number termistör serisine uyuyor."
    # Crystal
    if re.match(r"^(?:RH\d{2,3}|ABS\d|ABM\d|ECS-|NX\d{4}|FA-\d|TSX-\d|7M\d|9C\d|CST|SIT\d)", pn):
        return "CRY", 0.95, "Part Number kristal serisine uyuyor."
    # Inductor
    if re.match(r"^(?:NR\d{4}|CD\d{2,3}|MSS\d{4}|SRR\d{4}|IHLP)", pn):
        return "IND", 0.95, "Part Number indüktör serisine uyuyor."

    # ── 3. WEIGHTED TEXT KEYWORD SCORING (Cover: 4x weight, Rest: 1x weight) ──
    scores: dict[str, float] = {
        "RES": 0.0, "CAP": 0.0, "DIO": 0.0, "FET": 0.0, "TRA": 0.0,
        "IND": 0.0, "ICR": 0.0, "ICS": 0.0, "MCU": 0.0, "OPT": 0.0,
        "LED": 0.0, "CON": 0.0, "REL": 0.0, "FUS": 0.0, "VAR": 0.0,
        "TER": 0.0, "CRY": 0.0, "SWT": 0.0, "FER": 0.0, "POT": 0.0,
        "BUZ": 0.0, "TRF": 0.0, "TRI": 0.0, "SCR": 0.0, "AMP": 0.0,
        "FLT": 0.0, "ANT": 0.0, "BAH": 0.0,
    }

    def _score(cat: str, pattern: str, pts: float):
        # 4x weight if on cover page, 1x if in body
        if re.search(pattern, cover_lower):
            scores[cat] += pts * 4.0
        elif re.search(pattern, text_lower):
            scores[cat] += pts * 1.0

    PARASITIC_RESISTANCE_RE = r"\b(?:coil|contact|thermal|junction|insulation|equivalent\s+series|esr|dc|dcr|series|lead|conductor|winding|ground)\s+resistance\b"
    PARASITIC_CAPACITANCE_RE = r"\b(?:load|shunt|junction|stray|parasitic|input|output|coupling|feedthrough|case)\s+capacitance\b"

    non_resistor_dominant = bool(
        re.search(r"\b(?:crystals?|oscillators?|resonators?|inductors?|chokes?|relays?|transformers?)\b", cover_lower)
        or re.search(r"\b(?:crystals?|oscillators?|resonators?|inductors?|chokes?|relays?|transformers?)\b", text_lower)
    )
    non_capacitor_dominant = bool(
        re.search(r"\b(?:crystals?|oscillators?|resonators?|diodes?|rectifiers?|tvs)\b", cover_lower)
        or re.search(r"\b(?:crystals?|oscillators?|resonators?|diodes?|rectifiers?|tvs)\b", text_lower)
    )

    # Resistor
    if any(k in text_lower or k in pn.casefold() for k in ("royal ohm", "royalohm", "royal-ohm")):
        scores["RES"] += 30.0
    if not non_resistor_dominant:
        if not re.search(PARASITIC_RESISTANCE_RE, text_lower):
            _score("RES", r"\b(?:resistors?|resistance|thick\s+film|chip\s+resistor|metal\s+film)\b", 12.0)
        else:
            _score("RES", r"\b(?:resistors?|thick\s+film|chip\s+resistor|metal\s+film)\b", 12.0)
        _score("RES", r"\b(?:kohm|mohm|kω|mω|ohm|ω)\b", 8.0)
    else:
        # Non-resistor component dominant (crystal, inductor, relay, etc.):
        # Suppress bare ohm and bare resistance scoring to avoid parasitic contamination
        _score("RES", r"\b(?:chip\s+resistors?|metal\s+film\s+resistors?|thick\s+film\s+chip\s+resistors?)\b", 12.0)

    _score("RES", r"\b(?:1/16w|1/10w|1/8w|1/4w|1/2w|1w|2w)\b", 6.0)

    # Capacitor
    if not non_capacitor_dominant:
        if not re.search(PARASITIC_CAPACITANCE_RE, text_lower) and not re.search(r"\bcj\s*\d+", text_lower):
            _score("CAP", r"\b(?:capacitors?|capacitance|mlcc)\b", 12.0)
        else:
            _score("CAP", r"\b(?:capacitors?|mlcc)\b", 12.0)
        _score("CAP", r"\b(?:dielectric|x7r|c0g|np0|x5r|y5v)\b", 8.0)
        _score("CAP", r"\b(?:picofarad|nanofarad|microfarad|pf|nf|µf|uf)\b", 6.0)
    else:
        # Non-capacitor component dominant (crystal, diode, etc.):
        # Suppress bare capacitance and bare pf/nf/uf scoring to avoid parasitic contamination
        _score("CAP", r"\b(?:ceramic\s+chip\s+capacitors?|multilayer\s+ceramic\s+capacitors?|mlcc)\b", 12.0)
        _score("CAP", r"\b(?:dielectric|x7r|c0g|np0|x5r|y5v)\b", 8.0)

    # Inductor
    _score("IND", r"\b(?:power\s+inductor|chip\s+inductor|smd\s+inductor|inductors?|inductance|choke|choke\s+coil)\b", 15.0)
    _score("IND", r"\b(?:microhenry|millihenry|µh|uh|nh|mh)\b", 10.0)
    _score("IND", r"\b(?:isat|irms|dcr|dc\s+resistance|saturation\s+current)\b", 6.0)

    # Diode
    _score("DIO", r"\b(?:schottky|zener|rectifiers?|diodes?|surface\s+mount\s+rectifiers?)\b", 15.0)
    _score("DIO", r"\b(?:forward\s+voltage|reverse\s+leakage|vf|ir)\b", 5.0)

    # MOSFET
    _score("FET", r"\b(?:mosfet|n-channel|p-channel|drain-source|rds\(on\))\b", 15.0)

    # Transistor (BJT)
    _score("TRA", r"\b(?:bipolar\s+transistor|npn|pnp|collector-emitter|hfe)\b", 15.0)

    # Voltage Regulator
    _score("ICR", r"\b(?:voltage\s+regulator|ldo|low\s+dropout|linear\s+regulator|voltage\s+reference)\b", 15.0)

    # SMPS IC
    _score("ICS", r"\b(?:buck\s+converter|boost\s+converter|step-down|step-up|switching\s+regulator|pwm\s+controller)\b", 15.0)

    # Microcontroller
    _score("MCU", r"\b(?:microcontroller|arm\s+cortex|flash\s+mcu|mcu|flash\s+memory)\b", 15.0)

    # Optocoupler
    _score("OPT", r"\b(?:optocoupler|photocoupler|optoisolator|isolation\s+voltage\s+5000v)\b", 15.0)

    # LED
    _score("LED", r"\b(?:light\s+emitting\s+diode|luminous\s+intensity|dominant\s+wavelength|chip\s+led)\b", 15.0)

    # Connector
    if not re.search(r"\b(?:antenna|relay|transformer|optocoupler)\b", cover_lower):
        _score("CON", r"\b(?:connector|header|housing|crimp|terminal\s+block|wire-to-board)\b", 15.0)
        if any(w in text_lower for w in ("molex", "jst", "phoenix contact", "wago", "ttaf", "amtek")):
            scores["CON"] += 12.0

    # Antenna
    _score("ANT", r"\b(?:antennas?|vswr|gain\s*(?:dbi)?|radiation\s+pattern|rf\s+data|peak\s+gain|average\s+gain|feed\s+point|polarization|triple\s+band)\b", 18.0)

    # Relay
    _score("REL", r"\b(?:relay|contact\s+rating|coil\s+voltage|spst|spdt)\b", 15.0)

    # Fuse
    _score("FUS", r"\b(?:resettable\s+fuse|ptc\s+fuse|tripping\s+current|hold\s+current|fuse)\b", 15.0)

    # Varistor
    _score("VAR", r"\b(?:varistors?|metal\s+oxide\s+varistors?|siov|clamping\s+voltage|max\s+continuous\s+voltage)\b", 18.0)

    # Thermistor
    _score("TER", r"\b(?:thermistor|ntc|ptc|b25/50|b25/85|zero\s+power\s+resistance)\b", 15.0)

    # Crystal
    _score("CRY", r"\b(?:quartz\s+crystal|crystal\s+unit|crystal\s+oscillator|ceramic\s+resonator|crystals?|resonators?|frequency\s+tolerance|load\s+capacitance)\b", 15.0)

    # Switch
    _score("SWT", r"\b(?:tactile\s+switch|tact\s+switch|push\s+button|slide\s+switch|dip\s+switch)\b", 15.0)

    # Ferrite Bead
    _score("FER", r"\b(?:ferrite\s+bead|ferrite\s+chip|impedance\s+at\s+100mhz)\b", 15.0)

    # Potentiometer
    _score("POT", r"\b(?:potentiometer|trimmer|trimpot|single\s+turn|multi\s+turn)\b", 15.0)

    # Buzzer
    _score("BUZ", r"\b(?:piezo\s+buzzer|magnetic\s+buzzer|active\s+buzzer|electromagnetic\s+buzzer|buzzers?|sound\s+pressure\s+level|resonant\s+frequency|oscillation\s+frequency|sound\s+output|蜂鸣器)\b", 18.0)

    # Battery Holder / Clip
    _score("BAH", r"\b(?:battery\s+(?:holder|clip|retainer|contact)|coin\s+cell\s+(?:holder|retainer)|pil\s+(?:yuvası|tutucu|klipsi))\b", 18.0)

    # Transformer
    _score("TRF", r"\b(?:transformer|turns\s+ratio|primary\s+inductance|isolation\s+voltage)\b", 15.0)

    # Amplifier
    _score("AMP", r"\b(?:operational\s+amplifier|op\s+amp|gain\s+bandwidth|slew\s+rate)\b", 15.0)

    best_type = max(scores, key=scores.get)
    best_score = scores[best_type]

    if best_score >= 8.0:
        conf = min(0.98, 0.5 + (best_score / 35.0))
        display_name = MATERIAL_TYPE_NAMES.get(best_type, best_type)
        return best_type, conf, f"Datasheet metnindeki ağırlıklı teknik anahtar kelimelerden tespit edildi ({display_name})."

    # ── 4. CONTEXTUAL FALLBACK ──
    pn_lower = pn.casefold()
    if any(k in pn_lower for k in ("res", "royal")):
        return "RES", 0.70, "Parça kodu veya üretici bilgisinden direnç olarak tespit edildi."
    if any(k in pn_lower for k in ("buz", "fuet", "hnb", "fst")):
        return "BUZ", 0.85, "Parça kodundan buzzer olarak tespit edildi."
    if any(k in pn_lower for k in ("kls5-sp", "battery", "coin-cell")):
        return "BAH", 0.85, "Parça kodundan pil yuvası/klipsi olarak tespit edildi."
    if any(k in pn_lower for k in ("b722", "varistor")):
        return "VAR", 0.85, "Parça kodundan varistör olarak tespit edildi."
    if any(k in pn_lower for k in ("cap", "uf", "nf", "pf", "mlcc")):
        return "CAP", 0.70, "Parça kodu veya birim bilgisinden kapasitör olarak tespit edildi."
    if any(k in pn_lower for k in ("dio", "bat", "bzx", "1n4")):
        return "DIO", 0.70, "Parça kodundan diyot olarak tespit edildi."
    if any(k in pn_lower for k in ("con", "molex", "jst", "hdr", "header", "term", "crimp")):
        return "CON", 0.70, "Parça kodu veya marka bilgisinden konnektör olarak tespit edildi."
    if any(k in pn_lower for k in ("ind", "choke")):
        return "IND", 0.70, "Parça kodundan bobin/indüktör olarak tespit edildi."

    # Look for known manufacturer in text
    if any(k in text_lower for k in ("molex", "jst", "ttaf", "amtek", "phoenix contact")):
        return "CON", 0.75, "Datasheet metnindeki konnektör üreticisinden tespit edildi (CON)."
    if any(k in text_lower for k in ("royal ohm", "royalohm")):
        return "RES", 0.85, "Datasheet metnindeki direnç üreticisinden tespit edildi (RES)."

    if "ohm" in text_lower and not non_resistor_dominant and not re.search(PARASITIC_RESISTANCE_RE, text_lower):
        return "RES", 0.35, "Datasheet metnindeki 'ohm' ifadesinden direnç olarak tahmin edildi."
    if any(w in text_lower for w in ("housing", "connector", "circuits", "molex")):
        return "CON", 0.35, "Datasheet metnindeki konnektör terimlerinden tahmin edildi."

    return "BİLİNMİYOR", 0.0, "Malzeme tipi kesin belirlenemedi."


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

    # 9. Degson Terminal Blocks: e.g. DG8HS-1.2-12P-17-100SH/AH, DG10HS-1.0-02P-17-00AH
    degson_m = re.match(r"^DG(\d+)HS?-([0-9.]+)-(\d+)P", pn)
    if degson_m:
        series_num, pitch_code, pin_count = degson_m.groups()
        specs["manufacturer"] = "Degson"
        specs["type_connector"] = "Wire to Wire Klemens"
        specs["pin_type"] = "Dişi"
        specs["mounting"] = "Cable"
        specs["pin_number"] = f"{int(pin_count)} pin"
        # Pitch code mapping: 1.2 in DG8HS series represents 8.2mm pitch
        if pitch_code in ("1.2", "8.2"):
            specs["pitch"] = "8.2mm"
        elif pitch_code in ("1.0", "10.0"):
            specs["pitch"] = "10.0mm"
        elif pitch_code in ("1.5", "15.0"):
            specs["pitch"] = "15.0mm"
        return specs

    # 10. Raltron Crystals: e.g. RH100-8.000-16-3030-TR-NS1
    ral_m = re.match(r"^RH(?:\d+)?-([0-9.]+)-([0-9.]+)", pn)
    if ral_m or pn.startswith("RALTRON"):
        specs["manufacturer"] = "Raltron"
        if ral_m:
            f_val, cl_val = ral_m.groups()
            if "." in f_val:
                f_val = f_val.rstrip("0").rstrip(".")
            specs["frequency_str"] = f"{f_val}MHz"
            if "." in cl_val:
                cl_val = cl_val.rstrip("0").rstrip(".")
            specs["load_capacitance"] = f"{cl_val} pF"
            specs["mounting"] = "SMD"
            if pn.startswith("RH100"):
                specs["package"] = "4-SMD"
        return specs

    return specs


def extract_pdf_tables(file_bytes: bytes) -> list[list[list[str]]]:
    """Extract all 2D tables from PDF bytes using pdfplumber without any external API or network calls."""
    tables: list[list[list[str]]] = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    for t in page_tables:
                        clean_table: list[list[str]] = []
                        for row in t:
                            if row:
                                clean_table.append([str(c).strip() if c is not None else "" for c in row])
                        if clean_table:
                            tables.append(clean_table)
    except Exception:
        pass
    return tables


def match_part_number_in_tables(file_bytes: bytes | None, part_number: str) -> dict[str, Any]:
    """Search 2D tables in PDF for exact row matching part_number and extract its column attributes.

    Provides deterministic, coordinate-accurate parameter extraction for multi-variant catalogs.
    """
    if not file_bytes or not part_number or len(part_number) < 2:
        return {}

    pn_clean = part_number.strip().upper()
    variant_specs: dict[str, Any] = {}

    try:
        tables = extract_pdf_tables(file_bytes)
        for table in tables:
            for row in table:
                # Find if any cell in this row matches the part number
                matched_col = -1
                for c_idx, cell in enumerate(row):
                    if cell.upper() == pn_clean or re.search(rf"\b{re.escape(pn_clean)}\b", cell, re.I):
                        matched_col = c_idx
                        break

                if matched_col >= 0:
                    for c_idx, cell in enumerate(row):
                        if c_idx == matched_col or not cell:
                            continue

                        # 1. Dimensions / Lengths (between 1.0 and 3000.0 mm)
                        num_m = re.match(r"^(\d+(?:\.\d+)?)(?:\s*mm)?$", cell)
                        if num_m:
                            val = float(num_m.group(1))
                            dash_suffix = pn_clean.split("-")[-1] if "-" in pn_clean else ""
                            if 1.0 <= val <= 3000.0 and (not dash_suffix.isdigit() or val != float(dash_suffix)):
                                if "cable_length" not in variant_specs:
                                    variant_specs["cable_length"] = f"{int(val) if val == int(val) else val}mm"

                        # 2. Connector Types
                        conn_m = re.search(r"\b(MHF(?:-TYPE|4-TYPE|\s*4)?|SMA|U\.FL|IPEX|MCX|MMCX|Type-C|USB|RJ45|JST|Molex)\b", cell, re.I)
                        if conn_m and "mating_connector" not in variant_specs:
                            variant_specs["mating_connector"] = conn_m.group(1).upper().replace("-TYPE", "")

                        # 3. Pin Counts
                        pins_m = re.search(r"\b(\d{1,3})\s*(?:pin|circuits?|pos|poles?|way)\b", cell, re.I)
                        if pins_m and "pin_number" not in variant_specs:
                            variant_specs["pin_number"] = f"{pins_m.group(1)} pin"

                        # 4. Pitch
                        pitch_m = re.search(r"\b(1\.00?|1\.25|1\.27|1\.50?|2\.00?|2\.50?|2\.54|3\.00?|3\.81|3\.96|4\.20?|5\.00?|5\.08)\s*mm\b", cell, re.I)
                        if pitch_m and "pitch" not in variant_specs:
                            variant_specs["pitch"] = pitch_m.group(0).lower()

                    if variant_specs:
                        return variant_specs
    except Exception:
        pass

    return variant_specs


def _extract_part_variant_table_row(text: str, part_number: str) -> dict[str, Any]:
    """Find the specific row or table entry corresponding to *part_number* in multi-variant datasheets.

    Extracts variant-specific dimensions, cable lengths, connectors, pin counts, voltages, etc.
    This works universally for antennas, connectors, diodes, passives, and IC families.
    """
    if not part_number or len(part_number) < 2:
        return {}
    variant_specs: dict[str, Any] = {}
    pn_clean = part_number.strip()

    for line in text.splitlines():
        line_s = line.strip()
        if not line_s:
            continue
        # Does this line contain the part number as a distinct token?
        m = re.search(rf"(?:^|[\s|;,]){re.escape(pn_clean)}(?:[\s|;,]|$)", line_s)
        if m:
            after_pn = line_s[m.end():].strip()
            tokens = after_pn.split()

            # Find dimensions / lengths (between 1.0 and 3000.0)
            dims = []
            for t in tokens:
                num_m = re.match(r"^(\d+(?:\.\d+)?)(?:mm)?$", t)
                if num_m:
                    val = float(num_m.group(1))
                    dash_suffix = pn_clean.split("-")[-1] if "-" in pn_clean else ""
                    if 1.0 <= val <= 3000.0 and (not dash_suffix.isdigit() or val != float(dash_suffix)):
                        dims.append(val)

            if dims and "cable_length" not in variant_specs:
                val = dims[0]
                variant_specs["cable_length"] = f"{int(val) if val == int(val) else val}mm"

            # Connectors (MHF, MHF4, SMA, U.FL, IPEX, etc.)
            conn_m = re.search(r"\b(MHF(?:-TYPE|4-TYPE|\s*4)?|SMA|U\.FL|IPEX|MCX|MMCX|Type-C|USB|RJ45|JST|Molex)\b", line_s, re.I)
            if conn_m and "mating_connector" not in variant_specs:
                variant_specs["mating_connector"] = conn_m.group(1).upper().replace("-TYPE", "")

            # Pin count
            pins_m = re.search(r"\b(\d{1,3})\s*(?:pin|circuits?|pos|poles?|way)\b", line_s, re.I)
            if pins_m and "pin_number" not in variant_specs:
                variant_specs["pin_number"] = f"{pins_m.group(1)} pin"

    return variant_specs


def extract_component_specs(
    material_type: str,
    text: str,
    part_number: str = "",
    is_catalog: bool = False,
    file_bytes: bytes | None = None,
) -> dict[str, Any]:
    """Extract physical and electrical specifications from decoded PN + 2D tables + full text."""
    # Support alternate argument orderings: (text, material_type, part_number) or (text, part_number, material_type)
    if material_type in MATERIAL_TYPE_NAMES:
        pass
    elif text in MATERIAL_TYPE_NAMES:
        material_type, text = text, material_type
    elif part_number in MATERIAL_TYPE_NAMES:
        material_type, text, part_number = part_number, material_type, text
    elif len(material_type) > len(text) and len(text) <= 4:
        material_type, text = text, material_type

    specs = decode_manufacturer_pn(part_number)
    pn = part_number.strip().upper()
    pn_lower = pn.lower()
    text_lower = text.casefold()
    is_cat = is_catalog or is_catalog_document_name(part_number)

    # 1. Native 2D Table Extraction via pdfplumber
    if file_bytes:
        table_specs = match_part_number_in_tables(file_bytes, part_number)
        for k, v in table_specs.items():
            if k not in specs:
                specs[k] = v

    # 2. Universal Line-by-Line Regex Fallback for multi-part catalogs
    variant_specs = _extract_part_variant_table_row(text, part_number)
    for k, v in variant_specs.items():
        if k not in specs:
            specs[k] = v

    # Mounting
    if "mounting" not in specs:
        if material_type == "CON":
            if any(k in text_lower for k in ("smd", "smt", "surface mount")):
                specs["mounting"] = "SMD"
            else:
                specs["mounting"] = "TH"
        elif any(k in text_lower for k in ("through hole", "through-hole", "tht", "dip", "radial", "leaded", "wave soldering", "terminal pins")):
            specs["mounting"] = "TH"
        elif any(k in text_lower for k in ("smd", "smt", "surface mount", "chip")):
            specs["mounting"] = "SMD"
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
            elif re.search(r"\b(?:sma|do-214ac)\b", text_lower):
                specs["package"] = "SMA"
            elif re.search(r"\b(?:smb|do-214aa)\b", text_lower):
                specs["package"] = "SMB"
            elif "to-220" in text_lower:
                specs["package"] = "TO-220"
            elif "to-92" in text_lower:
                specs["package"] = "TO-92"

    # Manufacturer
    if "manufacturer" not in specs:
        for kw, name, code in KNOWN_MANUFACTURERS:
            if re.search(rf"\b{re.escape(kw)}\b", text_lower):
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
        # Standard M1-M7 surface-mount rectifier series (SMA, 1A)
        m_pn = re.search(r"(?:^|[^A-Za-z0-9])M([1-7])(?:[^A-Za-z0-9]|$)", pn)
        if not m_pn:
            m_pn = re.search(r"\bM([1-7])\b", text)
        if m_pn:
            m_digit = m_pn.group(1)
            v_map = {"1": "50V", "2": "100V", "3": "200V", "4": "400V", "5": "600V", "6": "800V", "7": "1000V"}
            specs["voltage"] = v_map[m_digit]
            specs["current"] = "1A"
            specs["value_diode"] = f"M{m_digit}"
            if "package" not in specs:
                specs["package"] = "SMA"
            if "mounting" not in specs:
                specs["mounting"] = "SMD"

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
            elif "rectifier" in text_lower or m_pn:
                specs["type_diode"] = "Standart Doğrultucu"

        # Diode Value / PN
        if "value_diode" not in specs:
            specs["value_diode"] = part_number.strip()

        # Automotive AEC-Q101 qualification
        if "aec-q101" in text_lower or "aec q101" in text_lower:
            specs["special_def"] = "Automotive AEC-Q101"

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

    elif material_type == "IND":
        # Inductance
        if "inductance_str" not in specs and not is_cat:
            ind_dec = _extract_inductance(text_lower)
            if ind_dec is not None:
                specs["inductance_str"] = _format_ind(ind_dec)

        # Size inductor / dimensions
        if "size_inductor" not in specs:
            size_m = re.search(r"\b(\d+(?:\.\d+)?\s*[x*]\s*\d+(?:\.\d+)?(?:\s*[x*]\s*\d+(?:\.\d+)?)?\s*mm)\b", text_lower)
            if size_m:
                specs["size_inductor"] = size_m.group(1).replace(" ", "")

    elif material_type == "LED":
        # Color
        if "led_color" not in specs:
            color_map = [
                ("red-blue", "Red-Blue"),
                ("ultra orange red", "Ultra Orange Red"),
                ("orange", "Orange"),
                ("amber", "Amber"),
                ("yellow-green", "Yellow-Green"),
                ("yellow", "Yellow"),
                ("green", "Green"),
                ("blue", "Blue"),
                ("white", "White"),
                ("red", "Red"),
                ("kırmızı", "Red"),
                ("yeşil", "Green"),
                ("mavi", "Blue"),
                ("sarı", "Yellow"),
                ("beyaz", "White"),
                ("rgb", "RGB"),
                ("infrared", "IR"),
                ("ir", "IR"),
            ]
            for pat, col in color_map:
                if re.search(rf"\b{pat}\b", text_lower):
                    specs["led_color"] = col
                    break

    elif material_type == "FUS":
        # Type Fuse
        if "type_fuse" not in specs:
            if any(k in text_lower for k in ("resettable", "ptc")):
                specs["type_fuse"] = "Resettable Fuse"
            elif "thermal" in text_lower:
                specs["type_fuse"] = "Thermal Fuse"
            elif "holder" in text_lower:
                specs["type_fuse"] = "Fuse Holder"
            else:
                specs["type_fuse"] = "Fuses"
        # Fuse Current
        if "fuse_current" not in specs and not is_cat:
            cur_m = re.search(r"(?:rated\s+current|current\s+rating|ihold|hold\s+current)[:\s=]+(\d+(?:\.\d+)?\s*(?:ma|a))\b", text_lower)
            if not cur_m:
                cur_m = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:a|ma)\b", text_lower)
            if cur_m:
                specs["fuse_current"] = cur_m.group(0).replace(" ", "").upper()

    elif material_type == "VAR":
        # Check part number coding for disc diameter (e.g. B72207 -> 7mm, S07 -> 7mm, B72210 -> 10mm, B72214 -> 14mm)
        pn_var_m = re.search(r"(?:B722|S)0?(\d{1,2})", pn)
        if pn_var_m and "varistor_size" not in specs:
            specs["varistor_size"] = f"{int(pn_var_m.group(1))}mm"

        # AC Voltage (Check 3-digit EIA voltage in PN first to avoid matching catalog range like 11...1100 VRMS)
        if "voltage_ac" not in specs:
            var_pn_v = re.search(r"(?:S0?|K)(\d{3})", pn)
            v_var_map = {
                "110": "11VAC", "140": "14VAC", "170": "17VAC", "200": "20VAC", "250": "25VAC",
                "300": "30VAC", "350": "35VAC", "400": "40VAC", "500": "50VAC", "600": "60VAC",
                "750": "75VAC", "950": "95VAC", "131": "130VAC", "141": "140VAC", "151": "150VAC",
                "171": "175VAC", "211": "210VAC", "231": "230VAC", "251": "250VAC", "271": "275VAC",
                "301": "300VAC", "321": "320VAC", "351": "350VAC", "381": "385VAC", "421": "420VAC",
                "461": "460VAC", "511": "510VAC", "551": "550VAC", "681": "680VAC",
            }
            if var_pn_v and var_pn_v.group(1) in v_var_map:
                specs["voltage_ac"] = v_var_map[var_pn_v.group(1)]
            else:
                vac_m = re.search(r"\b(\d+)\s*(?:vac|v\s*ac|vrms)\b", text_lower)
                if vac_m:
                    specs["voltage_ac"] = f"{vac_m.group(1)}VAC"
                elif "271" in part_number or "275" in text_lower:
                    specs["voltage_ac"] = "275VAC"

        # Varistor Size
        if "varistor_size" not in specs:
            size_m = re.search(r"(?:size|diameter|çap)[:\s=]+(\d+)\s*mm\b", text_lower)
            if not size_m:
                size_m = re.search(r"\b(\d{1,2})d\b", text_lower)
            if size_m:
                specs["varistor_size"] = f"{size_m.group(1)}mm"

        # Mounting
        if "mounting" not in specs:
            if any(k in text_lower for k in ("leaded", "radial", "through hole", "tht")):
                specs["mounting"] = "TH"

    elif material_type == "TER":
        # Type Thermistor
        if "type_thermistor" not in specs:
            if "ptc" in text_lower:
                specs["type_thermistor"] = "PTC"
            else:
                specs["type_thermistor"] = "NTC"
        # Resistance
        if "resistance_str" not in specs and not is_cat:
            res_dec = _extract_resistance(text_lower)
            if res_dec is not None:
                specs["resistance_str"] = _format_res(res_dec)
        # Tolerance
        if "tolerance" not in specs and not is_cat:
            tol_dec = _extract_tolerance(text_lower)
            if tol_dec is not None:
                specs["tolerance"] = f"%{int(tol_dec)}"

    elif material_type == "CRY":
        # Frequency / Value
        if "frequency_str" in specs:
            fm = re.match(r"^(\d+(?:\.\d+)?)\s*(mhz|khz)$", str(specs["frequency_str"]), re.I)
            if fm:
                v = fm.group(1)
                if "." in v:
                    v = v.rstrip("0").rstrip(".")
                u = "MHz" if fm.group(2).lower() == "mhz" else "kHz"
                specs["frequency_str"] = f"{v}{u}"
        else:
            freq_m = re.search(r"\b(\d+(?:\.\d+)?)\s*(mhz|khz)\b", text_lower)
            if freq_m:
                val_num = freq_m.group(1)
                if "." in val_num:
                    val_num = val_num.rstrip("0").rstrip(".")
                unit = "MHz" if freq_m.group(2) == "mhz" else "kHz"
                specs["frequency_str"] = f"{val_num}{unit}"

        # Load Capacitance
        if "load_capacitance" not in specs:
            cap_m = re.search(r"\b(?:load\s+capacitance|load\s+cap|cl)[:\s=]+(\d+(?:\.\d+)?)\s*pf\b", text_lower)
            if cap_m:
                cap_val = cap_m.group(1)
                if "." in cap_val:
                    cap_val = cap_val.rstrip("0").rstrip(".")
                specs["load_capacitance"] = f"{cap_val} pF"

        # ESR (Equivalent Series Resistance)
        if "esr" not in specs:
            esr_m = re.search(r"\b(?:equivalent\s+series\s+resistance|esr)[:\s=]+(\d+(?:\.\d+)?)\s*(?:ohm|ω|q|Ω)\b", text_lower)
            if esr_m:
                esr_val = esr_m.group(1)
                if "." in esr_val:
                    esr_val = esr_val.rstrip("0").rstrip(".")
                specs["esr"] = f"{esr_val}Ohm"

        # Package
        if specs.get("package") not in ("4-SMD", "2-SMD", "HC-49/SMD", "HC-49S/SMD", "HC-49/U"):
            if (
                any(k in text_lower for k in ("4-smd", "4 smd", "4-pad", "4 pad", "3225", "3.2x2.5", "3.2 x 2.5", "5032", "5.0x3.2", "2016", "2.0x1.6", "7050", "7.0x5.0"))
                or re.search(r"\b4x1(?:\.\d+)?\b", text_lower)
                or re.search(r"\b4\s*pins?\b", text_lower)
                or pn.startswith("RH100")
            ):
                specs["package"] = "4-SMD"
            elif any(k in text_lower for k in ("hc-49/smd", "hc-49smd", "hc-49s/smd", "hc-49us/smd")):
                specs["package"] = "HC-49/SMD"
            elif any(k in text_lower for k in ("2-smd", "2 smd", "2-pad", "2 pad")) or re.search(r"\b2\s*pins?\b", text_lower):
                specs["package"] = "2-SMD"
            elif any(k in text_lower for k in ("hc-49/u", "hc-49u", "hc-49s", "hc-49us")):
                specs["package"] = "HC-49/SMD" if specs.get("mounting") == "SMD" else "HC-49/U"

        # Mounting
        if str(specs.get("package", "")).endswith("SMD"):
            specs.setdefault("mounting", "SMD")
        if "mounting" not in specs or specs["mounting"] != "SMD":
            if any(k in text_lower for k in ("surface mount", "smd", "smt")):
                specs["mounting"] = "SMD"
            elif any(k in text_lower for k in ("through hole", "through-hole", "tht")):
                specs["mounting"] = "TH"

        # Manufacturer
        if "manufacturer" not in specs:
            if "raltron" in text_lower or re.match(r"^RH(?:\d+)?-[0-9.]+", pn) or pn.startswith("RALTRON"):
                specs["manufacturer"] = "Raltron"

    elif material_type == "POT":
        # Resistance
        if "resistance_str" not in specs and not is_cat:
            res_dec = _extract_resistance(text_lower)
            if res_dec is not None:
                specs["resistance_str"] = _format_res(res_dec)
        # Turns
        if "pot_turns" not in specs:
            if any(k in text_lower for k in ("multi", "10 turn", "10 tur", "10-turn")):
                specs["pot_turns"] = "Multi"
            elif "slide" in text_lower:
                specs["pot_turns"] = "Slide"
            else:
                specs["pot_turns"] = "1 Tur"

    elif material_type == "REL":
        # Coil Voltage
        if "relay_coil_voltage" not in specs:
            cv_m = re.search(r"(?:coil|bobin)[:\s=]+(\d+)\s*v(?:dc)?\b", text_lower)
            if cv_m:
                specs["relay_coil_voltage"] = f"Coil:{cv_m.group(1)}VDC"
        # Contact Voltage
        if "relay_contact_voltage" not in specs:
            cont_v = re.search(r"\b(120|125|220|230|250)\s*vac\b", text_lower)
            if cont_v:
                specs["relay_contact_voltage"] = f"{cont_v.group(1)}VAC"
        # Contact Current
        if "relay_contact_current" not in specs:
            cont_i = re.search(r"\b(\d+)\s*a\b", text_lower)
            if cont_i:
                specs["relay_contact_current"] = f"{cont_i.group(1)}A"
        # Contact Type
        if "relay_contact_type" not in specs:
            if any(k in text_lower for k in ("spst-no", "1 form a", "normally open")):
                specs["relay_contact_type"] = "NO"
            elif any(k in text_lower for k in ("spst-nc", "1 form b", "normally closed")):
                specs["relay_contact_type"] = "NC"
            elif any(k in text_lower for k in ("spdt", "1 form c", "changeover")):
                specs["relay_contact_type"] = "NO/NC"

    elif material_type == "SWT":
        if "type_switch" not in specs:
            if any(k in text_lower for k in ("tact", "tactile")):
                specs["type_switch"] = "Tact Switch"
            elif "push" in text_lower:
                specs["type_switch"] = "Push Button"
            elif "slide" in text_lower:
                specs["type_switch"] = "Slide Switch"
            elif "toggle" in text_lower:
                specs["type_switch"] = "Toggle Switch"
            elif "dip" in text_lower:
                specs["type_switch"] = "DIP Switch"

    elif material_type == "FER":
        if "ferrite_type" not in specs:
            if "toroid" in text_lower:
                specs["ferrite_type"] = "Toroid"
            else:
                specs["ferrite_type"] = "Axial"

    elif material_type == "BUZ":
        # Voltage (3V, 5V, 6V, 9V, 12V)
        if "buzzer_voltage" not in specs:
            bv_m = (
                re.search(r"(?:rated\s*voltage|额定电压)\s*(?:\([^\)]+\))?\s*(?:vo-p|vdc|v)?\s*[:\s]*(\d+(?:\.\d+)?)", text_lower)
                or re.search(r"(?:operating\s*voltage)\s*(?:\([^\)]+\))?\s*(?:vo-p|vdc|v)?\s*[:\s]*(\d+(?:\.\d+)?)", text_lower)
                or re.search(r"\b(\d+(?:\.\d+)?)\s*v(?:dc|o-p)?\s+rated\b", text_lower)
                or re.search(r"\b(3|5|6|9|12)\s*v(?:dc|o-p)?\b", text_lower)
            )
            if bv_m:
                val = float(bv_m.group(1))
                if val in (3, 5, 6, 9, 12):
                    specs["buzzer_voltage"] = f"{int(val)}V"

        # Frequency (Hz / kHz -> standard dictionary values e.g. 4kHz, 2.7kHz)
        if "buzzer_frequency" not in specs:
            bf_hz_m = (
                re.search(r"(?:oscillation\s*frequency|resonant\s*frequency|rated\s*frequency|共振频率)\s*(?:[^\n\r:0-9]{0,15})[:\s]*(\d{4})", text_lower)
                or re.search(r"\b(\d{4})\s*hz\b", text_lower)
            )
            if bf_hz_m:
                hz = int(bf_hz_m.group(1))
                hz_map = {
                    1000: "1kHz", 2000: "2kHz", 2300: "2.3kHz", 2400: "2.4kHz", 2500: "2.5kHz",
                    2700: "2.7kHz", 2730: "2.73kHz", 2900: "2.9kHz", 3000: "3kHz", 3100: "3.10kHz",
                    4000: "4kHz"
                }
                closest = min(hz_map.keys(), key=lambda k: abs(k - hz))
                if abs(closest - hz) <= 150:
                    specs["buzzer_frequency"] = hz_map[closest]
                else:
                    specs["buzzer_frequency"] = f"{round(hz / 1000.0, 1)}kHz"
            else:
                bf_khz_m = re.search(r"\b(\d+(?:\.\d+)?)\s*khz\b", text_lower)
                if bf_khz_m:
                    specs["buzzer_frequency"] = f"{bf_khz_m.group(1)}kHz"

        # Sound Level (dB -> e.g. 75dB, 85dB)
        if "sound_level" not in specs:
            spl_m = (
                re.search(r"\b(5[0-9]|[6-9][0-9]|1[0-2][0-9])\s*db\b", text_lower)
                or re.search(r"(?:sound\s*(?:pressure\s*level|output)|声压值|soundpressurelevel)\s*(?:at\s*10\s*cm)?\s*(?:\([^\)]+\))?\s*(?:db)?\s*(?:[^\n\r:0-9]{0,15})[:\s]*(?:min\.|≥|>=)?\s*([5-9]\d|1[0-2]\d)", text_lower)
            )
            if spl_m:
                db_val = int(spl_m.group(1))
                db_choices = [50, 60, 70, 75, 80, 83, 85, 87, 90]
                closest_db = min(db_choices, key=lambda d: abs(d - db_val))
                if abs(closest_db - db_val) <= 5:
                    specs["sound_level"] = f"{closest_db}dB"
                else:
                    specs["sound_level"] = f"{db_val}dB"

        # Buzzer Size
        if "buzzer_size" not in specs:
            if "hnb09" in pn.lower() or "9mm" in text_lower:
                specs["buzzer_size"] = "9mm"
            elif "5020" in pn or "fst-5020" in text_lower or "5mm" in text_lower:
                specs["buzzer_size"] = "5mm"
            elif "1514" in pn or "15.4" in text_lower:
                specs["buzzer_size"] = "14mm"
            else:
                sz_m = re.search(r"(?:dimensions|size)\s*[:\s]*ø?\s*(\d+(?:\.\d+)?)\s*mm", text_lower)
                if sz_m:
                    specs["buzzer_size"] = f"{sz_m.group(1)}mm"

        # Mounting
        if "mounting" not in specs:
            if any(k in text_lower for k in ("pin", "pins", "wave soldering", "through hole", "tht")):
                specs["mounting"] = "TH"
            elif any(k in text_lower for k in ("smd", "smt", "reflow", "surface mount", "reel")):
                specs["mounting"] = "SMD"

    elif material_type == "BAH":
        if "type_battery_holder" not in specs:
            if "l-kls5-sp" in pn.lower() or "clip" in text_lower or "retainer" in text_lower:
                specs["type_battery_holder"] = "Coin Cell Retainer"
            elif "holder" in text_lower or "yuvası" in text_lower:
                specs["type_battery_holder"] = "CR2032 Pil Yuvası"
            else:
                specs["type_battery_holder"] = "Pil Tutucu"
        if "mounting" not in specs:
            specs["mounting"] = "SMD" if "smd" in text_lower or "surface" in text_lower or "l-kls5" in pn.lower() else "TH"
        if "manufacturer" not in specs:
            if "kls" in pn.lower() or "kls" in text_lower:
                specs["manufacturer"] = "Kls"
                specs["manufacturer_code"] = "KL"

    elif material_type == "TRF":
        if "type_transformer" not in specs:
            if "isolation" in text_lower:
                specs["type_transformer"] = "Isolation Transformer"
            elif "smps" in text_lower:
                specs["type_transformer"] = "SMPS Transformer"
            elif "choke" in text_lower:
                specs["type_transformer"] = "Choke Coil"
            else:
                specs["type_transformer"] = "Standard Transformer"

    elif material_type == "ANT":
        # Mounting: TH (Board/PCB Antenna), SMD (Chip), Cable (External Cable)
        if any(k in text_lower for k in ("pcb antenna", "board antenna", "embedded antenna", "adhesive backing on the pcb")):
            specs["mounting"] = "TH"
        elif any(k in text_lower for k in ("smd", "surface mount", "chip antenna")):
            specs["mounting"] = "SMD"
        elif any(k in text_lower for k in ("through hole", "tht", "dip")):
            specs["mounting"] = "TH"
        elif any(k in text_lower for k in ("cable", "coaxial", "micro-coaxial", "mating connector", "adhesive")):
            specs["mounting"] = "Cable"
        else:
            specs["mounting"] = "TH"

        # Protocol / Application
        if "wifi" in text_lower or "wi-fi" in text_lower:
            specs["protocol"] = "WiFi"
        elif "bluetooth" in text_lower or "zigbee" in text_lower:
            specs["protocol"] = "BLE"
        elif "gps" in text_lower or "gnss" in text_lower:
            specs["protocol"] = "GPS"
        elif "lte" in text_lower or "5g" in text_lower or "cellular" in text_lower:
            specs["protocol"] = "Cellular"

        # Type Antenna: 1:Chip, 2:Wire, 3:RF
        if "chip" in text_lower:
            specs["type_antenna"] = "Chip"
        elif "wire" in text_lower:
            specs["type_antenna"] = "Wire"
        else:
            specs["type_antenna"] = "RF"

    # Common for ICs and discrete parts: set part number as value
    if material_type in ("OPT", "ICS", "MCU", "TRI", "SCR", "AMP", "FLT", "ICL", "ICM", "ICC", "ICD", "ICB"):
        specs["value_pn"] = part_number.strip()

    return specs


CORE_ATTRIBUTES_BY_TYPE: dict[str, set[str]] = {
    "RES": {"mounting", "package", "value", "tolerance", "wattage", "compositionresistors", "specialdefinitionresistors", "q"},
    "CAP": {"mounting", "package", "value", "voltage", "tolerance", "temperature", "specialdefinitioncapacitors", "q"},
    "DIO": {"mounting", "package", "typediode", "value", "specialdefinitionall", "q"},
    "CON": {"mounting", "typeconnector", "pintypeofconnector", "pitchconnector", "pitch", "pinnumber", "angle", "q"},
    "IND": {"mounting", "package", "value", "sizeinductor", "q"},
    "FET": {"mounting", "package", "value", "fetchanneltype", "q"},
    "TRA": {"mounting", "package", "value", "typetransistor", "specialdefinitiontransistor", "q"},
    "ICR": {"mounting", "package", "typeregulator", "regulatoroutputvoltage", "regulatorcurrent", "voltage", "q"},
    "ICS": {"mounting", "package", "value", "q"},
    "MCU": {"mounting", "package", "value", "q"},
    "LED": {"mounting", "package", "ledcolor", "q"},
    "OPT": {"mounting", "package", "value", "q"},
    "REL": {"mounting", "relaycoilvoltage", "relaycontactvoltage", "relaycontactcurrent", "relaycontacttype", "q"},
    "FUS": {"mounting", "package", "typefuse", "fusecurrent", "q"},
    "VAR": {"mounting", "package", "varistorsize", "voltageac", "q"},
    "TER": {"mounting", "package", "typethermistor", "value", "tolerance", "q"},
    "CRY": {"mounting", "value", "q"},
    "POT": {"mounting", "value", "potturnnumber", "wattage", "q"},
    "SWT": {"mounting", "typeswitch", "q"},
    "FER": {"ferritetypes", "q"},
    "BUZ": {"mounting", "buzzerratedvoltage", "buzzerfrequency", "soundlevelbuzzer", "q"},
    "TRF": {"mounting", "typetransformer", "q"},
    "TRI": {"mounting", "package", "value", "q"},
    "SCR": {"mounting", "package", "value", "q"},
    "FLT": {"mounting", "package", "filtertype", "value", "q"},
    "AMP": {"mounting", "package", "q"},
    "ANT": {"mounting", "typeantenna", "q"},
    "BAH": {"mounting", "typebatteryholder", "q"},
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

        # Sequential ERP Counter (e.g. Serial Antenna Number, Serial Connector Number, Serial Crystal Number, etc.)
        # These are internal sequential counters in DEKA ERP (e.g. 000023). They do not exist in manufacturer datasheets.
        if "serial" in attr_norm:
            counter_val = specs.get("serial_counter") or "0000XX"
            mappings.append(PropertyMapping(
                attribute_name=attr_display_name,
                sort_no=sort_no,
                extracted_value="DEKA ERP Dahili Sıra Sayacı (Örn: 000023)",
                matched_erp_value="Otomatik Sayaç",
                erp_code=counter_val,
                prefix=entries[0].prefix or "",
                status="ERP SAYAÇ (Dahili)",
            ))
            code_tokens.append(counter_val)
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
                elif e.value == "C" and target in ("C", "CABLE"):
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
                # 1st pass: exact word boundary match avoiding array multipliers like 2x, 4x
                for e in entries:
                    desc = e.value_description or ""
                    if re.search(r"\b\d+x", desc, re.I):
                        continue
                    if re.search(rf"\b{re.escape(pkg)}\b", desc, re.I):
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
                # 2nd pass: fallback
                if not matched_entry:
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

            elif material_type == "IND":
                target_val = specs.get("inductance_str", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        i_entry = _extract_inductance(e.value_description.casefold())
                        i_target = _extract_inductance(target_val.casefold())
                        if i_entry is not None and i_target is not None and i_entry == i_target:
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                        elif target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "YAKIN EŞLEŞTİ"
                            break

            elif material_type in ("TER", "POT"):
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

            elif material_type == "CRY":
                target_val = specs.get("frequency_str", "")
                extracted_val_str = target_val
                if target_val:
                    for e in entries:
                        if target_val.casefold() in e.value_description.casefold():
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
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

            else:
                target_val = specs.get("value_pn", "")
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
            target_mfr_code = specs.get("manufacturer_code", "")
            extracted_val_str = target_mfr
            if target_mfr or target_mfr_code:
                # Pass 1: exact code match (e.g. 'TC' == e.value)
                if target_mfr_code:
                    for e in entries:
                        if target_mfr_code.upper() == e.value.strip().upper():
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                # Pass 2: exact description match (e.g. 'TE Connectivity' == 'TE Connectivity')
                if not matched_entry and target_mfr:
                    for e in entries:
                        if target_mfr.casefold() == e.value_description.strip().rstrip(",").casefold():
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                # Pass 3: word boundary match (e.g. '\bTE Connectivity\b' or '\bSamsung\b')
                if not matched_entry and target_mfr:
                    for e in entries:
                        if re.search(rf"\b{re.escape(target_mfr)}\b", e.value_description, re.IGNORECASE):
                            matched_entry = e
                            status = "TAM EŞLEŞTİ"
                            break
                # Pass 4: substring match ONLY if length >= 4 (to avoid 2-letter false matches like 'TE' matching 'Qualitek')
                if not matched_entry and target_mfr and len(target_mfr) >= 4:
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

        # 12. DIODE TYPE
        elif "typediode" in attr_norm:
            td = specs.get("type_diode", "")
            extracted_val_str = td
            if td:
                for e in entries:
                    if td.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 13. FET CHANNEL TYPE
        elif "fetchannel" in attr_norm:
            fc = specs.get("fet_channel", "")
            extracted_val_str = fc
            if fc:
                for e in entries:
                    if fc.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 14. TRANSISTOR TYPE
        elif "typetransistor" in attr_norm:
            tt = specs.get("tra_type", "")
            extracted_val_str = tt
            if tt:
                for e in entries:
                    if tt.casefold() == e.value_description.strip().casefold() or tt.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 15. REGULATOR ATTRIBUTES
        elif "typeregulator" in attr_norm:
            tr = specs.get("type_regulator", "")
            extracted_val_str = tr
            if tr:
                for e in entries:
                    if tr.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "regulatoroutputvoltage" in attr_norm:
            rv = specs.get("regulator_voltage", "")
            extracted_val_str = rv
            if rv:
                for e in entries:
                    if rv.casefold() in e.value_description.casefold() or e.value_description.strip().startswith(rv):
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 16. INDUCTOR SIZE
        elif "sizeinductor" in attr_norm:
            si = specs.get("size_inductor", "")
            extracted_val_str = si
            if si:
                for e in entries:
                    if si.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 17. LED COLOR
        elif "ledcolor" in attr_norm:
            lc = specs.get("led_color", "")
            extracted_val_str = lc
            if lc:
                for e in entries:
                    if lc.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 18. FUSE ATTRIBUTES
        elif "typefuse" in attr_norm:
            tf = specs.get("type_fuse", "")
            extracted_val_str = tf
            if tf:
                for e in entries:
                    if tf.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "fusecurrent" in attr_norm:
            fc = specs.get("fuse_current", "")
            extracted_val_str = fc
            if fc:
                for e in entries:
                    if fc.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 19. VARISTOR ATTRIBUTES
        elif "voltageac" in attr_norm:
            vac = specs.get("voltage_ac", "")
            extracted_val_str = vac
            if vac:
                for e in entries:
                    if vac.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "varistorsize" in attr_norm:
            vs = specs.get("varistor_size", "")
            extracted_val_str = vs
            if vs:
                for e in entries:
                    if vs.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 20. THERMISTOR TYPE
        elif "typethermistor" in attr_norm:
            tt = specs.get("type_thermistor", "")
            extracted_val_str = tt
            if tt:
                for e in entries:
                    if tt.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 21. POT TURNS
        elif "potturn" in attr_norm:
            pt = specs.get("pot_turns", "")
            extracted_val_str = pt
            if pt:
                for e in entries:
                    if pt.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 22. RELAY ATTRIBUTES
        elif "relaycoilvoltage" in attr_norm:
            rcv = specs.get("relay_coil_voltage", "")
            extracted_val_str = rcv
            if rcv:
                for e in entries:
                    if rcv.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "relaycontactvoltage" in attr_norm:
            rcvolt = specs.get("relay_contact_voltage", "")
            extracted_val_str = rcvolt
            if rcvolt:
                for e in entries:
                    if rcvolt.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "relaycontactcurrent" in attr_norm:
            rci = specs.get("relay_contact_current", "")
            extracted_val_str = rci
            if rci:
                for e in entries:
                    if rci.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "relaycontacttype" in attr_norm:
            rct = specs.get("relay_contact_type", "")
            extracted_val_str = rct
            if rct:
                for e in entries:
                    if rct.casefold() == e.value_description.strip().casefold() or rct.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break

        # 23. SWITCH / FERRITE / BUZZER / TRANSFORMER
        elif "typeswitch" in attr_norm:
            ts = specs.get("type_switch", "")
            extracted_val_str = ts
            if ts:
                for e in entries:
                    if ts.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "ferritetypes" in attr_norm:
            ft = specs.get("ferrite_type", "")
            extracted_val_str = ft
            if ft:
                for e in entries:
                    if ft.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "buzzerratedvoltage" in attr_norm:
            bv = specs.get("buzzer_voltage", "")
            extracted_val_str = bv
            if bv:
                for e in entries:
                    if bv.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "buzzerfrequency" in attr_norm:
            bf = specs.get("buzzer_frequency", "")
            extracted_val_str = bf
            if bf:
                for e in entries:
                    if bf.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "soundlevelbuzzer" in attr_norm:
            sl = specs.get("sound_level", "")
            extracted_val_str = sl
            if sl:
                for e in entries:
                    if sl.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "buzzersize" in attr_norm:
            bsz = specs.get("buzzer_size", "")
            extracted_val_str = bsz
            if bsz:
                for e in entries:
                    if bsz.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "typebatteryholder" in attr_norm:
            tbh = specs.get("type_battery_holder", "")
            extracted_val_str = tbh
            if tbh:
                for e in entries:
                    if tbh.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "typetransformer" in attr_norm:
            ttrf = specs.get("type_transformer", "")
            extracted_val_str = ttrf
            if ttrf:
                for e in entries:
                    if ttrf.casefold() in e.value_description.casefold():
                        matched_entry = e
                        status = "TAM EŞLEŞTİ"
                        break
        elif "typeantenna" in attr_norm:
            ta = specs.get("type_antenna", "")
            extracted_val_str = ta
            if ta:
                for e in entries:
                    if ta.casefold() in e.value_description.casefold():
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


def build_rich_description(material_type: str, specs: dict[str, Any], text: str = "") -> str:
    """Build a rich, industry-standard engineering description for any electronic component."""
    mfr = specs.get("manufacturer", "")
    text_lower = text.casefold()

    if material_type == "ANT":
        # e.g.: RF Board Antenna, WiFi, 50mm MHF Cable, TE Connectivity
        type_ant = specs.get("type_antenna", "RF")
        struct = "Board Antenna" if specs.get("mounting") == "TH" or "pcb" in text_lower or "board" in text_lower else "Antenna"
        proto = specs.get("protocol") or ("WiFi" if "wifi" in text_lower else "")
        cable_parts = []
        if specs.get("cable_length"):
            cable_parts.append(specs["cable_length"])
        if specs.get("mating_connector"):
            cable_parts.append(specs["mating_connector"])
        if cable_parts:
            cable_parts.append("Cable")
        cable_str = " ".join(cable_parts)

        tokens = [f"{type_ant} {struct}".strip()]
        if proto:
            tokens.append(proto)
        if cable_str:
            tokens.append(cable_str)
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "RES":
        # e.g.: Thick Film Chip Resistor, 255K 1% 1/10W, 0603 SMD, Automotive AEC-Q200, Yageo
        comp = specs.get("composition", "Thick Film")
        title = f"{comp} Chip Resistor" if specs.get("mounting") == "SMD" else f"{comp} Resistor"
        val_parts = []
        if specs.get("resistance_str"):
            val_parts.append(specs["resistance_str"])
        if specs.get("tolerance"):
            val_parts.append(specs["tolerance"])
        if specs.get("wattage"):
            val_parts.append(specs["wattage"])
        val_str = " ".join(val_parts)
        pkg_str = f"{specs.get('package', '')} {specs.get('mounting', 'SMD')}".strip()
        tokens = [title]
        if val_str:
            tokens.append(val_str)
        if pkg_str:
            tokens.append(pkg_str)
        if specs.get("special_def") and specs["special_def"] != "General":
            tokens.append(specs["special_def"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "CAP":
        # e.g.: Multilayer Ceramic Chip Capacitor, 100nF 50V 10%, 0805 SMD X7R, Samsung
        is_mlcc = specs.get("special_def") in ("X7R", "X5R", "C0G", "NP0", "Y5V", "X7S") or specs.get("package") in ("0402", "0603", "0805", "1206")
        title = "Multilayer Ceramic Chip Capacitor" if is_mlcc else "Capacitor"
        val_parts = []
        if specs.get("capacitance_str"):
            val_parts.append(specs["capacitance_str"])
        if specs.get("voltage"):
            val_parts.append(specs["voltage"])
        if specs.get("tolerance"):
            val_parts.append(specs["tolerance"])
        val_str = " ".join(val_parts)
        pkg_parts = []
        if specs.get("package"):
            pkg_parts.append(specs["package"])
        if specs.get("mounting"):
            pkg_parts.append(specs["mounting"])
        if specs.get("special_def"):
            pkg_parts.append(specs["special_def"])
        pkg_str = " ".join(pkg_parts)
        tokens = [title]
        if val_str:
            tokens.append(val_str)
        if pkg_str:
            tokens.append(pkg_str)
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "DIO":
        # e.g.: Schottky Barrier Diode, 40V 200mA, SOT-23 SMD, Diodes Inc
        t_dio = specs.get("type_diode", "Diode")
        title = f"{t_dio} Diode" if "diode" not in t_dio.lower() else t_dio
        ratings = []
        if specs.get("voltage"):
            ratings.append(specs["voltage"])
        if specs.get("current"):
            ratings.append(specs["current"])
        ratings_str = " ".join(ratings)
        pkg_str = f"{specs.get('package', '')} {specs.get('mounting', '')}".strip()
        tokens = [title]
        if specs.get("value_diode"):
            tokens.append(specs["value_diode"])
        if ratings_str:
            tokens.append(ratings_str)
        if pkg_str:
            tokens.append(pkg_str)
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "CON":
        # e.g.: Wire to Board Konnektör, 2.00mm, 4 Pin, 90°(Yatık), Molex
        title = specs.get("type_connector", "Konnektör")
        tokens = [title]
        if specs.get("pitch"):
            tokens.append(specs["pitch"])
        if specs.get("pin_number"):
            tokens.append(specs["pin_number"])
        if specs.get("angle"):
            tokens.append(specs["angle"])
        if specs.get("pin_type"):
            tokens.append(specs["pin_type"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "IND":
        # e.g.: Power Inductor, 10uH 1.5A, SMD, Coilcraft
        tokens = ["Power Inductor"]
        if specs.get("inductance_str"):
            tokens.append(specs["inductance_str"])
        if specs.get("size_inductor"):
            tokens.append(specs["size_inductor"])
        if specs.get("package"):
            tokens.append(specs["package"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "CRY":
        mount_pkg_parts = []
        if specs.get("mounting"):
            mount_pkg_parts.append(specs["mounting"])
        if specs.get("package"):
            if specs["package"] != specs.get("mounting"):
                mount_pkg_parts.append(specs["package"])
        mount_pkg_str = " ".join(mount_pkg_parts).strip()

        tokens = []
        if mount_pkg_str:
            tokens.append(mount_pkg_str)
        if specs.get("frequency_str"):
            tokens.append(specs["frequency_str"])
        if specs.get("load_capacitance"):
            tokens.append(specs["load_capacitance"])
        if specs.get("esr"):
            tokens.append(specs["esr"])
        if mfr:
            tokens.append(mfr)
        if not tokens:
            tokens.append("Quartz Crystal")
        return ", ".join(tokens)

    elif material_type == "BUZ":
        # e.g.: Magnetic Buzzer, SMD, 3V, 4kHz, 75dB, FUET
        struct = "Buzzer"
        if "piezo" in text_lower:
            struct = "Piezo Buzzer"
        elif "magnetic" in text_lower or "electromagnetic" in text_lower:
            struct = "Magnetic Buzzer"
        tokens = [struct]
        if specs.get("mounting"):
            tokens.append(specs["mounting"])
        if specs.get("buzzer_voltage"):
            tokens.append(specs["buzzer_voltage"])
        if specs.get("buzzer_frequency"):
            tokens.append(specs["buzzer_frequency"])
        if specs.get("sound_level"):
            tokens.append(specs["sound_level"])
        if specs.get("buzzer_size"):
            tokens.append(specs["buzzer_size"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "VAR":
        # e.g.: SIOV Leaded Metal Oxide Varistor, 275VAC, 7mm Disc, Epcos
        tokens = ["Metal Oxide Varistor"]
        if specs.get("voltage_ac"):
            tokens.append(specs["voltage_ac"])
        if specs.get("varistor_size"):
            tokens.append(specs["varistor_size"])
        if specs.get("mounting"):
            tokens.append(specs["mounting"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    elif material_type == "BAH":
        # e.g.: Coin Cell Battery Retainer / Clip, SMD, Ningbo KLS
        t_bah = specs.get("type_battery_holder", "Battery Clip / Retainer")
        tokens = [t_bah]
        if specs.get("mounting"):
            tokens.append(specs["mounting"])
        if mfr:
            tokens.append(mfr)
        return ", ".join(tokens)

    # General fallback
    mat_name = MATERIAL_TYPE_NAMES.get(material_type, material_type)
    tokens = [mat_name]
    for key in ("value_pn", "voltage", "current", "package", "mounting"):
        if specs.get(key):
            tokens.append(str(specs[key]))
    if mfr:
        tokens.append(mfr)
    return ", ".join(tokens)


def analyze_datasheet(
    stream_or_path: BinaryIO | bytes | Path | str,
    part_number: str | None = None,
    material_type_override: str | None = None,
    rulebook: RuleBook | None = None,
    filename: str = "datasheet.pdf",
    client: VisionLLMClient | None = None,
    serial_counter: str | None = None,
) -> dict[str, Any]:
    """Analyze a single datasheet PDF with multimodal hybrid fusion and map to DEKA ERP format."""
    # 1. Normalize stream_or_path
    if isinstance(stream_or_path, (bytes, bytearray)):
        file_bytes = bytes(stream_or_path)
    elif isinstance(stream_or_path, (str, Path)):
        p = Path(stream_or_path)
        if not filename or filename == "datasheet.pdf":
            filename = p.name
        try:
            file_bytes = p.read_bytes()
        except Exception:
            file_bytes = b""
    elif hasattr(stream_or_path, "read"):
        try:
            pos = stream_or_path.tell() if hasattr(stream_or_path, "tell") else None
            file_bytes = stream_or_path.read()
            if isinstance(file_bytes, str):
                file_bytes = file_bytes.encode("utf-8", errors="replace")
            if pos is not None and hasattr(stream_or_path, "seek"):
                stream_or_path.seek(pos)
        except Exception:
            file_bytes = b""
    else:
        file_bytes = b""

    if rulebook is None:
        rulebook = get_default_rulebook()

    # 2. Priority 1 (Fast-Path): PN Decoding
    raw_pn = (part_number or Path(filename).stem).strip()
    raw_pn = re.sub(r"\.pdf$", "", raw_pn, flags=re.I).strip()
    pn = raw_pn

    pn_specs: dict[str, Any] = {}
    if pn:
        try:
            pn_specs = decode_manufacturer_pn(pn) or {}
        except Exception:
            pn_specs = {}

    # Check if PN already resolves material type and core specs
    if material_type_override and material_type_override.upper() != "AUTO":
        mat_type = material_type_override.upper()
        conf = 1.0
        reason = f"Kullanıcı tarafından seçildi ({mat_type})."
    else:
        mat_type, conf, reason = classify_material_type("", pn, filename=filename)
        if mat_type == "BİLİNMİYOR" and pn_specs:
            if "resistance_str" in pn_specs or "wattage" in pn_specs:
                mat_type = "RES"
                conf = 0.90
                reason = "Parça numarası kod çözümünden direnç olarak tespit edildi."
            elif "capacitance_str" in pn_specs or "capacitance_pf" in pn_specs:
                mat_type = "CAP"
                conf = 0.90
                reason = "Parça numarası kod çözümünden kapasitör olarak tespit edildi."
            elif "pitch" in pn_specs or "type_connector" in pn_specs or "pin_number" in pn_specs:
                mat_type = "CON"
                conf = 0.90
                reason = "Parça numarası kod çözümünden konnektör olarak tespit edildi."

    pn_has_core = False
    if mat_type == "RES" and pn_specs.get("resistance_str"):
        pn_has_core = True
    elif mat_type == "CAP" and (pn_specs.get("capacitance_str") or pn_specs.get("capacitance_pf")):
        pn_has_core = True
    elif mat_type == "CON" and pn_specs.get("pitch") and pn_specs.get("pin_number"):
        pn_has_core = True
    elif mat_type in ("DIO", "FET", "TRA", "ICR", "ICS", "MCU", "BUZ", "BAH", "VAR") and bool(pn_specs):
        pn_has_core = True
    elif pn_specs and any(k in pn_specs for k in ("resistance_str", "capacitance_str", "capacitance_pf", "pin_number", "pitch", "voltage")):
        pn_has_core = True

    # 3. Priority 2: Text Extraction
    error_msg = None
    if pn_has_core:
        # Fast-Path: when PN decoding already resolved all core specs, bypass slow PDF text extraction
        text = ""
        clean_text = ""
        analyzed_pages = [1]
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            page_count = len(reader.pages)
        except Exception:
            page_count = 1
        is_catalog = False
        detected_mpns = [pn] if pn else []
    else:
        try:
            text, page_count, analyzed_pages = extract_pdf_data(file_bytes, part_number=pn)
        except Exception as exc:
            text, page_count, analyzed_pages = "", 0, []
            error_msg = f"PDF okunamadı: {str(exc)}"

        clean_text = text.strip()

        is_catalog = is_catalog_document_name(pn)
        detected_mpns = extract_part_numbers_from_text(text)

        # If user didn't enter a manual override, and text has a single unique genuine MPN
        if is_catalog and not part_number and len(detected_mpns) == 1:
            pn = detected_mpns[0]
            is_catalog = False
            if not pn_specs:
                try:
                    pn_specs = decode_manufacturer_pn(pn) or {}
                except Exception:
                    pn_specs = {}

        # 4. Material Type Classification (fallback to document text if still unknown)
        if mat_type == "BİLİNMİYOR":
            mat_type, conf, reason = classify_material_type(text, pn, filename=filename)
            if mat_type == "BİLİNMİYOR" and pn_specs:
                if "resistance_str" in pn_specs or "wattage" in pn_specs:
                    mat_type = "RES"
                    conf = 0.90
                    reason = "Parça numarası kod çözümünden direnç olarak tespit edildi."
                elif "capacitance_str" in pn_specs or "capacitance_pf" in pn_specs:
                    mat_type = "CAP"
                    conf = 0.90
                    reason = "Parça numarası kod çözümünden kapasitör olarak tespit edildi."
                elif "pitch" in pn_specs or "type_connector" in pn_specs or "pin_number" in pn_specs:
                    mat_type = "CON"
                    conf = 0.90
                    reason = "Parça numarası kod çözümünden konnektör olarak tespit edildi."

        # Re-check pn_has_core if PN was discovered from text
        if mat_type == "RES" and pn_specs.get("resistance_str"):
            pn_has_core = True
        elif mat_type == "CAP" and (pn_specs.get("capacitance_str") or pn_specs.get("capacitance_pf")):
            pn_has_core = True
        elif mat_type == "CON" and pn_specs.get("pitch") and pn_specs.get("pin_number"):
            pn_has_core = True
        elif mat_type in ("DIO", "FET", "TRA", "ICR", "ICS", "MCU", "BUZ", "BAH", "VAR") and bool(pn_specs):
            pn_has_core = True
        elif pn_specs and any(k in pn_specs for k in ("resistance_str", "capacitance_str", "capacitance_pf", "pin_number", "pitch", "voltage")):
            pn_has_core = True

    # Return early if unreadable empty document and unknown
    if mat_type == "BİLİNMİYOR" and error_msg:
        return {
            "filename": filename,
            "part_number": pn,
            "material_type": "BİLİNMİYOR",
            "material_type_name": "Bilinmiyor",
            "inferred_mounting": "",
            "confidence": 0.0,
            "suggested_description": "",
            "rich_description": "",
            "suggested_erp_code": "",
            "erp_code": "",
            "page_count": 0,
            "analyzed_pages": [],
            "summary_text": "",
            "error": error_msg,
            "is_catalog": False,
            "detected_part_numbers": [],
            "analysis_method": "deterministic_text",
            "properties": [],
            "_mappings": [],
            "raw_specs": {},
            "specs": {},
        }

    # Deterministic spec extraction
    # Fast-Path: when pn_has_core is True, do not pass file_bytes to avoid slow 2D CAD vector table extraction
    effective_file_bytes = None if pn_has_core else file_bytes
    specs = extract_component_specs(mat_type, text, pn, is_catalog=is_catalog, file_bytes=effective_file_bytes)
    for k, v in pn_specs.items():
        if k not in specs or specs[k] is None or specs[k] == "":
            specs[k] = v
    if serial_counter is not None:
        specs["serial_counter"] = serial_counter

    # 5. Priority 3: Vision Fusion Trigger
    is_text_empty = len(clean_text) == 0
    are_critical_missing = False

    if mat_type == "CON":
        are_critical_missing = not (specs.get("pitch") and specs.get("pin_number") and specs.get("type_connector"))
    elif mat_type == "RES":
        are_critical_missing = not specs.get("resistance_str")
    elif mat_type == "CAP":
        are_critical_missing = not (specs.get("capacitance_str") or specs.get("capacitance_pf"))
    elif mat_type == "IND":
        are_critical_missing = not specs.get("inductance_str")
    elif mat_type == "BİLİNMİYOR":
        are_critical_missing = True

    used_vision = False
    if file_bytes and not pn_has_core and (is_text_empty or are_critical_missing):
        vlm = client if client is not None else VisionLLMClient()
        try:
            if vlm.is_available():
                rendered = select_and_render_pages(file_bytes, target_pn=pn, max_pages=2)
                images = [p["base64_png"] for p in rendered if "base64_png" in p]
                if images:
                    vlm_hint = mat_type if mat_type != "BİLİNMİYOR" else None
                    vlm_res = vlm.extract_specs_from_images(images, target_pn=pn, material_type_hint=vlm_hint)
                    if vlm_res:
                        vlm_specs = vlm_res.get("specs") or {}
                        # Merge specs: deterministic/PN specs take precedence, VLM fills any empty/None field
                        for k, v in vlm_specs.items():
                            if v is not None and (k not in specs or specs[k] is None or specs[k] == ""):
                                specs[k] = v

                        vlm_mat = vlm_res.get("material_type")
                        if vlm_mat:
                            vlm_mat_clean = vlm_mat.strip().upper()
                            vlm_mat_normalized = NORMALIZE_VLM_TYPE.get(vlm_mat_clean, vlm_mat_clean)
                            if mat_type == "BİLİNMİYOR" or not mat_type:
                                mat_type = vlm_mat_normalized
                                conf = float(vlm_res.get("confidence", 0.85))
                                reason = f"Vision AI tarafından tespit edildi ({MATERIAL_TYPE_NAMES.get(mat_type, mat_type)})."
                        elif vlm_res.get("confidence"):
                            conf = max(conf, float(vlm_res.get("confidence", 0.0)))

                        if not specs.get("manufacturer") and vlm_res.get("manufacturer"):
                            specs["manufacturer"] = vlm_res["manufacturer"]

                        used_vision = True
        except Exception:
            pass

    # 6. Analysis Method Determination
    if used_vision:
        analysis_method = "hybrid_vision"
    elif pn_has_core:
        analysis_method = "fast_path"
    else:
        analysis_method = "deterministic_text"

    # 7. ERP Property Mapping & Presentation
    mat_type_name = MATERIAL_TYPE_NAMES.get(mat_type, mat_type)
    specs["is_catalog"] = is_catalog
    mounting = specs.get("mounting", "SMD")

    if rulebook is not None and mat_type != "BİLİNMİYOR":
        mappings, suggested_desc, suggested_code = map_specs_to_erp(rulebook, mat_type, specs)
    else:
        mappings, suggested_desc, suggested_code = [], "", ""

    rich_desc = build_rich_description(mat_type, specs, text=text)

    if is_catalog:
        conf = min(conf, 0.65)
        mfr = specs.get("manufacturer", "Üretici")
        summary = (
            f"⚠️ Katalog/Seri Dokümanı Tespit Edildi ({raw_pn}). "
            f"Bu doküman {mfr} genel ürün serisi kataloğudur (birden fazla kılıf ve değer içerir). "
            f"Kesin kılıf, direnç, tolerans ve güç eşleşmesi için lütfen tam parça kodunu (örn: RC0805JR-07300RL) giriniz."
        )
    elif mat_type == "BİLİNMİYOR":
        summary = "Malzeme tipi kesin belirlenemedi."
    else:
        summary = f"{mat_type_name} olarak tespit edildi ({int(conf * 100)}% güven). {reason}"

    serialized_specs = {
        k: str(v) if isinstance(v, Decimal) else v
        for k, v in specs.items()
    }

    return {
        "filename": filename,
        "part_number": pn,
        "material_type": mat_type,
        "material_type_name": mat_type_name,
        "inferred_mounting": mounting,
        "confidence": round(conf, 2),
        "suggested_description": suggested_desc,
        "rich_description": rich_desc or suggested_desc,
        "suggested_erp_code": suggested_code,
        "erp_code": suggested_code,
        "page_count": page_count,
        "analyzed_pages": analyzed_pages,
        "summary_text": summary,
        "error": error_msg,
        "is_catalog": is_catalog,
        "detected_part_numbers": detected_mpns,
        "analysis_method": analysis_method,
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
            for p in mappings
        ],
        "_mappings": mappings,
        "raw_specs": serialized_specs,
        "specs": serialized_specs,
    }


def analyze_single_datasheet(
    filename: str,
    file_bytes: bytes,
    rulebook: RuleBook,
    part_number_override: str | None = None,
    material_type_override: str | None = None,
) -> DatasheetAnalysisResult:
    """Analyze a single datasheet PDF and map its properties to DEKA ERP format."""
    res = analyze_datasheet(
        stream_or_path=file_bytes,
        part_number=part_number_override,
        material_type_override=material_type_override,
        rulebook=rulebook,
        filename=filename,
    )
    return DatasheetAnalysisResult(
        filename=res["filename"],
        part_number=res["part_number"],
        material_type=res["material_type"],
        material_type_name=res["material_type_name"],
        inferred_mounting=res["inferred_mounting"],
        confidence=res["confidence"],
        suggested_description=res["suggested_description"],
        rich_description=res["rich_description"],
        suggested_erp_code=res["suggested_erp_code"],
        properties=res["_mappings"],
        raw_specs=res["raw_specs"],
        page_count=res["page_count"],
        analyzed_pages=res["analyzed_pages"],
        summary_text=res["summary_text"],
        error=res["error"],
        is_catalog=res["is_catalog"],
        detected_part_numbers=res["detected_part_numbers"],
        analysis_method=res["analysis_method"],
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
        "Önerilen ERP Kodu",
        "Zengin Mühendislik Açıklaması",
        "Sözlük ERP Açıklaması",
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
            item.suggested_erp_code,
            item.rich_description or item.suggested_description,
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
