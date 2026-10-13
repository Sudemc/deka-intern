"""Semantic parser for electrical component values and descriptions.

Normalizes capacitance, resistance, inductance, voltage, tolerance,
and mounting types so equivalent values (e.g. 100nF == 0.1uF) do not
trigger false contradiction alerts.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
import math
import re
import unicodedata


# Capacitance unit multipliers to Farad
_CAP_MULTIPLIERS = {
    "pf": Decimal("1e-12"),
    "nf": Decimal("1e-9"),
    "uf": Decimal("1e-6"),
    "µf": Decimal("1e-6"),
    "mf": Decimal("1e-3"),
    "f": Decimal("1"),
}

# Resistance unit multipliers to Ohm
_RES_MULTIPLIERS = {
    "mr": Decimal("1e-3"),
    "mohm": Decimal("1e-3"),
    "mω": Decimal("1e-3"),
    "r": Decimal("1"),
    "ohm": Decimal("1"),
    "ω": Decimal("1"),
    "k": Decimal("1e3"),
    "kohm": Decimal("1e3"),
    "kω": Decimal("1e3"),
    "m": Decimal("1e6"),
    "meg": Decimal("1e6"),
    "mohm_mega": Decimal("1e6"),
}

# Inductance unit multipliers to Henry
_IND_MULTIPLIERS = {
    "nh": Decimal("1e-9"),
    "uh": Decimal("1e-6"),
    "µh": Decimal("1e-6"),
    "mh": Decimal("1e-3"),
    "h": Decimal("1"),
}


@dataclass(frozen=True)
class ElectricalSpecs:
    capacitance_farad: Decimal | None = None
    resistance_ohm: Decimal | None = None
    inductance_henry: Decimal | None = None
    voltage_volt: Decimal | None = None
    tolerance_pct: Decimal | None = None
    mounting: str | None = None  # 'SMD' or 'TH'
    package: str | None = None   # e.g. '0603', '0805', '1206'


def normalize_code_key(value: str) -> str:
    """Normalize a part number by removing non-alphanumeric characters.
    
    E.g. 'RC0603FR-0710KL' -> 'rc0603fr0710kl'
    """
    normalized = unicodedata.normalize("NFKD", str(value).casefold())
    return "".join(char for char in normalized if char.isalnum())


def normalize_mounting(value: str | None) -> str | None:
    """Normalize mounting type to 'SMD', 'TH', or None."""
    if not value:
        return None
    val = " ".join(str(value).split()).upper()
    val_clean = re.sub(r"[^A-Z]", "", val)
    if val_clean in {"S", "SMD", "SMT", "SURFACEMOUNT"}:
        return "SMD"
    if val_clean in {"T", "TH", "THT", "THROUGHHOLE", "DIP"}:
        return "TH"
    if "SMD" in val or "SMT" in val:
        return "SMD"
    if "THROUGH" in val or "THT" in val or "TH " in val or " TH" in val:
        return "TH"
    return None


def parse_specs(text: str) -> ElectricalSpecs:
    """Extract electrical specifications from a descriptive string."""
    clean = " " + text.replace(",", " ").replace(";", " ") + " "
    clean_lower = clean.casefold()

    cap = _extract_capacitance(clean_lower)
    res = _extract_resistance(clean_lower)
    ind = _extract_inductance(clean_lower)
    volt = _extract_voltage(clean_lower)
    tol = _extract_tolerance(clean_lower)
    mount = _extract_mounting(clean)
    pkg = _extract_package(clean)

    return ElectricalSpecs(
        capacitance_farad=cap,
        resistance_ohm=res,
        inductance_henry=ind,
        voltage_volt=volt,
        tolerance_pct=tol,
        mounting=mount,
        package=pkg,
    )


def _extract_capacitance(text: str) -> Decimal | None:
    # Matches: 100nf, 0.1uf, 10uf, 100000pf, 4.7uf, 4u7, etc.
    # Pattern 1: standard number + unit
    match = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(pf|nf|uf|µf|mf|f)(?![a-z0-9])", text)
    if match:
        num, unit = match.groups()
        return Decimal(num) * _CAP_MULTIPLIERS.get(unit, Decimal("1"))
    # Pattern 2: R-code like 4u7 -> 4.7 uF
    match_rcode = re.search(r"(?<![a-z0-9])(\d+)(u|p|n)(\d+)(?![a-z0-9])", text)
    if match_rcode:
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit = f"{unit_char}f"
        if unit in _CAP_MULTIPLIERS:
            return Decimal(num_str) * _CAP_MULTIPLIERS[unit]
    return None


def _extract_resistance(text: str) -> Decimal | None:
    # Matches: 10k, 10kohm, 10 kohm, 10000 ohm, 4k7, 0r1, 100mr, 1m, 330r, 330 ohm
    # Exclude capacitance or voltage matches like 10uf, 50v
    # Pattern 1: R-code like 4k7, 0r1, 2m2
    match_rcode = re.search(r"(?<![a-z0-9])(\d+)([rkmo])(\d+)(?![a-z0-9])", text)
    if match_rcode:
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit_map = {"r": "r", "k": "k", "m": "m", "o": "r"}
        unit = unit_map[unit_char]
        return Decimal(num_str) * _RES_MULTIPLIERS[unit]

    # Pattern 2: number + unit (k, kohm, ohm, r, etc.)
    match = re.search(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(kohm|kω|mohm|mω|ohm|ω|mr|k|r)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        # Avoid matching 'm' as mega unless explicit 'mohm' or 'meg'
        if unit == "k":
            return Decimal(num) * Decimal("1e3")
        if unit in _RES_MULTIPLIERS:
            return Decimal(num) * _RES_MULTIPLIERS[unit]

    # Pattern 3: explicit "xxx ohm" or "xxx R"
    match_ohm = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:ohm|r)(?![a-z0-9])", text)
    if match_ohm:
        return Decimal(match_ohm.group(1))

    return None


def _extract_inductance(text: str) -> Decimal | None:
    match = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(nh|uh|µh|mh|h)(?![a-z0-9])", text)
    if match:
        num, unit = match.groups()
        return Decimal(num) * _IND_MULTIPLIERS.get(unit, Decimal("1"))
    return None


def _extract_voltage(text: str) -> Decimal | None:
    # Matches: 50v, 50 v, 50volt, 6.3v, 6v3, 3v3
    match_vcode = re.search(r"(?<![a-z0-9])(\d+)v(\d+)(?![a-z0-9])", text)
    if match_vcode:
        whole, frac = match_vcode.groups()
        return Decimal(f"{whole}.{frac}")

    match = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:v|volt|vdc)(?!\.?reg)(?![a-z0-9])", text)
    if match:
        return Decimal(match.group(1))
    return None


def _extract_tolerance(text: str) -> Decimal | None:
    # Matches: 1%, %1, +/-1%, ±1%, 5%, 10%
    # Pattern 1: number followed by % (e.g. 1%, 5%, +/-1%)
    match_suff = re.search(r"(?<![a-z0-9])(?:[±\+\/-]*\s*)(\d+(?:\.\d+)?)\s*%", text)
    if match_suff:
        val = Decimal(match_suff.group(1))
        if val <= 50:
            return val

    # Pattern 2: % followed by number (e.g. %1, %5)
    match_pref = re.search(r"%\s*(\d+(?:\.\d+)?)(?![a-z0-9])", text)
    if match_pref:
        val = Decimal(match_pref.group(1))
        if val <= 50:
            return val

    return None


def _extract_mounting(text: str) -> str | None:
    return normalize_mounting(text)


def _extract_package(text: str) -> str | None:
    # Common packages: 0201, 0402, 0603, 0805, 1206, 1210, 2010, 2512, SOT-23, SOD-123, etc.
    match = re.search(r"\b(0201|0402|0603|0805|1206|1210|2010|2512|SOT-?23|SOD-?123|TO-?220)\b", text, re.I)
    if match:
        return match.group(1).upper().replace("-", "")
    return None


def compare_descriptions(customer_text: str, erp_text: str) -> list[str]:
    """Compare customer and ERP descriptions semantically.
    
    Returns a list of issue descriptions if there are genuine contradictions.
    Returns [] if they are compatible or equivalent.
    """
    cust_clean = " ".join(customer_text.split()).casefold()
    erp_clean = " ".join(erp_text.split()).casefold()

    if not cust_clean and not erp_clean:
        return ["Müşteri Value ve ERP Description boş"]
    if not cust_clean and erp_clean:
        return ["Müşteri verisi boş, ERP verisi kullanılabilir"]
    if cust_clean and not erp_clean:
        return ["ERP Description boş"]

    # If normalized strings are identical, no issue
    if cust_clean == erp_clean:
        return []

    cust_specs = parse_specs(customer_text)
    erp_specs = parse_specs(erp_text)

    contradictions: list[str] = []

    # Check Capacitance
    if cust_specs.capacitance_farad is not None and erp_specs.capacitance_farad is not None:
        if cust_specs.capacitance_farad != erp_specs.capacitance_farad:
            contradictions.append(
                f"Açıklama Çelişkisi (Kapasite farklı: Müşteri {_format_cap(cust_specs.capacitance_farad)}, ERP {_format_cap(erp_specs.capacitance_farad)})"
            )

    # Check Resistance
    if cust_specs.resistance_ohm is not None and erp_specs.resistance_ohm is not None:
        if cust_specs.resistance_ohm != erp_specs.resistance_ohm:
            contradictions.append(
                f"Açıklama Çelişkisi (Direnç farklı: Müşteri {_format_res(cust_specs.resistance_ohm)}, ERP {_format_res(erp_specs.resistance_ohm)})"
            )

    # Check Voltage
    if cust_specs.voltage_volt is not None and erp_specs.voltage_volt is not None:
        if cust_specs.voltage_volt != erp_specs.voltage_volt:
            contradictions.append(
                f"Açıklama Çelişkisi (Voltaj farklı: Müşteri {cust_specs.voltage_volt}V, ERP {erp_specs.voltage_volt}V)"
            )

    # Check Tolerance
    if cust_specs.tolerance_pct is not None and erp_specs.tolerance_pct is not None:
        if cust_specs.tolerance_pct != erp_specs.tolerance_pct:
            contradictions.append(
                f"Açıklama Çelişkisi (Tolerans farklı: Müşteri %{cust_specs.tolerance_pct}, ERP %{erp_specs.tolerance_pct})"
            )

    # Check Package
    if cust_specs.package is not None and erp_specs.package is not None:
        if cust_specs.package != erp_specs.package:
            contradictions.append(
                f"Açıklama Çelişkisi (Kılıf/Paket farklı: Müşteri {cust_specs.package}, ERP {erp_specs.package})"
            )

    if contradictions:
        return contradictions

    # If any specific electrical spec was verified to match, we consider it compatible!
    has_matching_spec = (
        (cust_specs.capacitance_farad is not None and cust_specs.capacitance_farad == erp_specs.capacitance_farad)
        or (cust_specs.resistance_ohm is not None and cust_specs.resistance_ohm == erp_specs.resistance_ohm)
        or (cust_specs.voltage_volt is not None and cust_specs.voltage_volt == erp_specs.voltage_volt)
    )
    if has_matching_spec:
        return []

    # If customer text is a subset of tokens in ERP text
    cust_tokens = set(re.findall(r"[a-z0-9]+", cust_clean))
    erp_tokens = set(re.findall(r"[a-z0-9]+", erp_clean))
    if cust_tokens and cust_tokens.issubset(erp_tokens):
        return []

    return []


def _format_cap(farad: Decimal) -> str:
    if farad >= Decimal("1e-6"):
        return f"{farad / Decimal('1e-6'):g}uF"
    if farad >= Decimal("1e-9"):
        return f"{farad / Decimal('1e-9'):g}nF"
    return f"{farad / Decimal('1e-12'):g}pF"


def _format_res(ohm: Decimal) -> str:
    if ohm >= Decimal("1e6"):
        return f"{ohm / Decimal('1e6'):g}M"
    if ohm >= Decimal("1e3"):
        return f"{ohm / Decimal('1e3'):g}k"
    return f"{ohm:g}R"
