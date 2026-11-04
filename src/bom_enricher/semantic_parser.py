"""Semantic parser for electrical component values and descriptions.

Normalizes capacitance, resistance, inductance, voltage, tolerance,
mounting types, frequencies, colors, and polarities so equivalent values
(e.g. 100nF == 0.1uF) do not trigger false contradiction alerts while genuine
discrepancies are reliably flagged.
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
    "\u03bcf": Decimal("1e-6"),
    "mf": Decimal("1e-3"),
    "f": Decimal("1"),
}

# Resistance unit multipliers to Ohm
_RES_MULTIPLIERS = {
    "mr": Decimal("1e-3"),
    "mohm": Decimal("1e-3"),
    "mohms": Decimal("1e-3"),
    "mω": Decimal("1e-3"),
    "r": Decimal("1"),
    "ohm": Decimal("1"),
    "ohms": Decimal("1"),
    "ω": Decimal("1"),
    "k": Decimal("1e3"),
    "kohm": Decimal("1e3"),
    "kohms": Decimal("1e3"),
    "kω": Decimal("1e3"),
    "m": Decimal("1e6"),
    "meg": Decimal("1e6"),
    "megohm": Decimal("1e6"),
    "megohms": Decimal("1e6"),
    "mohm_mega": Decimal("1e6"),
}

# Inductance unit multipliers to Henry
_IND_MULTIPLIERS = {
    "nh": Decimal("1e-9"),
    "uh": Decimal("1e-6"),
    "µh": Decimal("1e-6"),
    "\u03bch": Decimal("1e-6"),
    "mh": Decimal("1e-3"),
    "h": Decimal("1"),
}

# Frequency unit multipliers to Hertz
_FREQ_MULTIPLIERS = {
    "hz": Decimal("1"),
    "khz": Decimal("1e3"),
    "mhz": Decimal("1e6"),
    "ghz": Decimal("1e9"),
}

# Current unit multipliers to Ampere
_CURRENT_MULTIPLIERS = {
    "a": Decimal("1"),
    "ma": Decimal("1e-3"),
    "ua": Decimal("1e-6"),
    "µa": Decimal("1e-6"),
    "\u03bca": Decimal("1e-6"),
}

# Power unit multipliers to Watt
_POWER_MULTIPLIERS = {
    "w": Decimal("1"),
    "mw": Decimal("1e-3"),
    "kw": Decimal("1e3"),
}

_COLOR_MAP = {
    "red": "RED",
    "kırmızı": "RED",
    "kirmizi": "RED",
    "green": "GREEN",
    "yeşil": "GREEN",
    "yesil": "GREEN",
    "yellow": "YELLOW",
    "sarı": "YELLOW",
    "sari": "YELLOW",
    "blue": "BLUE",
    "mavi": "BLUE",
    "white": "WHITE",
    "beyaz": "WHITE",
    "amber": "AMBER",
    "orange": "ORANGE",
    "turuncu": "ORANGE",
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
    frequency_hz: Decimal | None = None
    current_amp: Decimal | None = None
    power_watt: Decimal | None = None
    dielectric: str | None = None
    category: str | None = None
    color: str | None = None
    polarity: str | None = None
    frequencies: frozenset[Decimal] = frozenset()
    currents: frozenset[Decimal] = frozenset()
    powers: frozenset[Decimal] = frozenset()
    dielectrics: frozenset[str] = frozenset()
    categories: frozenset[str] = frozenset()
    colors: frozenset[str] = frozenset()
    polarities: frozenset[str] = frozenset()


@dataclass(frozen=True)
class MultiElectricalSpecs:
    capacitances: frozenset[Decimal] = frozenset()
    resistances: frozenset[Decimal] = frozenset()
    inductances: frozenset[Decimal] = frozenset()
    voltages: frozenset[Decimal] = frozenset()
    tolerances: frozenset[Decimal] = frozenset()
    packages: frozenset[str] = frozenset()
    mountings: frozenset[str] = frozenset()
    frequencies: frozenset[Decimal] = frozenset()
    currents: frozenset[Decimal] = frozenset()
    powers: frozenset[Decimal] = frozenset()
    dielectrics: frozenset[str] = frozenset()
    categories: frozenset[str] = frozenset()
    colors: frozenset[str] = frozenset()
    polarities: frozenset[str] = frozenset()


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


def _normalize_mega(text: str) -> str:
    """Convert uppercase M / MOhm / Meg resistor notations to megohm before casefolding.
    
    Prevents 10M / 10MOhm from being conflated with milliohms (10mOhm) or losing case information.
    Guards against hyphenated part numbers (e.g. 7M-27.120, VLP8040T-150M).
    """
    return re.sub(
        r"(?<![a-zA-Z0-9\-_])(?:(\d+\s*[xX*]\s*))?(\d+(?:\.\d+)?)\s*(?:M[oO][hH][mM]s?|M[oO]hms?|MΩ|Meg|MEG|M)(?![a-zA-Z0-9\-_])",
        lambda m: (m.group(1) or "") + m.group(2) + "megohm",
        text,
    )


def _strip_jedec(text: str) -> str:
    """Strip standard JEDEC part codes (1Nxxxx, 2Nxxxx, 3Nxxxx) so they aren't parsed as Farad/Henry."""
    return re.sub(r"\b[123]n\d{3,4}[a-z]?\b", " ", text, flags=re.I)


def parse_specs(text: str) -> ElectricalSpecs:
    """Extract electrical specifications from a descriptive string.
    
    Backward-compatible: returns the first extracted specification for each property.
    """
    clean = " " + text.replace(",", " ").replace(";", " ") + " "
    clean_mega = _normalize_mega(clean)
    clean_lower = clean_mega.casefold()

    cap = _extract_capacitance(clean_lower)
    res = _extract_resistance(clean_lower)
    ind = _extract_inductance(clean_lower)
    volt = _extract_voltage(clean_lower)
    tol = _extract_tolerance(clean_lower)
    mount = _extract_mounting(clean)
    pkg = _extract_package(clean)
    freq = _extract_frequency(clean_lower)
    curr = _extract_current(clean_lower)
    pwr = _extract_power(clean_lower)
    diel = _extract_dielectric(clean_lower)
    cat = _extract_category(clean_lower)
    col = _extract_color(clean_lower)
    pol = _extract_polarity(clean_lower)

    freqs = _extract_frequencies(clean_lower)
    currs = _extract_currents(clean_lower)
    pwrs = _extract_powers(clean_lower)
    diels = _extract_dielectrics(clean_lower)
    cats = _extract_categories(clean_lower)
    cols = _extract_colors(clean_lower)
    pols = _extract_polarities(clean_lower)

    return ElectricalSpecs(
        capacitance_farad=cap,
        resistance_ohm=res,
        inductance_henry=ind,
        voltage_volt=volt,
        tolerance_pct=tol,
        mounting=mount,
        package=pkg,
        frequency_hz=freq,
        current_amp=curr,
        power_watt=pwr,
        dielectric=diel,
        category=cat,
        color=col,
        polarity=pol,
        frequencies=freqs,
        currents=currs,
        powers=pwrs,
        dielectrics=diels,
        categories=cats,
        colors=cols,
        polarities=pols,
    )


def parse_specs_set(text: str) -> MultiElectricalSpecs:
    """Extract all electrical specifications from a descriptive string as sets."""
    clean = " " + text.replace(",", " ").replace(";", " ") + " "
    clean_mega = _normalize_mega(clean)
    clean_lower = clean_mega.casefold()

    return MultiElectricalSpecs(
        capacitances=_extract_capacitances(clean_lower),
        resistances=_extract_resistances(clean_lower),
        inductances=_extract_inductances(clean_lower),
        voltages=_extract_voltages(clean_lower),
        tolerances=_extract_tolerances(clean_lower),
        packages=_extract_packages(clean),
        mountings=_extract_mountings(clean),
        frequencies=_extract_frequencies(clean_lower),
        currents=_extract_currents(clean_lower),
        powers=_extract_powers(clean_lower),
        dielectrics=_extract_dielectrics(clean_lower),
        categories=_extract_categories(clean_lower),
        colors=_extract_colors(clean_lower),
        polarities=_extract_polarities(clean_lower),
    )


def _extract_capacitances(text: str) -> frozenset[Decimal]:
    caps: set[Decimal] = set()
    text_clean = _strip_jedec(text)
    # Pattern 1: standard number + unit (supports optional array prefix)
    for match in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(pf|nf|uf|µf|\u03bcf|mf|f)(?![a-z0-9])",
        text_clean,
    ):
        num, unit = match.groups()
        caps.add(Decimal(num) * _CAP_MULTIPLIERS.get(unit, Decimal("1")))
    # Pattern 2: R-code like 4u7 -> 4.7 uF
    for match_rcode in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)(u|p|n|µ|\u03bc)(\d+)(?![a-z0-9])",
        text_clean,
    ):
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit = "uf" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}f"
        if unit in _CAP_MULTIPLIERS:
            caps.add(Decimal(num_str) * _CAP_MULTIPLIERS[unit])
    return frozenset(caps)


def _extract_capacitance(text: str) -> Decimal | None:
    text_clean = _strip_jedec(text)
    match = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(pf|nf|uf|µf|\u03bcf|mf|f)(?![a-z0-9])",
        text_clean,
    )
    if match:
        num, unit = match.groups()
        return Decimal(num) * _CAP_MULTIPLIERS.get(unit, Decimal("1"))
    match_rcode = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)(u|p|n|µ|\u03bc)(\d+)(?![a-z0-9])",
        text_clean,
    )
    if match_rcode:
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit = "uf" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}f"
        if unit in _CAP_MULTIPLIERS:
            return Decimal(num_str) * _CAP_MULTIPLIERS[unit]
    return None


def _extract_resistances(text: str) -> frozenset[Decimal]:
    res: set[Decimal] = set()
    # Pattern 1: R-code like 4k7, 0r1, 2m2
    for match_rcode in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)([rkmo])(\d+)(?![a-z0-9])",
        text,
    ):
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit_map = {"r": "r", "k": "k", "m": "m", "o": "r"}
        unit = unit_map[unit_char]
        res.add(Decimal(num_str) * _RES_MULTIPLIERS[unit])

    # Pattern 2: number + unit (megohm, meg, k, kohm, ohm, r, etc.)
    for match in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(megohms?|meg|kohms?|kω|mohms?|mω|ohms?|ω|mr|k|r)(?![a-z0-9])",
        text,
    ):
        num, unit = match.groups()
        if unit == "k":
            res.add(Decimal(num) * Decimal("1e3"))
        elif unit in _RES_MULTIPLIERS:
            res.add(Decimal(num) * _RES_MULTIPLIERS[unit])

    # Pattern 3: explicit "xxx ohm" or "xxx R"
    for match_ohm in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(?:ohm|r)(?![a-z0-9])",
        text,
    ):
        res.add(Decimal(match_ohm.group(1)))

    return frozenset(res)


def _extract_resistance(text: str) -> Decimal | None:
    match_rcode = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)([rkmo])(\d+)(?![a-z0-9])",
        text,
    )
    if match_rcode:
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit_map = {"r": "r", "k": "k", "m": "m", "o": "r"}
        unit = unit_map[unit_char]
        return Decimal(num_str) * _RES_MULTIPLIERS[unit]

    match = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(megohms?|meg|kohms?|kω|mohms?|mω|ohms?|ω|mr|k|r)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        if unit == "k":
            return Decimal(num) * Decimal("1e3")
        if unit in _RES_MULTIPLIERS:
            return Decimal(num) * _RES_MULTIPLIERS[unit]

    match_ohm = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(?:ohm|r)(?![a-z0-9])",
        text,
    )
    if match_ohm:
        return Decimal(match_ohm.group(1))

    return None


def _extract_inductances(text: str) -> frozenset[Decimal]:
    inds: set[Decimal] = set()
    text_clean = _strip_jedec(text)
    for match in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(nh|uh|µh|\u03bch|mh|h)(?![a-z0-9])",
        text_clean,
    ):
        num, unit = match.groups()
        inds.add(Decimal(num) * _IND_MULTIPLIERS.get(unit, Decimal("1")))
    # Require explicit 'h' after prefix so 4u7, 4n7, 2m2 are never parsed as inductance
    for match_code in re.finditer(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)(u|µ|\u03bc|m|n)h(\d+)(?![a-z0-9])",
        text_clean,
    ):
        whole, unit_char, frac = match_code.groups()
        num_str = f"{whole}.{frac}"
        unit = "uh" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}h"
        if unit in _IND_MULTIPLIERS:
            inds.add(Decimal(num_str) * _IND_MULTIPLIERS[unit])
    return frozenset(inds)


def _extract_inductance(text: str) -> Decimal | None:
    text_clean = _strip_jedec(text)
    match = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+(?:\.\d+)?)\s*(nh|uh|µh|\u03bch|mh|h)(?![a-z0-9])",
        text_clean,
    )
    if match:
        num, unit = match.groups()
        return Decimal(num) * _IND_MULTIPLIERS.get(unit, Decimal("1"))
    match_code = re.search(
        r"(?<![a-z0-9])(?:\d+\s*[x*]\s*)?(\d+)(u|µ|\u03bc|m|n)h(\d+)(?![a-z0-9])",
        text_clean,
    )
    if match_code:
        whole, unit_char, frac = match_code.groups()
        num_str = f"{whole}.{frac}"
        unit = "uh" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}h"
        if unit in _IND_MULTIPLIERS:
            return Decimal(num_str) * _IND_MULTIPLIERS[unit]
    return None


def _extract_voltages(text: str) -> frozenset[Decimal]:
    volts: set[Decimal] = set()
    for match_vcode in re.finditer(r"(?<![a-z0-9])(\d+)v(\d+)(?![a-z0-9])", text):
        whole, frac = match_vcode.groups()
        volts.add(Decimal(f"{whole}.{frac}"))

    for m in re.finditer(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:v|volt|vdc)(?!\.?reg)(?!\.?\d)(?![a-z0-9])", text):
        val = Decimal(m.group(1))
        if 1950 <= val <= 2050 or val > 1000:
            continue
        volts.add(val)
    return frozenset(volts)


def _extract_voltage(text: str) -> Decimal | None:
    match_vcode = re.search(r"(?<![a-z0-9])(\d+)v(\d+)(?![a-z0-9])", text)
    if match_vcode:
        whole, frac = match_vcode.groups()
        return Decimal(f"{whole}.{frac}")

    for m in re.finditer(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:v|volt|vdc)(?!\.?reg)(?!\.?\d)(?![a-z0-9])", text):
        val = Decimal(m.group(1))
        if 1950 <= val <= 2050 or val > 1000:
            continue
        return val
    return None


def _extract_tolerances(text: str) -> frozenset[Decimal]:
    tols: set[Decimal] = set()
    for match_suff in re.finditer(r"(?<![a-z0-9])(?:[±\+\/-]*\s*)(\d+(?:\.\d+)?)\s*%", text):
        val = Decimal(match_suff.group(1))
        if val <= 50:
            tols.add(val)

    for match_pref in re.finditer(r"%\s*(\d+(?:\.\d+)?)(?![a-z0-9])", text):
        val = Decimal(match_pref.group(1))
        if val <= 50:
            tols.add(val)

    return frozenset(tols)


def _extract_tolerance(text: str) -> Decimal | None:
    match_suff = re.search(r"(?<![a-z0-9])(?:[±\+\/-]*\s*)(\d+(?:\.\d+)?)\s*%", text)
    if match_suff:
        val = Decimal(match_suff.group(1))
        if val <= 50:
            return val

    match_pref = re.search(r"%\s*(\d+(?:\.\d+)?)(?![a-z0-9])", text)
    if match_pref:
        val = Decimal(match_pref.group(1))
        if val <= 50:
            return val

    return None


def _extract_mountings(text: str) -> frozenset[str]:
    mounts: set[str] = set()
    clean = " " + text + " "
    if re.search(r"(?<![a-z0-9])(smd|smt|surface\s*mount)(?![a-z0-9])", clean, re.I):
        mounts.add("SMD")
    if re.search(r"(?<![a-z0-9])(th|tht|through\s*hole|dip)(?![a-z0-9])", clean, re.I):
        mounts.add("TH")
    if not mounts:
        single = normalize_mounting(text)
        if single:
            mounts.add(single)
    return frozenset(mounts)


def _extract_mounting(text: str) -> str | None:
    return normalize_mounting(text)


def _extract_packages(text: str) -> frozenset[str]:
    pkgs: set[str] = set()
    for match in re.finditer(r"\b(0201|0402|0603|0805|1206|1210|2010|2512|SOT-?23|SOD-?123|TO-?220)\b", text, re.I):
        pkgs.add(match.group(1).upper().replace("-", ""))
    return frozenset(pkgs)


def _extract_package(text: str) -> str | None:
    match = re.search(r"\b(0201|0402|0603|0805|1206|1210|2010|2512|SOT-?23|SOD-?123|TO-?220)\b", text, re.I)
    if match:
        return match.group(1).upper().replace("-", "")
    return None


def _extract_frequencies(text: str) -> frozenset[Decimal]:
    freqs: set[Decimal] = set()
    for match in re.finditer(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(ghz|mhz|khz|hz)(?![a-z0-9])",
        text,
    ):
        num, unit = match.groups()
        freqs.add(Decimal(num) * _FREQ_MULTIPLIERS[unit])
    return frozenset(freqs)


def _extract_frequency(text: str) -> Decimal | None:
    match = re.search(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(ghz|mhz|khz|hz)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        return Decimal(num) * _FREQ_MULTIPLIERS[unit]
    return None


def _extract_currents(text: str) -> frozenset[Decimal]:
    currs: set[Decimal] = set()
    for match in re.finditer(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(ma|ua|µa|\u03bca|a)(?![a-z0-9])",
        text,
    ):
        num, unit = match.groups()
        currs.add(Decimal(num) * _CURRENT_MULTIPLIERS[unit])
    return frozenset(currs)


def _extract_current(text: str) -> Decimal | None:
    match = re.search(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(ma|ua|µa|\u03bca|a)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        return Decimal(num) * _CURRENT_MULTIPLIERS[unit]
    return None


def _extract_powers(text: str) -> frozenset[Decimal]:
    pwrs: set[Decimal] = set()
    # Fractional watts: 1/4W, 1/8W, 1/10W, etc.
    for match_frac in re.finditer(
        r"(?<![a-z0-9])(\d+)/(\d+)\s*w(?![a-z0-9])",
        text,
    ):
        num, den = match_frac.groups()
        if Decimal(den) != 0:
            pwrs.add(Decimal(num) / Decimal(den))
    # Standard: 1W, 250mW, etc.
    for match in re.finditer(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(mw|kw|w)(?![a-z0-9])",
        text,
    ):
        num, unit = match.groups()
        pwrs.add(Decimal(num) * _POWER_MULTIPLIERS[unit])
    return frozenset(pwrs)


def _extract_power(text: str) -> Decimal | None:
    match_frac = re.search(
        r"(?<![a-z0-9])(\d+)/(\d+)\s*w(?![a-z0-9])",
        text,
    )
    if match_frac:
        num, den = match_frac.groups()
        if Decimal(den) != 0:
            return Decimal(num) / Decimal(den)
    match = re.search(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(mw|kw|w)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        return Decimal(num) * _POWER_MULTIPLIERS[unit]
    return None


def _extract_dielectrics(text: str) -> frozenset[str]:
    diels: set[str] = set()
    for match in re.finditer(r"\b(c0g|np0|x7r|x5r|y5v|x8r|x6s|x7s|x7t|x5s)\b", text):
        diels.add(match.group(1).upper())
    return frozenset(diels)


def _extract_dielectric(text: str) -> str | None:
    match = re.search(r"\b(c0g|np0|x7r|x5r|y5v|x8r|x6s|x7s|x7t|x5s)\b", text)
    if match:
        return match.group(1).upper()
    return None


def _extract_categories(text: str) -> frozenset[str]:
    cats: set[str] = set()
    if re.search(r"\b(res|resistor|direnç|direnc|potentiometer|pot)\b", text):
        cats.add("RESISTOR")
    if re.search(r"\b(cap|capacitor|kondansatör|kondansator|tantalum|ceramic|seramik|electrolytic)\b", text):
        cats.add("CAPACITOR")
    if re.search(r"\b(inductor|indüktör|induktor|bobin|choke|ferrite|bead)\b", text):
        cats.add("INDUCTOR")
    if re.search(r"\b(diode|diyot|rectifier|zener|schottky|tvs|transil|varistor)\b", text) or re.search(r"\b1n\d{3,4}[a-z]?\b", text):
        cats.add("DIODE")
    if re.search(r"\bleds?\b", text):
        cats.add("LED")
    if re.search(r"\b(transistor|tranzistör|mosfet|fet|bjt|igbt)\b", text) or re.search(r"\b2n\d{3,4}[a-z]?\b", text):
        cats.add("TRANSISTOR")
    if re.search(r"\b(crystal|kristal|oscillator|osilator|resonator|xtal)\b", text):
        cats.add("CRYSTAL")
    if re.search(r"\b(fuse|sigorta)\b", text):
        cats.add("FUSE")
    if re.search(r"\b(connector|konektör|konektor|header|socket|plug|terminal)\b", text):
        cats.add("CONNECTOR")
    if re.search(r"\b(relay|röle|role)\b", text):
        cats.add("RELAY")
    if re.search(r"\b(switch|anahtar|butona?)\b", text):
        cats.add("SWITCH")
    if re.search(r"\b(ic|entegre|microcontroller|mcu|regulator)\b", text):
        cats.add("IC")
    return frozenset(cats)


def _extract_category(text: str) -> str | None:
    cats = _extract_categories(text)
    if cats:
        return next(iter(cats))
    return None


def _extract_colors(text: str) -> frozenset[str]:
    cols: set[str] = set()
    for match in re.finditer(
        r"\b(red|kırmızı|kirmizi|green|yeşil|yesil|yellow|sarı|sari|blue|mavi|white|beyaz|amber|orange|turuncu)\b",
        text,
    ):
        word = match.group(1)
        if word in _COLOR_MAP:
            cols.add(_COLOR_MAP[word])
    return frozenset(cols)


def _extract_color(text: str) -> str | None:
    match = re.search(
        r"\b(red|kırmızı|kirmizi|green|yeşil|yesil|yellow|sarı|sari|blue|mavi|white|beyaz|amber|orange|turuncu)\b",
        text,
    )
    if match and match.group(1) in _COLOR_MAP:
        return _COLOR_MAP[match.group(1)]
    return None


def _extract_polarities(text: str) -> frozenset[str]:
    pols: set[str] = set()
    if re.search(r"\bnpn\b", text):
        pols.add("NPN")
    if re.search(r"\bpnp\b", text):
        pols.add("PNP")
    if re.search(r"\b(?:n[- ]?ch(?:annel)?|n[- ]?kanal)\b", text):
        pols.add("N-CHANNEL")
    if re.search(r"\b(?:p[- ]?ch(?:annel)?|p[- ]?kanal)\b", text):
        pols.add("P-CHANNEL")
    return frozenset(pols)


def _extract_polarity(text: str) -> str | None:
    pols = _extract_polarities(text)
    if pols:
        return next(iter(pols))
    return None


def _calc_ratio(a: Decimal, b: Decimal) -> Decimal:
    """Calculate ratio between two electrical values with zero-guard."""
    min_val = min(a, b)
    if min_val <= Decimal("0"):
        return Decimal("Infinity")
    return max(a, b) / min_val


def compare_descriptions(customer_text: str, erp_text: str, mpn_matched: bool = False) -> list[str]:
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

    cust_specs = parse_specs_set(customer_text)
    erp_specs = parse_specs_set(erp_text)

    contradictions: list[str] = []

    # 0. Cross-category conflict: Capacitance vs Resistance/Inductance (only if not mpn_matched)
    if not mpn_matched:
        composite_keywords = r"\b(rc[- ]?net|rc[- ]?network|rc[- ]?filter|filter|filtre|f\.?bead|ferrite\s*bead|bead)\b"
        is_cust_composite = (
            bool(re.search(composite_keywords, cust_clean))
            or (bool(cust_specs.capacitances) and bool(cust_specs.resistances or cust_specs.inductances))
            or (bool(cust_specs.resistances) and bool(cust_specs.inductances))
        )
        is_erp_composite = (
            bool(re.search(composite_keywords, erp_clean))
            or (bool(erp_specs.capacitances) and bool(erp_specs.resistances or erp_specs.inductances))
            or (bool(erp_specs.resistances) and bool(erp_specs.inductances))
        )

        if not is_cust_composite and not is_erp_composite:
            cust_is_cap_only = bool(cust_specs.capacitances) and not cust_specs.resistances and not cust_specs.inductances
            erp_is_res_or_ind_only = bool(erp_specs.resistances or erp_specs.inductances) and not erp_specs.capacitances

            cust_is_res_or_ind_only = bool(cust_specs.resistances or cust_specs.inductances) and not cust_specs.capacitances
            erp_is_cap_only = bool(erp_specs.capacitances) and not erp_specs.resistances and not erp_specs.inductances

            if (cust_is_cap_only and erp_is_res_or_ind_only) or (cust_is_res_or_ind_only and erp_is_cap_only):
                c_type = "Kapasitör" if cust_is_cap_only else ("Direnç" if cust_specs.resistances else "Endüktans")
                e_type = "Kapasitör" if erp_is_cap_only else ("Direnç" if erp_specs.resistances else "Endüktans")
                contradictions.append(
                    f"Açıklama Çelişkisi (Malzeme Tipi farklı: Müşteri {c_type}, ERP {e_type})"
                )

    # 1. Check Capacitance
    if cust_specs.capacitances and erp_specs.capacitances:
        if not (cust_specs.capacitances & erp_specs.capacitances):
            best_pair = min(
                ((c, e) for c in cust_specs.capacitances for e in erp_specs.capacitances),
                key=lambda p: _calc_ratio(p[0], p[1]),
            )
            ratio = _calc_ratio(best_pair[0], best_pair[1])
            if ratio.is_infinite():
                ratio_note = ", sıfır vs değerli kapasite farkı"
            elif ratio >= 10:
                ratio_note = f", {int(ratio)} kat fark!"
            else:
                ratio_note = ""
            contradictions.append(
                f"Açıklama Çelişkisi (Kritik Değer: Kapasite farklı{ratio_note} - Müşteri {_format_cap(best_pair[0])}, ERP {_format_cap(best_pair[1])})"
            )

    # 2. Check Resistance
    if cust_specs.resistances and erp_specs.resistances:
        if not (cust_specs.resistances & erp_specs.resistances):
            best_pair = min(
                ((r, e) for r in cust_specs.resistances for e in erp_specs.resistances),
                key=lambda p: _calc_ratio(p[0], p[1]),
            )
            ratio = _calc_ratio(best_pair[0], best_pair[1])
            if ratio.is_infinite():
                ratio_note = ", 0R vs değerli direnç farkı"
            elif ratio >= 10:
                ratio_note = f", {int(ratio)} kat fark!"
            else:
                ratio_note = ""
            contradictions.append(
                f"Açıklama Çelişkisi (Kritik Değer: Direnç farklı{ratio_note} - Müşteri {_format_res(best_pair[0])}, ERP {_format_res(best_pair[1])})"
            )

    # 3. Check Inductance
    if cust_specs.inductances and erp_specs.inductances:
        if not (cust_specs.inductances & erp_specs.inductances):
            best_pair = min(
                ((l, e) for l in cust_specs.inductances for e in erp_specs.inductances),
                key=lambda p: _calc_ratio(p[0], p[1]),
            )
            ratio = _calc_ratio(best_pair[0], best_pair[1])
            if ratio.is_infinite():
                ratio_note = ", sıfır vs değerli endüktans farkı"
            elif ratio >= 10:
                ratio_note = f", {int(ratio)} kat fark!"
            else:
                ratio_note = ""
            contradictions.append(
                f"Açıklama Çelişkisi (Kritik Değer: Endüktans farklı{ratio_note} - Müşteri {_format_ind(best_pair[0])}, ERP {_format_ind(best_pair[1])})"
            )

    # 4. Check Voltage
    if cust_specs.voltages and erp_specs.voltages:
        if not (cust_specs.voltages & erp_specs.voltages):
            is_clamp_or_protection = bool(
                re.search(
                    r"\b(clamp|clamping|standoff|tvs|transil|varistor|esd|surge|protection|zener|zen|diode|vrwm|vbr)\b",
                    cust_clean + " " + erp_clean,
                )
            )
            is_capacitor = bool(cust_specs.capacitances or erp_specs.capacitances)

            # When MPN is verified identical, skip secondary clamp/breakdown voltage ratings on non-capacitors
            skip_voltage = mpn_matched and (is_clamp_or_protection or not is_capacitor)
            if not skip_voltage:
                best_pair = min(
                    ((v, e) for v in cust_specs.voltages for e in erp_specs.voltages),
                    key=lambda p: abs(p[0] - p[1]),
                )
                contradictions.append(
                    f"Açıklama Çelişkisi (Voltaj farklı: Müşteri {best_pair[0]}V, ERP {best_pair[1]}V)"
                )

    # 5. Check Tolerance
    if cust_specs.tolerances and erp_specs.tolerances:
        if not (cust_specs.tolerances & erp_specs.tolerances):
            if not mpn_matched:
                best_pair = min(
                    ((t, e) for t in cust_specs.tolerances for e in erp_specs.tolerances),
                    key=lambda p: abs(p[0] - p[1]),
                )
                contradictions.append(
                    f"Açıklama Çelişkisi (Tolerans farklı: Müşteri %{best_pair[0]}, ERP %{best_pair[1]})"
                )

    # 6. Check Package
    if cust_specs.packages and erp_specs.packages:
        if not (cust_specs.packages & erp_specs.packages):
            best_c_pkg = next(iter(cust_specs.packages))
            best_e_pkg = next(iter(erp_specs.packages))
            contradictions.append(
                f"Açıklama Çelişkisi (Kılıf/Paket farklı: Müşteri {best_c_pkg}, ERP {best_e_pkg} - SMT dizgide pad'e oturmaz!)"
            )

    # 7. Check Frequency
    if cust_specs.frequencies and erp_specs.frequencies:
        if not (cust_specs.frequencies & erp_specs.frequencies):
            best_pair = min(
                ((f, e) for f in cust_specs.frequencies for e in erp_specs.frequencies),
                key=lambda p: abs(p[0] - p[1]),
            )
            contradictions.append(
                f"Açıklama Çelişkisi (Frekans farklı: Müşteri {_format_freq(best_pair[0])}, ERP {_format_freq(best_pair[1])})"
            )

    # 8. Check Color
    if cust_specs.colors and erp_specs.colors:
        if not (cust_specs.colors & erp_specs.colors):
            c_col = next(iter(cust_specs.colors))
            e_col = next(iter(erp_specs.colors))
            contradictions.append(
                f"Açıklama Çelişkisi (Renk farklı: Müşteri {c_col}, ERP {e_col})"
            )

    # 9. Check Polarity / Channel
    if cust_specs.polarities and erp_specs.polarities:
        if not (cust_specs.polarities & erp_specs.polarities):
            c_pol = next(iter(cust_specs.polarities))
            e_pol = next(iter(erp_specs.polarities))
            contradictions.append(
                f"Açıklama Çelişkisi (Polarite/Kanal farklı: Müşteri {c_pol}, ERP {e_pol})"
            )

    return contradictions


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


def _format_ind(henry: Decimal) -> str:
    if henry >= Decimal("1"):
        return f"{henry:g}H"
    if henry >= Decimal("1e-3"):
        return f"{henry / Decimal('1e-3'):g}mH"
    if henry >= Decimal("1e-6"):
        return f"{henry / Decimal('1e-6'):g}uH"
    return f"{henry / Decimal('1e-9'):g}nH"


def _format_freq(hz: Decimal) -> str:
    if hz >= Decimal("1e9"):
        return f"{hz / Decimal('1e9'):g}GHz"
    if hz >= Decimal("1e6"):
        return f"{hz / Decimal('1e6'):g}MHz"
    if hz >= Decimal("1e3"):
        return f"{hz / Decimal('1e3'):g}kHz"
    return f"{hz:g}Hz"
