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


@dataclass(frozen=True)
class ElectricalSpecs:
    capacitance_farad: Decimal | None = None
    resistance_ohm: Decimal | None = None
    inductance_henry: Decimal | None = None
    voltage_volt: Decimal | None = None
    tolerance_pct: Decimal | None = None
    mounting: str | None = None  # 'SMD' or 'TH'
    package: str | None = None   # e.g. '0603', '0805', '1206'


@dataclass(frozen=True)
class MultiElectricalSpecs:
    capacitances: frozenset[Decimal] = frozenset()
    resistances: frozenset[Decimal] = frozenset()
    inductances: frozenset[Decimal] = frozenset()
    voltages: frozenset[Decimal] = frozenset()
    tolerances: frozenset[Decimal] = frozenset()
    packages: frozenset[str] = frozenset()
    mountings: frozenset[str] = frozenset()


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
    """Extract electrical specifications from a descriptive string.
    
    Backward-compatible: returns the first extracted specification for each property.
    """
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


def parse_specs_set(text: str) -> MultiElectricalSpecs:
    """Extract all electrical specifications from a descriptive string as sets."""
    clean = " " + text.replace(",", " ").replace(";", " ") + " "
    clean_lower = clean.casefold()

    return MultiElectricalSpecs(
        capacitances=_extract_capacitances(clean_lower),
        resistances=_extract_resistances(clean_lower),
        inductances=_extract_inductances(clean_lower),
        voltages=_extract_voltages(clean_lower),
        tolerances=_extract_tolerances(clean_lower),
        packages=_extract_packages(clean),
        mountings=_extract_mountings(clean),
    )


def _extract_capacitances(text: str) -> frozenset[Decimal]:
    caps: set[Decimal] = set()
    # Pattern 1: standard number + unit
    for match in re.finditer(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(pf|nf|uf|µf|\u03bcf|mf|f)(?![a-z0-9])", text):
        num, unit = match.groups()
        caps.add(Decimal(num) * _CAP_MULTIPLIERS.get(unit, Decimal("1")))
    # Pattern 2: R-code like 4u7 -> 4.7 uF
    for match_rcode in re.finditer(r"(?<![a-z0-9])(\d+)(u|p|n|µ|\u03bc)(\d+)(?![a-z0-9])", text):
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit = "uf" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}f"
        if unit in _CAP_MULTIPLIERS:
            caps.add(Decimal(num_str) * _CAP_MULTIPLIERS[unit])
    return frozenset(caps)


def _extract_capacitance(text: str) -> Decimal | None:
    match = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(pf|nf|uf|µf|\u03bcf|mf|f)(?![a-z0-9])", text)
    if match:
        num, unit = match.groups()
        return Decimal(num) * _CAP_MULTIPLIERS.get(unit, Decimal("1"))
    match_rcode = re.search(r"(?<![a-z0-9])(\d+)(u|p|n|µ|\u03bc)(\d+)(?![a-z0-9])", text)
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
    for match_rcode in re.finditer(r"(?<![a-z0-9])(\d+)([rkmo])(\d+)(?![a-z0-9])", text):
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit_map = {"r": "r", "k": "k", "m": "m", "o": "r"}
        unit = unit_map[unit_char]
        res.add(Decimal(num_str) * _RES_MULTIPLIERS[unit])

    # Pattern 2: number + unit (k, kohm, ohm, r, etc.)
    for match in re.finditer(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(kohms?|kω|mohms?|mω|ohms?|ω|mr|k|r)(?![a-z0-9])",
        text,
    ):
        num, unit = match.groups()
        if unit == "k":
            res.add(Decimal(num) * Decimal("1e3"))
        elif unit in _RES_MULTIPLIERS:
            res.add(Decimal(num) * _RES_MULTIPLIERS[unit])

    # Pattern 3: explicit "xxx ohm" or "xxx R"
    for match_ohm in re.finditer(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:ohm|r)(?![a-z0-9])", text):
        res.add(Decimal(match_ohm.group(1)))

    return frozenset(res)


def _extract_resistance(text: str) -> Decimal | None:
    match_rcode = re.search(r"(?<![a-z0-9])(\d+)([rkmo])(\d+)(?![a-z0-9])", text)
    if match_rcode:
        whole, unit_char, frac = match_rcode.groups()
        num_str = f"{whole}.{frac}"
        unit_map = {"r": "r", "k": "k", "m": "m", "o": "r"}
        unit = unit_map[unit_char]
        return Decimal(num_str) * _RES_MULTIPLIERS[unit]

    match = re.search(
        r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(kohms?|kω|mohms?|mω|ohms?|ω|mr|k|r)(?![a-z0-9])",
        text,
    )
    if match:
        num, unit = match.groups()
        if unit == "k":
            return Decimal(num) * Decimal("1e3")
        if unit in _RES_MULTIPLIERS:
            return Decimal(num) * _RES_MULTIPLIERS[unit]

    match_ohm = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(?:ohm|r)(?![a-z0-9])", text)
    if match_ohm:
        return Decimal(match_ohm.group(1))

    return None


def _extract_inductances(text: str) -> frozenset[Decimal]:
    inds: set[Decimal] = set()
    for match in re.finditer(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(nh|uh|µh|\u03bch|mh|h)(?![a-z0-9])", text):
        num, unit = match.groups()
        inds.add(Decimal(num) * _IND_MULTIPLIERS.get(unit, Decimal("1")))
    for match_code in re.finditer(r"(?<![a-z0-9])(\d+)(u|µ|\u03bc|m|n)h?(\d+)(?![a-z0-9])", text):
        whole, unit_char, frac = match_code.groups()
        num_str = f"{whole}.{frac}"
        unit = "uh" if unit_char in ("u", "µ", "\u03bc") else f"{unit_char}h"
        if unit in _IND_MULTIPLIERS:
            inds.add(Decimal(num_str) * _IND_MULTIPLIERS[unit])
    return frozenset(inds)


def _extract_inductance(text: str) -> Decimal | None:
    match = re.search(r"(?<![a-z0-9])(\d+(?:\.\d+)?)\s*(nh|uh|µh|\u03bch|mh|h)(?![a-z0-9])", text)
    if match:
        num, unit = match.groups()
        return Decimal(num) * _IND_MULTIPLIERS.get(unit, Decimal("1"))
    match_code = re.search(r"(?<![a-z0-9])(\d+)(u|µ|\u03bc|m|n)h?(\d+)(?![a-z0-9])", text)
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

    # 1. Check Capacitance
    if cust_specs.capacitances and erp_specs.capacitances:
        if not (cust_specs.capacitances & erp_specs.capacitances):
            best_pair = min(
                ((c, e) for c in cust_specs.capacitances for e in erp_specs.capacitances),
                key=lambda p: max(p[0], p[1]) / min(p[0], p[1]),
            )
            ratio = max(best_pair[0], best_pair[1]) / min(best_pair[0], best_pair[1])
            if not mpn_matched or ratio >= 10:
                ratio_note = f", {int(ratio)} kat fark!" if ratio >= 10 else ""
                contradictions.append(
                    f"Açıklama Çelişkisi (Kritik Değer: Kapasite farklı{ratio_note} - Müşteri {_format_cap(best_pair[0])}, ERP {_format_cap(best_pair[1])})"
                )

    # 2. Check Resistance
    if cust_specs.resistances and erp_specs.resistances:
        if not (cust_specs.resistances & erp_specs.resistances):
            best_pair = min(
                ((r, e) for r in cust_specs.resistances for e in erp_specs.resistances),
                key=lambda p: max(p[0], p[1]) / min(p[0], p[1]),
            )
            ratio = max(best_pair[0], best_pair[1]) / min(best_pair[0], best_pair[1])
            if not mpn_matched or ratio >= 10:
                ratio_note = f", {int(ratio)} kat fark!" if ratio >= 10 else ""
                contradictions.append(
                    f"Açıklama Çelişkisi (Kritik Değer: Direnç farklı{ratio_note} - Müşteri {_format_res(best_pair[0])}, ERP {_format_res(best_pair[1])})"
                )

    # 3. Check Voltage
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

    # 4. Check Tolerance
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

    # 5. Check Package
    if cust_specs.packages and erp_specs.packages:
        if not (cust_specs.packages & erp_specs.packages):
            if not mpn_matched:
                best_c_pkg = next(iter(cust_specs.packages))
                best_e_pkg = next(iter(erp_specs.packages))
                contradictions.append(
                    f"Açıklama Çelişkisi (Kılıf/Paket farklı: Müşteri {best_c_pkg}, ERP {best_e_pkg} - SMT dizgide pad'e oturmaz!)"
                )

    if contradictions:
        return contradictions

    # If any specific electrical spec was verified to match, we consider it compatible!
    has_matching_spec = (
        bool(cust_specs.capacitances & erp_specs.capacitances)
        or bool(cust_specs.resistances & erp_specs.resistances)
        or bool(cust_specs.voltages & erp_specs.voltages)
        or bool(cust_specs.inductances & erp_specs.inductances)
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


def _format_ind(henry: Decimal) -> str:
    if henry >= Decimal("1"):
        return f"{henry:g}H"
    if henry >= Decimal("1e-3"):
        return f"{henry / Decimal('1e-3'):g}mH"
    if henry >= Decimal("1e-6"):
        return f"{henry / Decimal('1e-6'):g}uH"
    return f"{henry / Decimal('1e-9'):g}nH"
