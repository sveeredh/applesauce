# --- libraries ---
from tabulate import tabulate
import re
import time
import os
from collections import namedtuple
import pandas as pd
import sys
import contextlib
import io
from get_data import _unblock_csv_file, download_and_rename_ti_specs, download_and_rename_ti_zener_specs, download_and_rename_aos_specs, download_diodes_specs, download_nexperia_specs, download_littelfuse_specs, download_jiangsu_specs, download_semtech_specs
from find_package_matches import find_package_matches
from parsing import safe_strip, is_blocked_competitor_package, requires_exact_package_match, normalize_package, normalize_structured_package, to_numeric_val, normalize_aos_package, normalize_amazing_package, normalize_diodes_package, normalize_nexperia_package, normalize_littelfuse_package, normalize_jiangsu_package, normalize_semtech_package
from ti_scrape import load_excel_data, fetch_ti_specs_from_excel
from ti_zener_scrape import fetch_ti_zener_specs_from_excel
from central_scrape import fetch_central_specs, central_has_part
from comchip_scrape import fetch_comchip_specs, comchip_has_part
from aos_scrape import fetch_aos_specs_from_excel
from amazing_scrape import fetch_amazing_specs_from_excel
from diodes_scrape import fetch_diodes_specs_from_excel
from nexperia_scrape import fetch_nexperia_specs_from_excel
from littelfuse_scrape import fetch_littelfuse_specs_from_excel
from jiangsu_scrape import fetch_jiangsu_specs_from_excel
from semtech_scrape import fetch_semtech_specs_from_excel
from stm_scrape import fetch_stm_specs_from_excel
from panjit_scrape import fetch_panjit_specs_from_excel
from onsemi_scrape import fetch_onsemi_specs_from_excel
from anbon_scrape import fetch_anbon_specs
from anbon_scrape import fetch_anbon_specs
from diotec_scrape import fetch_diotec_specs_from_excel
from eaton_scrape import fetch_eaton_specs_from_excel
from eic_scrape import fetch_eic_specs_from_excel
from evvo_scrape import fetch_evvo_specs_from_excel, evvo_has_part
from goodark_scrape import fetch_goodark_specs_from_excel
from infineon_scrape import fetch_infineon_specs_from_excel
from vishay_scrape import fetch_vishay_specs_from_excel
from yangjie_scrape import fetch_yangjie_specs_from_excel, load_yangjie_xls
from mcc_scrape import fetch_mcc_specs_from_excel
from galaxy_scrape import fetch_galaxy_specs_from_excel
from leshan_scrape import fetch_leshan_specs_from_excel
from yenyo_scrape import fetch_yenyo_specs_from_excel
from nichtek_scrape import fetch_nichtek_specs_from_excel
from jjm_scrape import fetch_jjm_specs_from_excel, load_jjm_xlsx
from inpaq_scrape import fetch_inpaq_specs_from_excel

CACHE_DIR = ".applesauce_cache"

# ============================================================================
# Canonical package classification (DigiKey <-> TI)
#
# Each package+pin combo this program understands maps to a canonical code
# like "DFN0603_2", "X2SON_4", "SOT886_6". DigiKey strings are read primarily
# from "Supplier Device Package", falling back to "Package / Case" for pin
# disambiguation when needed. TI strings are read from the Package name +
# Pin count columns. Anything that doesn't match a known rule returns None
# and falls back to the legacy normalize_package() matching.
# ============================================================================

def _has_pin_marker(text, pin_n):
    """True if `text` has pin_n indicated via -N suffix, N- prefix, or a
    trailing token containing pin_n (e.g. 'P2', '-2L')."""
    t = text.upper()
    n = str(pin_n)
    # A digit after a decimal point belongs to a body dimension, not a pin
    # count ("DFNWB1.0x0.6-02L" is 2 pins, not 6).
    if re.search(rf'(^|[^0-9.]){n}-', t):
        return True
    # Pin counts are sometimes zero padded ("-02L", "-06L").
    if re.search(rf'-0*{n}([^0-9]|$)', t):
        return True
    if re.search(rf'-0*{n}[A-Z]', t):
        return True
    if re.search(rf'[A-Z]{n}(\b|$)', t):
        return True
    return False


def _has_pin_count_anywhere(text, pin_n):
    """Looser check: pin_n appears as a standalone-ish number in the string."""
    return bool(re.search(rf'(?<![0-9]){pin_n}(?![0-9])', text.upper()))


def _sot9x3_pin_count(pkg_str):
    """
    Pin count for a competitor package in TI's SOT-9X3 family.

    TI names this family SOT-9X3, where X is the pin count, so the middle digit
    of SOT-9x3 decides which TI variant a competitor part may cross to
    ('SOT-953' -> 5). SOT-923 is the exception: it is the 3-lead member of the
    family, so it counts as 3 rather than as its middle digit. An explicit
    lead/pin count in the string wins.
    Returns None if the string carries no usable pin information.
    """
    s = str(pkg_str or "").upper()

    m = re.search(r'(?:^|[^0-9A-Z])(\d+)\s*-?\s*(?:L\b|LD\b|LEADS?\b|PINS?\b)', s)
    if m and 2 <= int(m.group(1)) <= 8:
        return int(m.group(1))

    flat = re.sub(r'[\s\-_]+', '', s)

    # TI's own "SOT-9X3-{pins}" form (produced by _canonical_to_ti_pkg_with_pins)
    m = re.match(r'^SOT9X3(\d+)', flat)
    if m:
        return int(m.group(1))

    # Middle digit of SOT-9x3 is the pin count, except SOT-923 (3 leads).
    m = re.match(r'^SOT9(\d)3', flat)
    if m:
        return 3 if m.group(1) == '2' else int(m.group(1))

    return None


def _has_dim(text, *dim_variants):
    """True if any of the given dimension strings (e.g. '0.6x0.3', '0603')
    appear in text, accommodating separators like 'x', '.', '-', ' '."""
    t = text.upper().replace(' ', '')
    for variant in dim_variants:
        v = variant.upper().replace(' ', '')
        # Build a loose regex: digits/dots literally, x/X as separator, allow optional . and -
        pattern = re.escape(v).replace('X', '[Xx]').replace(r'\.', r'\.?')
        if re.search(pattern, t):
            return True
    return False


def _sot5x3_pin_count(pkg_str):
    """
    Pin count for a competitor package in TI's SOT-5X3 family.

    TI names this family SOT-5X3, where X is the pin count, so the pin count is
    what decides which TI variant a competitor part may cross to:
      - an explicit lead/pin count in the string wins ('SOT523-3L' -> 3)
      - otherwise the middle digit of SOT-5x3 is the pin count ('SOT-563' -> 6)
    Returns None if the string carries no usable pin information.
    """
    s = str(pkg_str or "").upper()

    # Explicit lead/pin count: "SOT523-3L", "SOT-563 6L", "SOT553-5PIN".
    # A separator must precede the number so a trailing 'L' on the body code
    # ("SOT553L") is not mistaken for a lead count.
    m = re.search(r'(?:^|[^0-9A-Z])(\d+)\s*-?\s*(?:L\b|LD\b|LEADS?\b|PINS?\b)', s)
    if m and 2 <= int(m.group(1)) <= 8:
        return int(m.group(1))

    flat = re.sub(r'[\s\-_]+', '', s)

    # TI's own "SOT-5X3-{pins}" form (produced by _canonical_to_ti_pkg_with_pins)
    m = re.match(r'^SOT5X3(\d+)', flat)
    if m:
        return int(m.group(1))

    # Middle digit of SOT-5x3 is the pin count (SOD-563 is the one SOD in
    # this family; SOD-523 / SOD-323 are separate packages).
    if flat == "SOT523":
        return "3"
    m = re.match(r'^SOT5(\d)3', flat)
    if m:
        return int(m.group(1))
    if flat.startswith("SOD563"):
        return 6

    if flat.startswith("SOT665"):
        return 2
    # SC-89 / SOT-89 write their pin count as a trailing number ("SC-89-3"),
    # so an explicit count wins over the family default below.
    m = re.match(r'^(?:SOT89|SC89)(\d+)[A-Z]*$', flat)
    if m and 2 <= int(m.group(1)) <= 8:
        return int(m.group(1))
    
    if flat.startswith("SOT89") or flat.startswith("SC89"):
        return 5
    return None


def _classify_digikey_package(supplier_pkg, case_pkg):
    """
    Returns a canonical code (e.g. 'DFN0603_2') for a DigiKey part based on
    Supplier Device Package (primary) and Package / Case (backup, for pin info).
    Returns None if no rule matches.
    """
    sup = str(supplier_pkg or "").strip().upper()
    case = str(case_pkg or "").strip().upper()
    combined = f"{sup} {case}"

    if not sup and not case:
        return None

    # A cell may list several packages ("SOT23, SOT323"), and Comchip writes a
    # single package as "<size code>/<package name>" ("0603C/SOD-523F"). Take
    # comma-separated entries left to right, but read each slash group right to
    # left: the package name after the slash is the real package, and the size
    # code before it would otherwise hijack a rule below ("0603" -> DFN0603).
    if ',' in sup or '/' in sup:
        for _pkg_token in sup.split(','):
            for _slash_part in reversed(_pkg_token.split('/')):
                _slash_part = _slash_part.strip()
                if not _slash_part:
                    continue
                _token_canonical = _classify_digikey_package(_slash_part, case_pkg)
                if _token_canonical:
                    return _token_canonical
        return None

    # Pin-bearing DFN forms first — they carry an explicit pin count.
    _hyphen_dfn = re.match(r'(?:(?:U|W|X\d?)-?)?[DQ]FN(\d+)-(\d+)[A-Z]*(?:\([^)]*\))?$', sup.replace(' ', ''), re.IGNORECASE)
    if _hyphen_dfn:
        _dfn_fam_map_h = {
            '0603': 'DFN0603', '1006': 'DFN1006', '1110': 'DFN1110',
            '1610': 'DFN1610', '1616': 'DFN1616', '2020': 'DFN2020', '2510': 'DFN2510', '3030': 'DFN3030'
        }
        _fam_h = _dfn_fam_map_h.get(_hyphen_dfn.group(1))
        if _fam_h:
            _n = int(_hyphen_dfn.group(2))
            # DFN0603 6-pin is TI's X2SON 6-pin (DPFR), 4-pin is X2SON 4-pin (DPWR)
            if _fam_h == 'DFN0603' and _n == 6:
                return "X2SON_6"
            if _fam_h == 'DFN0603' and _n == 4:
                return "X2SON_4"
            return f"{_fam_h}_{_n}"

    _amazing_dfn = re.match(r'(?:U|W|X\d?)?[DQ]FN(\d+)P(\d+)[A-Z]*$', sup.replace(' ', ''), re.IGNORECASE)
    if _amazing_dfn:
        dims = _amazing_dfn.group(1)
        pins = int(_amazing_dfn.group(2))
        _dfn_family_map = {
            '0603': 'DFN0603', '1006': 'DFN1006', '1110': 'DFN1110',
            '1610': 'DFN1610', '1616': 'DFN1616', '2020': 'DFN2020', '2510': 'DFN2510', '3030': 'DFN3030'
        }
        family = _dfn_family_map.get(dims)
        if family:
            if family == 'DFN0603' and pins == 6:
                return "X2SON_6"
            if family == 'DFN0603' and pins == 4:
                return "X2SON_4"
            return f"{family}_{pins}"

    # 6-UFDFN is TI's 6-pin DFN1616 (VEBR).
    _ufdfn = re.match(r'^(\d+)U?F?DFN', re.sub(r'[\s\-_]+', '', sup.upper()))
    if _ufdfn and int(_ufdfn.group(1)) == 6:
        return "DFN1616_6"

    # ===================== TI package equivalence aliases =====================
    # Checked first, most specific first. `flat` strips spaces AND hyphens so
    # "SOT-323-6L" -> "SOT3236L"; trailing letters are tolerated everywhere.
    flat = re.sub(r'[\s\-_]+', '', sup.upper())

    # -- 6-pin SC70 family (must precede 3-pin SC70 checks) --
    if re.match(r'^(SOT363|TSSOP6|SC706|SOT3236|SC886|SC706)[A-Z0-9]*$', flat):
        return "SC706_6"
    # -- 3-pin SC70 family --
    if re.match(r'^(SOTSC70|SOT3233)[A-Z0-9]*$', flat):
        return "SC703_3"
    # SC70 carries its pin count: SC70/SC70-3 -> 3-pin, SC70-6 -> 6-pin.
    # Any other count (SC70-5L etc.) is not a TI package we carry.
    _sc70 = re.match(r'^SC70(\d*)[A-Z]*$', flat)
    if _sc70:
        _p = _sc70.group(1)
        if _p in ('', '3'):
            return "SC703_3"
        if _p == '6':
            return "SC706_6"
        return None
    # -- SOT-9X3 (X is the pin count, taken from the competitor package) --
    if re.match(r'^(SOT9\d3|SOT9X3\d*)[A-Z0-9]*$', flat):
        _p9x3 = _sot9x3_pin_count(sup)
        return f"SOT9X3_{_p9x3}" if _p9x3 else None
    if re.match(r'^SOT3[A-Z]*$', flat):
        return "SOT9X3_3"
    # -- SOT-5X3 (X is the pin count, taken from the competitor package).
    # SOD-523 / SOD-323 are their own packages and are handled below - only
    # SOD-563 belongs to this family. --
    if re.match(r'^(SOT5\d3|SOD563|SOT5X3\d*|SOT665|SOT89|SC89)[A-Z0-9]*$', flat):
        _p5x3 = _sot5x3_pin_count(sup)
        return f"SOT5X3_{_p5x3}" if _p5x3 else None
    # -- SOD523 / SOD323 --
    if re.match(r'^(SC79|SOD5232|SOD523)[A-Z0-9]*$', flat):
        return "SOD523_2"
    if re.match(r'^(SC90|SOD323)[A-Z0-9]*$', flat):
        return "SOD323_2"
    # -- SOT-23 family --
    if re.match(r'^(TO236|SC59|SOT233)[A-Z0-9]*$', flat):
        return "SOT233_3"
    if re.match(r'^(SOT143|SOT1434|SOT1234|SOT234|SOT24|TSOT24)[A-Z0-9]*$', flat):
        return "SOT234_4"
    if re.match(r'^(SOT25|TSOT25|TSOP5|SOT235)[A-Z0-9]*$', flat):
        return "SOT235_5"
    if re.match(r'^(SOT457|TSOP6|SOT26|TSOT26|SC74|SOT236)[A-Z0-9]*$', flat):
        return "SOT236_6"
    # -- DFN1006 --
    # SOT-883 is a 3-lead package, written with or without the lead count.
    if re.match(r'^(SOT883|LLP10063)[A-Z0-9]*$', flat):
        return "DFN1006_3"
    if re.match(r'^(X1SON|SOD882)[A-Z]*$', flat):
        return "DFN1006_2"
    # -- DFN1616 6-pin -> TI DFN1616 6-pin (VEBR) --
    if re.match(r'^DFN16166[A-Z]*$', flat):
        return "DFN1616_6"
    if re.match(r'^USON1616\d*[A-Z]*$', flat):
        return "USON1616_6"
    _uson = re.match(r'^USON(\d+)[A-Z]*$', flat)
    if _uson:
        return f"USON_{int(_uson.group(1))}"
    if re.match(r'^USON[A-Z]*$', flat):
        return "USON_6"

    # -- DFN0603 / X2SON --
    # Onsemi X2DFNW2 = 2-pin DFN1006
    if flat == "X2DFNW2":
        return "DFN1006_2"
    _x2 = re.match(r'^X2SON(\d+)[A-Z]*$', flat)
    if re.match(r'^X2?DFN2L?$', flat):
        return "DFN1006_2"
    if _x2:
        _n = int(_x2.group(1))
        return f"X2SON_{_n}" if _plausible_pin_for_family("X2SON", _n) else None
    if re.match(r'^(X2SON|DFN0606|X2DFN8084|X2DFN|DFN4L)[A-Z0-9]*$', flat):
        return "X2SON_4"
    # -- DFN2510 --
    if re.match(r'^(QFN10L|QFN10|UDFN10|DFN10)[A-Z]*$', flat):
        return "DFN2510_10"
    if re.match(r'^XSON6[A-Z0-9]*$', flat):
        return "USON_6"
    if re.match(r'^(QFN6L|QFN6)[A-Z0-9]*$', flat):
        return "DFN2510_6"
    # =========================================================================

    # --- WLCSP quoted by body size, e.g. AOS's "WLCSP0.63x0.33A-3". The body
    # size names the TI DFN family (0.63x0.33 rounds to 0.6x0.3 -> DFN0603) and
    # the trailing number is the bump count. Read from the raw string, before
    # normalisation welds the dimension and the lead count together. ---
    if "WLCSP" in sup or "CSP" in sup:
        _csp_dim = re.search(r'(\d+\.\d+)\s*X\s*(\d+\.\d+)', sup)
        _csp_pins = re.search(r'[-\s](\d+)L?\s*$', sup)
        if _csp_dim and _csp_pins:
            _csp_family = {
                (0.6, 0.3): 'DFN0603', (1.0, 0.6): 'DFN1006', (1.1, 1.0): 'DFN1110',
                (1.6, 1.6): 'DFN1616', (2.0, 2.0): 'DFN2020', (2.5, 1.0): 'DFN2510',
                (3.0, 3.0): 'DFN3030',
            }.get((round(float(_csp_dim.group(1)), 1), round(float(_csp_dim.group(2)), 1)))
            if _csp_family:
                _csp_n = int(_csp_pins.group(1))
                if _csp_family == 'DFN0603' and _csp_n == 6:
                    return "X2SON_6"
                return f"{_csp_family}_{_csp_n}"

    # --- SOD523 (must be checked before DFN0603, since some DigiKey entries
    # mislabel it like "0603/SOD-523") ---
    if _has_dim(combined, "SOD523", "SOD-523"):
        return "SOD523_2"

    # --- SOD323 ---
    if _has_dim(combined, "SOD323", "SOD-323"):
        return "SOD323_2"

    # --- DFN1608, e.g. Comchip's "0603/DFN1608". The 0603 here is the IMPERIAL
    # code for a 1.6 x 0.8 mm body, not the metric 0603 (0.6 x 0.3 mm) that maps
    # to TI's DFN0603 below. The explicit DFN1608 settles it, so it is checked
    # first. ---
    if "DFN1608" in sup.replace('-', '').replace(' ', ''):
        return None
    
    # --- DFN0603 (2-pin) / X2SON (4-pin) ---
    # Onsemi-specific DFN package aliases
    sup_flat = re.sub(r'[\s\-_]+', '', sup.upper())

    if sup_flat.startswith("X2DFNW2"):
        return "DFN1006_2"

    if sup_flat.startswith("X3DFN2"):
        return "DFN0603_2"

    if sup_flat.startswith("XDFNW2") and "1.00X0.60X0.50" in sup_flat:
        return "DFN1006_2"
    # 0603-style dims: 0.6x0.3, 0.62x0.32, or literal '0603' with pin markers.
    is_0603_dim = _has_dim(sup, "0.6x0.3", ".6x0.3", "0.6x.3", "0.62x0.32", "0808")
    is_0603_token = bool(re.search(r'(^|[^0-9])0603([^0-9]|$)', sup)) or \
                    bool(re.search(r'^0?603$', sup.replace(' ', '')))
    is_0808_dim = _has_dim(sup, "0.8x0.8", "0808")

    if is_0808_dim and _has_pin_marker(combined, 4):
        return "X2SON_4"
    if "X2SON" in sup:
        return "X2SON_4"
    # DFN0606 / X2-DFN808-4 / DFN-4L -> DFN0603 4-pin (X2SON)
    if re.search(r'DFN-?0606', sup) or re.search(r'X2-?DFN', sup) or re.search(r'DFN-?4L', sup):
        return "X2SON_4"
    # A bare "DFN-2" / "DFN-2L" states a pin count but no body size. The only
    # 2-pin DFN in TI's catalogue is DFN0603, so that is what it crosses to.
    if re.match(r'^DFN2L?$', sup_flat):
        return "DFN0603_2"
    
    if (is_0603_dim or is_0603_token) and not _has_dim(sup, "0201"):
        # Exclude if Package/Case explicitly indicates a different pin count (e.g. 3-pin)
        if _has_pin_marker(case, 3) or _has_pin_count_anywhere(case, 3):
            return None  # not our 2-pin DFN0603; some other pin count
        return "DFN0603_2"

    if _has_dim(sup, "0201"):
        if _has_pin_marker(case, 3) or _has_pin_count_anywhere(case, 3):
            return None
        return "DFN0603_2"

    # --- SOD882 -> DFN1006 (same footprint) ---
    if _has_dim(sup, "SOD882", "SOD-882", "SOD 882"):
        return "DFN1006_2"

    # --- SOD-923 is its own 2-lead package and stays SOD923. Covers Diodes'
    # plain "SOD923" and Comchip's "0402C/SOD-923F" flat-lead variant alike. ---
    if "SOD923" in sup.replace('-', '').replace(' ', ''):
        return "SOD923_2"

    # --- 1005 metric -> DFN1006 ---
    if re.match(r'^1005$', sup.replace(' ', '')):
        return "DFN1006_2"

    # --- DFN1006 (1x0.6) - the lead count in the string wins, so
    # "Q05C/DFN1006-5L" is 5-pin. 2-pin only when the string does not say. ---
    if _has_dim(sup, "1x0.6", "1.0x0.6", "1006"):
        for _dfn1006_pins in (6, 5, 4, 3):
            if _has_pin_marker(combined, _dfn1006_pins):
                return f"DFN1006_{_dfn1006_pins}"
        return "DFN1006_2"

    # --- DFN{dims}-{pins} hyphen style e.g. DFN2020-3, DFN1006-2 ---

    # --- Amazing-style DFN{dims}P{pins}[letter] e.g. DFN2020P3E ---
    _amazing_dfn = re.match(r'(?:U|W|X\d?)?[DQ]FN(\d+)P(\d+)[A-Z]*$', sup.replace(' ', ''), re.IGNORECASE)
    if _amazing_dfn:
        dims = _amazing_dfn.group(1)
        pins = int(_amazing_dfn.group(2))
        # Map dims to our canonical family
        _dfn_family_map = {
            '0603': 'DFN0603', '1006': 'DFN1006', '1110': 'DFN1110',
            '1610': 'DFN1610', '1616': 'DFN1616', '2020': 'DFN2020', '2510': 'DFN2510', '3030': 'DFN3030'
        }
        family = _dfn_family_map.get(dims)
        if family:
            if family == 'DFN0603' and pins == 6:
                return "X2SON_6"
            if family == 'DFN0603' and pins == 4:
                return "X2SON_4"
            return f"{family}_{pins}"

    # --- DFN1616 (1.6x1.6, 6-pin, suffix VEBR) ---
    if _has_dim(sup, "1.6x1.6", "1616") and _has_pin_marker(combined, 6):
        return "DFN1616_6"

    # --- DFN1110 (3-pin) ---
    if _has_dim(sup, "1110", "1.1x1.0", "1.1x1"):
        return "DFN1110_3"

    # --- WDFN-6: the bare form carries a pin count but no dimensions. The
    # 6-pin WDFN body is 2 x 2, so it crosses to DFN2020. ---
    if re.match(r'^WDFN6L?$', flat):
        return "DFN2020_6"
    
    # --- DFN2020 (2x2) - the lead count in the string wins, so "DFN2x2-3L" is
    # 3-pin. 6-pin only when the string does not say otherwise. ---
    if _has_dim(sup, "2x2", "2020"):
        for _dfn2020_pins in (3, 6):
            if _has_pin_marker(combined, _dfn2020_pins):
                return f"DFN2020_{_dfn2020_pins}"
        return "DFN2020_6"

    # --- XSON6 / QFN-6L -> DFN2510 6-pin ---
    if re.search(r'XSON-?6', sup) or re.search(r'QFN-?6L', sup):
        return "DFN2510_6"

    # --- QFN-10L / uDFN-10 -> DFN2510 10-pin ---
    if re.search(r'QFN-?10L?', sup) or re.search(r'u?DFN-?10', sup, re.IGNORECASE):
        return "DFN2510_10"

    # --- DFN2510 (10-pin) ---
    if _has_dim(sup, "2510", "2.5x1.0", "2.5x1"):
        return "DFN2510_10"

    # --- SOT-886 is TI's 6-pin USON (DRYR) - it crosses to that and nothing else ---
    if _has_dim(sup, "1.45x1.0", "1.45x1", "886") and _has_pin_marker(combined, 6):
        return "USON_6"
    if "SOT886" in sup.replace('-', '') or "SOT-886" in sup:
        return "USON_6"
    if "XSON" in sup and _has_pin_marker(combined, 6):
        return "USON_6"

    # --- USON (6-pin, 1.6x1.6, suffix DPKR) ---
    if _has_dim(sup, "1.6x1.6") and _has_pin_marker(combined, 6) and "USON" in combined:
        return "USON_6"

    # --- DFN3030 (3x3) / 8-MSOP - 8-pin is DRBR, 6-pin is DRSR ---
    if _has_dim(sup, "3x3", "3030") or "MSOP-8" in sup:
        if _has_pin_marker(combined, 6):
            return "DFN3030_6"
        return "DFN3030_8"

    # --- DSBGA (4-pin) ---
    if "DSBGA" in sup:
        return "DSBGA_4"

    # --- SC70-3 / SOT323 ---
    if re.search(r'SC-?70-?3', sup) or "SOT323" in sup.replace('-', ''):
        return "SC703_3"

    # --- SC70-6 / SOT363 / SC-88 ---
    # --- SC70-6 / SOT363 / SC-88 ---
    # SC-88 is the 6-pin body but SC-88A is the 5-pin one (SOT-353), so the
    # test has to stop at the digits rather than matching SC88 as a substring.
    _sc_flat = sup.replace('-', '').replace(' ', '')
    if re.search(r'SC-?70-?6', sup) or "SOT363" in _sc_flat or re.search(r'SC88(?![A-Z0-9])', _sc_flat):
        return "SC706_6"

    # --- SC70-5 / SOT353 / SC-88A ---
    if re.search(r'SC-?70-?5', sup) or "SOT353" in _sc_flat or re.search(r'SC88A(?![A-Z0-9])', _sc_flat):
        return "SC705_5"

    # --- SC70 with no pin number -> default 3-pin; SC70 with other pin count -> no match ---
    if re.search(r'SC-?70', sup):
        pin_m = re.search(r'SC-?70-?(\d+)', sup)
        if not pin_m:
            return "SC703_3"  # bare SC70, default to 3-pin
        # Any other pin count (4, 5, etc.) is not a TI package we carry
        return None

    # --- SOT24/TSOT24 -> SOT-23-4 ---
    if re.search(r'TSOT-?24\b', sup) or re.search(r'SOT-?24\b', sup):
        return "SOT234_4"

    # --- SOT25/TSOT25/TSOP-5 -> SOT-23-5 ---
    if re.search(r'TSOT-?25', sup) or re.search(r'SOT-?25', sup) or re.search(r'TSOP-?5', sup):
        return "SOT235_5"
    

    # --- SOT457/TSOP-6/SOT26/TSOT26 -> SOT-23-6 ---
    if re.search(r'SOT-?457', sup) or re.search(r'TSOP-?6', sup) or re.search(r'SOT-?26', sup) or re.search(r'TSOT-?26', sup):
        return "SOT236_6"
    if re.search(r'SC-?74', sup):
            return "SOT235_6"

    # --- SOT-23 family ---
    if "SOT23" in sup.replace('-', '').replace(' ', '') or "SOT143" in sup.replace('-', '').replace(' ', '') or re.search(r'SOT-?123-?4', sup):
        sup_clean = sup.replace('-', '').replace(' ', '')
        if "SOT143" in sup_clean or re.search(r'SOT-?123-?4', sup):
            return "SOT234_4"
        if _has_pin_marker(combined, 5):
            return "SOT235_5"
        if _has_pin_marker(combined, 6):
            return "SOT236_6"
        if _has_pin_marker(combined, 4):
            return "SOT234_4"
        return "SOT233_3"

    # --- SOT-5X3 (X is the pin count, taken from the competitor package) ---
    if re.search(r'SOT-?5(\d|X)3', sup) or re.search(r'SOD-?563', sup) or re.search(r'SOT-?665', sup):
        _p5x3 = _sot5x3_pin_count(sup)
        return f"SOT5X3_{_p5x3}" if _p5x3 else None

    # --- SOT-9X3 (X is the pin count, taken from the competitor package) ---
    if re.search(r'SOT-?9(\d|X)3', sup):
        _p9x3 = _sot9x3_pin_count(sup)
        return f"SOT9X3_{_p9x3}" if _p9x3 else None

    # --- SOT-9X3 (1x1, 3-pin) ---
    if _has_dim(sup, "1x1") and _has_pin_marker(combined, 3):
        return "SOT9X3_3"

    # --- UQFN family ---
    if "UQFN" in sup or "U-QFN" in sup:
        if _has_dim(sup, "3.5x1.35", "3.5x1.4"):
            return "UQFN_14"
        if _has_dim(sup, "2.0x1.5", "2x1.5"):
            return "UQFN_10"
        if "10-UQFN" in case.replace(' ', '') or _has_pin_marker(combined, 10):
            return "UQFN_10"
        if _has_pin_marker(combined, 14):
            return "UQFN_14"

    # --- WQFN (4x4, 12-pin) ---
    if "WQFN" in sup:
        if _has_dim(sup, "4x4"):
            return "WQFN_12"
        if _has_pin_marker(combined, 12):
            return "WQFN_12"

    # --- WSON: 15-pin crosses to TI's WSON (DSMR); 6-pin crosses to TI's
    # 6-pin DFN3030 (DRSR), which is what that package was renamed to. ---
    if "WSON" in sup:
        if _has_dim(sup, "3x3") and _has_pin_marker(combined, 6):
            return "DFN3030_6"
        if _has_pin_marker(combined, 15) or "15-SON" in case.replace(' ', ''):
            return "WSON_15"
        if _has_pin_marker(combined, 6):
            return "DFN3030_6"

    if "15-SON" in sup.replace(' ', '') or "15SON" in sup.replace(' ', '').replace('-', ''):
        return "WSON_15"

    # --- TO-236 / SC-59 variants -> SOT-23-3. TO-263 (D2PAK) is a power
    # package, not a SOT-23, and has no TI equivalent here. ---
    if re.match(r'TO-?236[A-Z0-9-]*$', sup.replace(' ', ''), re.IGNORECASE) or re.search(r'SC-?59', sup):
        return "SOT233_3"

    # --- DFN-2L -> DFN1006 2-pin ---
    if re.match(r'DFN-?2L$', sup.replace(' ', ''), re.IGNORECASE):
        return "DFN1006_2"

    # --- Fallback: try classifying using Package / Case as the primary string ---
    if case and case != sup:
        return _classify_digikey_package(case, "")

    return None


def _classify_ti_package(ti_pkg_token, pin_val, ti_gpn=None):
    """
    Returns a canonical code for a TI package name + pin count, matching
    the same code space as _classify_digikey_package.
    """
    norm = normalize_package(ti_pkg_token)
    norm_u = norm.upper()

    if norm_u.startswith("SOD523"):
        return "SOD523_2"
    if norm_u.startswith("SOD323"):
        return "SOD323_2"
    if norm_u.startswith("DFN0603"):
        if pin_val == 6:
            return "X2SON_6"
        if pin_val == 4:
            return "X2SON_4"
        return f"DFN0603_{pin_val}" if pin_val else "DFN0603_2"
    if norm_u.startswith("X2SON"):
        return f"X2SON_{pin_val}" if pin_val else None
    if norm_u.startswith("DFN1006"):
        return f"DFN1006_{pin_val}" if pin_val else None
    if norm_u.startswith("DFN1616"):
        return "DFN1616_6"
    if norm_u.startswith("DFN1110"):
        return "DFN1110_3"
    if norm_u.startswith("DFN2020"):
        return f"DFN2020_{pin_val}" if pin_val else None
    if norm_u.startswith("DFN2510"):
        return "DFN2510_10"
    if "SOT886" in norm_u or "SOT-886" in ti_pkg_token.upper():
        return "USON_6"   # TI's SOT-886 and its 6-pin USON are the same package
    if norm_u.startswith("USON"):
        if pin_val == 6 and ti_gpn and \
                str(ti_gpn).upper().replace("-Q1", "").strip() in TI_USON1616_PARTS:
            return "USON1616_6"
        return f"USON_{pin_val}" if pin_val else None
    if norm_u.startswith("DFN3030"):
        # 8-pin is DRBR, 6-pin is DRSR - the pin count decides the part.
        return f"DFN3030_{pin_val}" if pin_val else None
    if norm_u.startswith("DSBGA"):
        return "DSBGA_4"
    if norm_u.startswith("SC703"):
        return "SC703_3"
    if norm_u.startswith("SC706"):
        return "SC706_6"
    if norm_u.startswith("SOT233"):
        return "SOT233_3"
    if norm_u.startswith("SOT234"):
        return "SOT234_4"
    if norm_u.startswith("SOT235"):
        return "SOT235_5"
    if norm_u.startswith("SOT236"):
        return "SOT236_6"
    if re.match(r"SOT5[X\d]3", norm_u):
        return f"SOT5X3_{pin_val}" if pin_val else None
    if re.match(r"SOT9[X\d]3", norm_u):
        if pin_val:
            return f"SOT9X3_{pin_val}"
        _p9x3_ti = _sot9x3_pin_count(norm_u)
        return f"SOT9X3_{_p9x3_ti}" if _p9x3_ti else "SOT9X3_3"
    if norm_u.startswith("UQFN"):
        if pin_val == 14: return "UQFN_14"
        if pin_val == 10: return "UQFN_10"
        return None
    if norm_u.startswith("WQFN"):
        return "WQFN_12"
    if norm_u.startswith("WSON"):
        # 15-pin is the only WSON TI still carries; the rest were renamed.
        return "WSON_15" if pin_val == 15 else None

    return None


# Canonical code -> TI OPN suffix
CANONICAL_SUFFIX_MAP = {
    "SOD523_2": "DYAR",
    "SOD323_2": "DYFR",
    "DFN0603_2": "DPLR",
    "X2SON_3": "DMYR",
    "DFN0603_3": "DMYR",
    "X2SON_4": "DPWR",
    "DFN0603_4": "DPWR",
    "X2SON_6": "DPFR",
    "DFN0603_6": "DPFR",
    "DFN1006_2": "DPYR",
    "DFN1006_3": "DMXR",   # ESD122 exception
    "DFN1616_6": "VEBR",
    "DFN1110_3": "DXAR",
    "DFN2020_6": "DRVR",
    "DFN2020_3": "DRVR",
    "DFN2510_10": "DQAR",
    "DFN2510_6":  "DRYR",
    "USON_6": "DRYR",   # TI's SOT-886 classifies here too
    # TPD4E001's USON is the 1.6x1.6 variant - only it may cross to a
    # competitor DFN1616-6, and it uses DPKR rather than DRYR.
    "USON1616_6": "DPKR",
    "DFN3030_8": "DRBR",
    "DFN3030_6": "DRSR",
    "DSBGA_4": "YZFR",
    "SC703_3": "DCKR",
    "SC706_6": "DCKR",
    "SOT233_3": "DBZR",
    "SOT234_4": "DZDR",
    "SOT235_5": "DBVR",
    "SOT236_6": "DBVR",
    "SOT5X3_2": "DRLR",
    "SOT5X3_3": "DRLR",
    "SOT5X3_5": "DRLR",
    "SOT5X3_6": "DRLR",
    "SOT9X3_3": "DRTR",
    "UQFN_14": "RVZR",
    "UQFN_10": "RSER",
    "WQFN_12": "RSFR",
    "WSON_15": "DSMR",
}

def load_and_cache(cache_filename, source_path, force_reload, loader_func):
    if not os.path.exists(source_path):
        return pd.DataFrame()
    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR)
    cache_path = os.path.join(CACHE_DIR, cache_filename)
    use_cache = False
    if os.path.exists(cache_path) and not force_reload:
        try:
            if os.path.getmtime(source_path) < os.path.getmtime(cache_path):
                use_cache = True
        except FileNotFoundError:
            use_cache = False
    if use_cache:
        return pd.read_pickle(cache_path)
    else:
        print(f"Processing and caching '{os.path.basename(source_path)}'...")
        df = loader_func()
        if df is not None and not df.empty:
            df.to_pickle(cache_path)
        return df

def load_digi_export(xlsx_path, csv_path, cache_filename, force_reload, label):
    """
    Loads a per-competitor DigiKey export, saved as either a workbook or a CSV.
    These carry the channel counts (and, through them, the direction) that the
    competitors' own catalogues leave out.
    """
    if os.path.exists(xlsx_path):
        return load_and_cache(cache_filename, xlsx_path, force_reload,
            lambda: load_excel_data(xlsx_path, header_row=0))
    if os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path, on_bad_lines='skip', encoding='utf-8-sig', low_memory=False)
            print(f"Loaded {len(df)} rows from '{csv_path}'.")
            return df
        except Exception as e:
            print(f"Warning: Could not load '{csv_path}'. Error: {e}")
            return pd.DataFrame()
    print(f"Warning: '{xlsx_path}' not found. {label} parts will be treated as single channel.")
    return pd.DataFrame()


PACKAGE_DISPLAY_MAP = {
    "SOT236": "SOT-23-6",
    "SOT235": "SOT-23-5",
    "SOT233": "SOT-23-3",
    "SC76": "SC-76",
    "SC79": "SC-79",
}

TI_SPECS_DATABASE_FILE = "ti_specs.xlsx"
TI_ZENER_SPECS_DATABASE_FILE = "ti_zener_specs.xlsx"
AOS_SPECS_DATABASE_FILE = "aos_specs.xlsx"
AMAZING_SPECS_DATABASE_FILE = "amazing_specs.xlsx"
DIODES_DL_DATABASE_FILE = "diodes_dl_specs.xlsx"
DIODES_PL_DATABASE_FILE = "diodes_pl_specs.xlsx"
DIODES_ZENER_DATABASE_FILE = "diodes_zener_specs.xlsx"
# Nexperia ships a catalogue per part type, plus an automotive catalogue for
# each. Everything in an automotive file is automotive grade; in the standard
# files, an automotive part is marked by a "-Q" suffix on the part number.
NEXPERIA_ZENER_DATABASE_FILE = "nexperia_specs_zener.xls"
NEXPERIA_ESD_DATABASE_FILE = "nexperia_specs_esd.xls"
NEXPERIA_TVS_DATABASE_FILE = "nexperia_specs_tvs.xls"
NEXPERIA_AUTO_ZENER_DATABASE_FILE = "nexperia_specs_auto_zener.xls"
NEXPERIA_AUTO_ESD_DATABASE_FILE = "nexperia_specs_auto_esd.xls"
NEXPERIA_AUTO_TVS_DATABASE_FILE = "nexperia_specs_auto_tvs.xls"
LITTELFUSE_TVS_ARRAY_FILE = "littelfuse_specs_tvs_array.xlsx"
LITTELFUSE_AUTO_TVS_FILE = "littelfuse_specs_auto_tvs.xlsx"
LITTELFUSE_TVS_FILE = "littelfuse_specs_tvs.xlsx"
LITTELFUSE_DIGI_FILE = "littelfuse_digi_specs.xlsx"
LITTELFUSE_DIGI_CSV = "littelfuse_digi_specs.csv"
JIANGSU_ESD_FILE = "jiangsu_specs_esd.xls"
JIANGSU_TVS_FILE = "jiangsu_specs_tvs.xls"
JIANGSU_ZENER_FILE = "jiangsu_specs_zener.xls"
GALAXY_ESD_FILE = "galaxy_esd_specs.xlsx"
GALAXY_TVS_FILE = "galaxy_tvs_specs.xlsx"
GALAXY_ZENER_FILE = "galaxy_zener_specs.xlsx"
VISHAY_ESD_FILE = "vishay_esd_specs.xlsx"
VISHAY_ZENER_FILE = "vishay_zener_specs.xlsx"
SEMTECH_SPECS_FILE = "semtech_specs.xlsx"
STM_SPECS_CSV_FILE = "stmicroelectronics_specs.csv"
STM_SPECS_DATABASE_FILE = "stmicroelectronics_specs.xlsx"
PANJIT_ESD_FILE = "panjit_esd_specs.xls"
PANJIT_TVS_FILE = "panjit_tvs_specs.xls"
PANJIT_ZENER_FILE = "panjit_zener_specs.xls"
ONSEMI_ESD_FILE = "onsemi_esd_specs.csv"
ONSEMI_ZENER_FILE = "onsemi_zener_specs.csv"
ANBON_ESD_FILE   = "anbon_esd_specs.xlsx"
ANBON_ZENER_FILE = "anbon_zener_specs.xlsx"
ANBON_TVS_FILE   = "anbon_tvs_specs.xlsx"
ANBON_ESD_FILE   = "anbon_esd_specs.xlsx"
ANBON_ZENER_FILE = "anbon_zener_specs.xlsx"
ANBON_TVS_FILE   = "anbon_tvs_specs.xlsx"
CENTRAL_ZENER_FILE = "central_zener_specs.xlsx"
COMCHIP_ESD_FILE = "comchip_esd_specs.xlsx"
COMCHIP_TVS_FILE = "comchip_tvs_specs.xlsx"
COMCHIP_ZENER_FILE = "comchip_zener_specs.xlsx"
COMCHIP_DIGI_CSV = "comchip_digi_specs.csv"
DIOTEC_ESD_FILE = "diotec_esd_specs.xlsx"
DIOTEC_TVS_FILE = "diotec_tvs_specs.xlsx"
DIOTEC_ZENER_FILE = "diotec_zener_specs.xlsx"
EATON_SPECS_CSV_FILE = "eaton_specs.csv"
EATON_SPECS_FILE = "eaton_specs.xlsx"
EIC_TVS_FILE = "eic_tvs_specs.xlsx"
EIC_ZENER_FILE = "eic_zener_specs.xlsx"
EVVO_TVS_FILE = "evvo_tvs_specs.xlsx"
EVVO_ZENER_FILE = "evvo_zener_specs.xlsx"
EVVO_DIGI_FILE = "evvo_digi_specs.xlsx"
EVVO_DIGI_CSV = "evvo_digi_specs.csv"
GOODARK_ESD_FILE = "goodark_esd_specs.xlsx"
GOODARK_TVS_FILE = "goodark_tvs_specs.xlsx"
GOODARK_ZENER_FILE = "goodark_zener_specs.xlsx"
GOODARK_DIGI_FILE = "goodark_digi_specs.xlsx"
GOODARK_DIGI_CSV = "goodark_digi_specs.csv"
AMAZING_DIGI_FILE = "amazing_digi_specs.xlsx"
AMAZING_DIGI_CSV = "amazing_digi_specs.csv"
INFINEON_SPECS_FILE = "infineon_specs.xlsx"
YANGJIE_ESD_FILE       = "yangjie_esd_specs.xls"        
YANGJIE_TVS_FILE       = "yangjie_tvs_specs.xls"
YANGJIE_ZENER_FILE     = "yangjie_zener_specs.xls"
YANGJIE_AUTO_ESD_FILE  = "yangjie_auto_esd_specs.xls"
YANGJIE_AUTO_TVS_FILE  = "yangjie_auto_tvs_specs.xls"
YANGJIE_AUTO_ZENER_FILE = "yangjie_auto_zener_specs.xls"
JJM_ZENER_FILE       = "jjm_zener_specs.xlsx"
JJM_TVS_FILE         = "jjm_tvs_specs.xlsx"
JJM_ESD_FILE         = "jjm_esd_specs.xlsx"
JJM_AUTO_ZENER_FILE  = "jjm_auto_zener_specs.xlsx"
JJM_AUTO_TVS_FILE    = "jjm_auto_tvs_specs.xlsx"
JJM_AUTO_ESD_FILE    = "jjm_auto_esd_specs.xlsx"
MCC_ESD_FILE           = "mcc_esd_specs.xlsx"
MCC_TVS_FILE           = "mcc_tvs_specs.xlsx"
MCC_ZENER_FILE         = "mcc_zener_specs.xlsx"
LESHAN_ESD_FILE = "leshan_esd_specs.xlsx"
LESHAN_TVS_FILE = "leshan_tvs_specs.xlsx"
LESHAN_ZENER_FILE = "leshan_zener_specs.xlsx"
YENYO_ESD_FILE = "yenyo_esd_specs.xlsx"
YENYO_TVS_FILE = "yenyo_tvs_specs.xlsx"
YENYO_ZENER_FILE = "yenyo_zener_specs.xlsx"
NICHTEK_ESD_FILE = "nichtek_esd_specs.xlsx"
NICHTEK_TVS_FILE = "nichtek_tvs_specs.xlsx"
NICHTEK_ZENER_FILE = "nichtek_zener_specs.xlsx"
INPAQ_ESD_FILE = "inpaq_esd_specs.xlsx"
INPAQ_AUTO_ESD_FILE = "inpaq_auto_esd_specs.xlsx"
INPAQ_TVS_FILE = "inpaq_tvs_specs.xlsx"
INPAQ_AUTO_TVS_FILE = "inpaq_auto_tvs_specs.xlsx"

def is_similar_voltage(v1_str, v2_str, tolerance_percent=20, two_sided=True):
    """
    v1 is the competitor's figure, v2 the TI part's. The two-sided form asks
    that the gap be within tolerance of BOTH, which is materially tighter than
    the percentage implies - 7 V vs 5.5 V is 21% off the competitor but 27%
    off TI. Pass two_sided=False to measure against the competitor alone.
    """
    v1 = to_numeric_val(v1_str, default_if_error=None)
    v2 = to_numeric_val(v2_str, default_if_error=None)
    if v1 is None or v2 is None:
        return False
    if v1 == 0 and v2 == 0: return True
    if v1 == 0 or v2 == 0: return False
    if abs(v1 - v2) / v1 * 100 > tolerance_percent:
        return False
    return not two_sided or abs(v1 - v2) / v2 * 100 <= tolerance_percent

def _find_part_number_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        for keyword in keywords:
            if keyword.lower() in str(col).lower():
                return col
    return None

def _find_exact_match(norm_input, df, part_col, raw_input=None):
    """
    Normalizing away punctuation collides real parts: "PDZ3.6B-Q" and
    "PDZ36B-Q" both reduce to "PDZ36BQ". So the part number as typed is
    checked first, and the normalized comparison is only the fallback.
    """
    series = df[part_col].astype(str).str.strip().str.upper()

    if raw_input:
        raw_mask = (series == str(raw_input).strip().upper())
        if raw_mask.any():
            return df.loc[raw_mask.idxmax()][part_col]

    df_normalized = series.str.replace(r'[\W_]+', '', regex=True)
    match_mask = (df_normalized == norm_input)
    if match_mask.any():
        return df.loc[match_mask.idxmax()][part_col]
    return None

def _find_partial_match(norm_input, df, part_col):
    df_normalized = df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    match_mask = df_normalized.str.contains(norm_input, na=False)
    if match_mask.any():
        return df.loc[match_mask.idxmax()][part_col]
    return None

def _find_reverse_partial_match(norm_input, df, part_col):
    """
    The input carries a suffix the database does not - "DT1452-02SOQ-7" against
    a sheet that lists "DT1452-02SOQ" - so a database part that is a PREFIX of
    the input is a hit. More than one can be: "DT145202SO" and "DT145202SOQ"
    are both prefixes of "DT145202SOQ7". The longest is the right answer,
    because it accounts for the most of what was actually typed. Taking the
    first in sheet order instead silently drops the Q (automotive) variant
    whenever the base part happens to be listed above it.
    """
    df_normalized = df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    non_empty = df_normalized[(df_normalized != '') & (df_normalized.notna())]
    if non_empty.empty:
        return None
    prefix_matches = non_empty[non_empty.apply(lambda db_part: norm_input.startswith(db_part))]
    if prefix_matches.empty:
        return None
    return df.loc[prefix_matches.str.len().idxmax()][part_col]

def _find_fuzzy_match(norm_input, df, part_col, min_len):
    if len(norm_input) <= min_len:
        return None
    df_normalized = df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    non_empty_mask = (df_normalized != '') & (df_normalized.notna())
    for i in range(len(norm_input) - 1, min_len - 1, -1):
        sliced_input = norm_input[:i]
        fuzzy_match_mask = df_normalized[non_empty_mask].str.startswith(sliced_input, na=False)
        if fuzzy_match_mask.any():
            matched_part = df.loc[fuzzy_match_mask.idxmax()][part_col]
            print(f"--> Fuzzy match (min_len={min_len}) found. Using closest match: '{matched_part}'")
            return matched_part
    return None


PART_COL = "Mfr Part #"
SUPPLIER_COL = "Supplier"


def _normalize_digikey_package(raw_pkg):
    """
    Handles DigiKey package strings like 'SC-79', 'SOD-523', '0402 (1005 Metric)',
    and comma-separated lists as seen in TI parts. Strips imperial/metric size
    suffixes and passes the result through the standard normalize_package.
    """
    if not raw_pkg or raw_pkg == "-":
        return raw_pkg

    # If there are multiple packages separated by commas, normalize each one.
    parts = [p.strip() for p in str(raw_pkg).split(',') if p.strip()]
    normalized = []
    for p in parts:
        # Strip trailing imperial/metric size annotations like '0402 (1005 Metric)'
        # keeping only the first token when it looks like a passive size code.
        p_clean = re.sub(r'\s*\(.*?\)', '', p).strip()
        normalized.append(normalize_package(p_clean))

    return ', '.join(normalized) if normalized else raw_pkg


def _direction_from_digikey_row(row, df):
    """
    Reads directionality out of a DigiKey row. Returns "Unidirectional",
    "Bidirectional", or None when the row does not say.

    Primary signal is the pair of channel-count columns - DigiKey populates only
    one of them. Falls back to any column that spells the direction out, which is
    how the zener CSV carries it (in Description rather than its own column).
    """
    def cell(col):
        val = str(row.get(col, "") if col in row.index else "").strip()
        if val.lower() in ("", "nan", "none", "-"):
            return ""
        # A channel column read as a float comes back "0.0", not "0", so treat
        # any numeric zero as empty rather than as a populated channel count.
        try:
            if float(val) == 0:
                return ""
        except (TypeError, ValueError):
            pass
        return val

    # "unidirectional" does not contain "bidirectional", so these can't collide.
    uni_col = next((c for c in df.columns if "unidirectional channels" in c.lower()), None)
    bi_col = next((c for c in df.columns if "bidirectional channels" in c.lower()), None)

    if bi_col and cell(bi_col):
        return "Bidirectional"
    if uni_col and cell(uni_col):
        return "Unidirectional"

    # Text fallback: DigiKey descriptions write it as "BI-DIR" / "UNI-DIR" as
    # often as they spell it out, so compare with separators stripped.
    for col in df.columns:
        v = re.sub(r'[\s\-_]+', '', cell(col).lower())
        if not v:
            continue
        if "bidir" in v:
            return "Bidirectional"
        if "unidir" in v:
            return "Unidirectional"
    return None


def _resolve_digikey_column(df, preferred, keywords):
    """
    Finds a column by name, tolerating header drift between CSV exports (a BOM
    on the first header, stray whitespace, a renamed variant). `preferred` is
    tried verbatim first, then a keyword search, then a comparison with all
    non-alphanumerics stripped. Returns None if nothing matches.
    """
    if df is None or df.empty:
        return None
    if preferred in df.columns:
        return preferred

    col = _find_part_number_column(df, keywords)
    if col is not None:
        return col

    def flat(s):
        return re.sub(r'[^a-z0-9]+', '', str(s).lower())

    targets = {flat(preferred)} | {flat(k) for k in keywords}
    for c in df.columns:
        if flat(c) in targets:
            return c
    return None


def fetch_digikey_direction(part_input, supplier_input, digikey_df, digikey_zener_df=None,
                            primary_label="digikey_specs"):
    """
    Looks a part up in the DigiKey CSVs purely to recover its direction, for
    competitor databases that carry no direction column of their own.
    digikey_specs is checked first (it has the uni/bi channel columns), the
    zener CSV second.

    Only strict matches are used - a fuzzy hit on the wrong part would stamp a
    wrong direction onto the cross, which is worse than leaving it unknown.
    """
    norm_input = re.sub(r'[\W_]+', '', str(part_input).upper())
    if not norm_input:
        return None
    norm_supplier = str(supplier_input or "").strip().lower()

    part_keywords = ["mfr part", "manufacturer part", "part number", "part #", "part#"]
    supplier_keywords = ["supplier", "manufacturer"]

    sources = [
        (digikey_df, PART_COL, SUPPLIER_COL, primary_label),
        (digikey_zener_df, "DigiKey Part #", "Manufacturer", "digikey_zener"),
    ]

    for df, preferred_part_col, preferred_supplier_col, label in sources:
        if df is None or df.empty:
            print(f"Direction lookup - {label} is not loaded, skipping.")
            continue

        part_col = _resolve_digikey_column(df, preferred_part_col, part_keywords)
        if part_col is None:
            print(f"Direction lookup - no part number column in {label}; headers are: {list(df.columns)}")
            continue
        if part_col != preferred_part_col:
            print(f"Direction lookup - using column '{part_col}' in {label} (expected '{preferred_part_col}').")

        supplier_col = _resolve_digikey_column(df, preferred_supplier_col, supplier_keywords)

        search_df = df
        if supplier_col and supplier_col in df.columns and norm_supplier:
            supplier_df = df[df[supplier_col].astype(str).str.lower().str.contains(norm_supplier, na=False)]
            if not supplier_df.empty:
                search_df = supplier_df

        match = (
            _find_exact_match(norm_input, df, part_col, raw_input=part_input) or
            _find_partial_match(norm_input, search_df, part_col) or
            _find_reverse_partial_match(norm_input, search_df, part_col)
        )
        if not match:
            # Last resort: the same base part in another package/reel suffix.
            # min_len=6 keeps the base part number intact, so the die - and with
            # it the direction - is the same one. Looser fuzzy matching is not
            # used here: a wrong direction hard-filters TI candidates.
            match = _find_fuzzy_match(norm_input, search_df, part_col, min_len=6)
            if match:
                print(f"Direction lookup - no exact hit in {label}, using closest part '{match}'.")
        if not match:
            print(f"Direction lookup - '{part_input}' not found in {label} ({len(search_df)} rows searched).")
            continue

        row = search_df[search_df[part_col].astype(str) == str(match)].iloc[0]
        direction = _direction_from_digikey_row(row, search_df)
        if direction:
            print(f"Direction from {label}: '{match}' is {direction}.")
            return direction
        print(f"Direction lookup - found '{match}' in {label} but the row states no direction.")

    return None


def fetch_comchip_channels(comchip_digi_df, *part_candidates):
    """
    Looks a Comchip part up in the DigiKey export by 'Mfr Part #' and returns its
    'Channels' value. Parts that are not listed there are single channel.
    """
    if comchip_digi_df is None or comchip_digi_df.empty:
        return "1"

    part_col = None
    chan_col = None
    for col in comchip_digi_df.columns:
        col_lower = str(col).strip().lower()
        if part_col is None and "mfr part" in col_lower:
            part_col = col
        if chan_col is None and "channel" in col_lower:
            chan_col = col
    if part_col is None or chan_col is None:
        return "1"

    normalized_col = comchip_digi_df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    for candidate in part_candidates:
        if not candidate:
            continue
        normalized_candidate = re.sub(r'[\W_]+', '', str(candidate).upper())
        if not normalized_candidate:
            continue
        match_rows = comchip_digi_df[normalized_col == normalized_candidate]
        if match_rows.empty:
            continue
        chan_digits = re.search(r'\d+', str(safe_strip(match_rows.iloc[0][chan_col])))
        if chan_digits:
            print(f"Channels for '{candidate}' from '{COMCHIP_DIGI_CSV}': {chan_digits.group(0)}.")
            return chan_digits.group(0)
        return "1"
    return "1"

def fetch_digi_vrwm(digi_df, label, *part_candidates):
    """
    Looks a part up in a per-competitor DigiKey export by 'Mfr Part #' and
    returns its standoff voltage, formatted the way the scrapers format theirs
    so the comparison table and the Vrwm tolerance check both read it. Simpler
    than fetch_digi_channels because there is one column, not a directional
    pair. Returns None when the export has nothing for the part, leaving the
    caller's own value in place.
    """
    if digi_df is None or digi_df.empty:
        return None

    part_col = None
    vrwm_col = None
    for col in digi_df.columns:
        col_lower = str(col).strip().lower()
        if part_col is None and "mfr part" in col_lower:
            part_col = col
        if vrwm_col is None and "reverse standoff" in col_lower:
            vrwm_col = col
    if part_col is None or vrwm_col is None:
        return None

    normalized_col = digi_df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    for candidate in part_candidates:
        if not candidate:
            continue
        normalized_candidate = re.sub(r'[\W_]+', '', str(candidate).upper())
        if not normalized_candidate:
            continue
        match_rows = digi_df[normalized_col == normalized_candidate]
        if match_rows.empty:
            continue
        for _, row in match_rows.iterrows():
            # "5V", "5.5 V" and a bare "5" all reduce to the same number.
            cell = str(safe_strip(row[vrwm_col]))
            # A dual-voltage part lists both rails -- "15V (Max), 24V (Max)".
            # The higher one is the part's standoff, so every figure attached to
            # a V is collected and the largest wins rather than the first.
            # Matching on the V keeps "(Max)" and any trailing conditions out of
            # it; a bare number with no unit is the fallback.
            volts = re.findall(r'(\d+(?:\.\d+)?)\s*V', cell, flags=re.IGNORECASE)
            if not volts:
                bare = re.search(r'\d+(?:\.\d+)?', cell)
                volts = [bare.group(0)] if bare else []
            if volts:
                value = f"{max(float(v) for v in volts):g} V"
                print(f"Vrwm for '{candidate}' from '{label}': {value}.")
                return value
        return None
    return None

def fetch_digi_direction(digi_df, label, *part_candidates):
    """
    Infers a part's polarity from which channel column the DigiKey export fills
    in: a count under 'Bidirectional Channels' means bidirectional, one under
    'Unidirectional Channels' means unidirectional. Returns None when neither is
    populated, or when both are -- that is ambiguous, not a reading.
    """
    if digi_df is None or digi_df.empty:
        return None

    part_col = None
    bi_col = None
    uni_col = None
    for col in digi_df.columns:
        col_lower = str(col).strip().lower()
        flat = col_lower.replace("-", "").replace(" ", "")
        if part_col is None and "mfr part" in col_lower:
            part_col = col
        if uni_col is None and flat.startswith("unidirectional"):
            uni_col = col
        elif bi_col is None and flat.startswith("bidirectional"):
            bi_col = col
    if part_col is None or (bi_col is None and uni_col is None):
        return None

    def _populated(row, col):
        # A blank, a dash and an explicit 0 all mean "not this polarity".
        if col is None:
            return False
        digits = re.search(r'\d+', str(safe_strip(row[col])))
        return bool(digits) and int(digits.group(0)) > 0

    normalized_col = digi_df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    for candidate in part_candidates:
        if not candidate:
            continue
        normalized_candidate = re.sub(r'[\W_]+', '', str(candidate).upper())
        if not normalized_candidate:
            continue
        match_rows = digi_df[normalized_col == normalized_candidate]
        if match_rows.empty:
            continue
        for _, row in match_rows.iterrows():
            bi = _populated(row, bi_col)
            uni = _populated(row, uni_col)
            if bi and not uni:
                print(f"Direction for '{candidate}' from '{label}': Bidirectional.")
                return "Bidirectional"
            if uni and not bi:
                print(f"Direction for '{candidate}' from '{label}': Unidirectional.")
                return "Unidirectional"
        return None
    return None

def fetch_digi_channels(digi_df, direction, label, *part_candidates):
    """
    Looks a part up in a per-competitor DigiKey export by 'Mfr Part #' and
    returns its channel count. The export splits the count across 'Bidirectional
    Channels' and 'Unidirectional Channels', so the part's own direction picks
    the column and the other one is the fallback. Returns None when the export
    has nothing for the part, leaving the caller's own value in place.
    """
    if digi_df is None or digi_df.empty:
        return None

    part_col = None
    bi_col = None
    uni_col = None
    for col in digi_df.columns:
        col_lower = str(col).strip().lower()
        flat = col_lower.replace("-", "").replace(" ", "")
        if part_col is None and "mfr part" in col_lower:
            part_col = col
        if uni_col is None and flat.startswith("unidirectional"):
            uni_col = col
        elif bi_col is None and flat.startswith("bidirectional"):
            bi_col = col
    if part_col is None or (bi_col is None and uni_col is None):
        return None

    ordered_cols = [c for c in ([uni_col, bi_col] if "uni" in str(direction).lower()
                                else [bi_col, uni_col]) if c is not None]

    normalized_col = digi_df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)
    for candidate in part_candidates:
        if not candidate:
            continue
        normalized_candidate = re.sub(r'[\W_]+', '', str(candidate).upper())
        if not normalized_candidate:
            continue
        match_rows = digi_df[normalized_col == normalized_candidate]
        if match_rows.empty:
            continue
        for col in ordered_cols:
            for _, row in match_rows.iterrows():
                chan_digits = re.search(r'\d+', str(safe_strip(row[col])))
                if chan_digits:
                    print(f"Channels for '{candidate}' from '{label}': {chan_digits.group(0)}.")
                    return chan_digits.group(0)
        return None
    return None

# Display spellings for the output Excel. The lookup logic elsewhere matches on
# the raw lowercase string, so this is applied only when writing, never before
# get_competitor_specs_leniently. Aliases mirror the substring checks in that
# function so anything the lookup accepts also normalizes here.
_COMPETITOR_NAME_ALIASES = {
    "Amazing Microelectronic Corp": ["amazing", "az"],
    "Anbon Semiconductor": ["anbon", "anb"],
    "Alpha & Omega Semiconductor Ltd": ["aos", "alpha", "alpha omega", "alpha & omega", "alpha & omega semiconductor", "alpha omega semiconductor"],
    "Central Semiconductor Corp": ["central", "centralsemi", "csi"],
    "Comchip Technology Co Ltd": ["comchip", "comchiptech"],
    "Diodes Inc": ["diodes", "diodes inc"],
    "Diotec Semiconductor Ag": ["diotec"],
    "Eaton Corp": ["eaton"],
    "Eic Semiconductor Co Ltd": ["eic", "eicsemi"],
    "EVER Semiconductor Co Ltd (EVVOSEMI)": ["evvo", "ever", "evvosemi", "ever semiconductor"],
    "Changzhou Galaxy Century Microelectronics Co Ltd": ["galaxy", "gme"],
    "Good-Ark Electronics Co Ltd": ["goodark", "good-ark", "good ark", "gsc"],
    "Infineon Technologies Ag": ["infineon"],
    "INPAQ Technology Co Ltd": ["inpaq", "inpaq technology"],
    "Jiangsu Changjiang Electronics Technology Co Ltd": ["jiangsu", "jsu"],
    "Jiangsu Jiejie Microelectronics Co Ltd": ["jjm", "jiejie", "jie jie"],
    "Leshan Radio Co Ltd": ["leshan", "lrc"],
    "Littelfuse Inc": ["littelfuse", "lf"],
    "Micro Commercial Components Corp": ["mcc"],
    "Nexperia": ["nexperia"],
    "NichTek": ["nichtek"],
    "Onsemi": ["onsemi", "on semi"],
    "Panjit International Inc": ["panjit"],
    "Semtech Corp": ["semtech"],
    "Stmicroelectronics": ["stm"],
    "Vishay Semiconductors": ["vishay"],
    "Yangzhou Yangjie Electronics Co Ltd": ["yangjie", "yj"],
    "Yenyo Technology": ["yenyo"],
}

COMPETITOR_CANONICAL_NAMES = {
    alias: canonical
    for canonical, aliases in _COMPETITOR_NAME_ALIASES.items()
    for alias in aliases
}

def canonical_competitor_name(raw_name):
    """
    Whatever was typed -> the display spelling. Exact alias match first, then a
    longest-first word-boundary search, so "LF Inc" still reaches Littelfuse
    without "on semi" claiming "anbon semiconductor". An unrecognized name is
    title-cased rather than dropped, so a new competitor still reads sensibly.
    """
    key = str(raw_name).strip().lower()
    if not key or key == "nan":
        return str(raw_name).strip()
    if key in COMPETITOR_CANONICAL_NAMES:
        return COMPETITOR_CANONICAL_NAMES[key]
    for alias in sorted(COMPETITOR_CANONICAL_NAMES, key=len, reverse=True):
        if re.search(r"(?<!\w)" + re.escape(alias) + r"(?!\w)", key):
            return COMPETITOR_CANONICAL_NAMES[alias]
    return str(raw_name).strip().title()

def get_competitor_specs_leniently(part_input, competitor_name, all_dfs):
    """
    Central dispatcher. Checks competitor-specific local databases first,
    falls back to DigiKey CSV if not found.
    """

    # Competitor sheets that weld the datasheet hyperlink's text onto the part
    # number (JJM's do) send people to the prompt with "JEB03CX-AU<spaces>DataSheet"
    # on the clipboard. Normalization strips the whitespace but not the word, so
    # the input reads as JEB03CXAUDATASHEET -- exact and partial matching miss,
    # and reverse-partial then resolves to the shorter base part in the
    # commercial book. Stripped here so every mode and every competitor
    # benefits, before norm_input is built.
    part_input = re.sub(r"\s{2,}data\s*sheet\s*$", "", str(part_input), flags=re.IGNORECASE).strip()
    print(f"\nPerforming lenient search for '{part_input}' from competitor '{competitor_name}'...")

    aos_specs_df, amazing_specs_df, diodes_df_map, nexperia_df_map, jjm_df_map, \
    littelfuse_df_map, littelfuse_digi_df, jiangsu_df_map, vishay_df_map, yangjie_df_map, mcc_df_map, galaxy_df_map, leshan_df_map, nichtek_df_map, yenyo_df_map, inpaq_df_map, semtech_specs_df, \
    stm_specs_df, panjit_df_map, onsemi_df_map, \
    anbon_esd_df, anbon_zener_df, anbon_tvs_df, central_zener_df, \
    comchip_esd_df, comchip_tvs_df, comchip_zener_df, comchip_digi_df, \
    diotec_df_map, eaton_specs_df, eic_df_map, evvo_df_map, evvo_digi_df, \
    goodark_df_map, goodark_digi_df, amazing_digi_df, infineon_specs_df, \
    ti_specs_df, ti_zener_specs_df = all_dfs

    norm_input = re.sub(r'[\W_]+', '', str(part_input).upper())
    if not norm_input:
        return None

    exact_part_name = None
    comp_specs = None

    is_comchip_part = any(k in competitor_name for k in ["comchip", "comchiptech"])
    if is_comchip_part:
        comp_specs = fetch_comchip_specs(part_input, comchip_esd_df, comchip_tvs_df, comchip_zener_df)
        if comp_specs:
            comp_specs["Channels"] = fetch_comchip_channels(
                comchip_digi_df, comp_specs.get("Device Name"), part_input)
            return comp_specs
        # Found in Comchip's sheets but rejected (axial lead). Stop here rather
        # than falling through to DigiKey, which would cross it anyway.
        if comchip_has_part(part_input, comchip_esd_df, comchip_tvs_df, comchip_zener_df):
            return None

    is_diotec_part = "diotec" in competitor_name
    if is_diotec_part:
        comp_specs = fetch_diotec_specs_from_excel(part_input, diotec_df_map)
        if comp_specs:
            return comp_specs

    is_eic_part = any(k in competitor_name for k in ["eic", "eicsemi"])
    if is_eic_part:
        comp_specs = fetch_eic_specs_from_excel(part_input, eic_df_map)
        if comp_specs:
            return comp_specs

    is_evvo_part = any(k in competitor_name for k in ["evvo", "evvosemi", "ever semiconductor"])
    if is_evvo_part:
        comp_specs = fetch_evvo_specs_from_excel(part_input, evvo_df_map)
        if comp_specs:
            if safe_strip(comp_specs.get("Direction")) == "-":
                digi_direction = fetch_digi_direction(
                    evvo_digi_df, EVVO_DIGI_FILE,
                    comp_specs.get("Device Name"), part_input)
                if digi_direction:
                    comp_specs["Direction"] = digi_direction
            # The export is authoritative where it has the part; otherwise the
            # description's own count stands, defaulting to a single channel.
            if "Channels" in comp_specs:
                comp_specs["Channels"] = fetch_digi_channels(
                    evvo_digi_df, comp_specs.get("Direction"), EVVO_DIGI_FILE,
                    comp_specs.get("Device Name"), part_input) or comp_specs.get("Channels", "1")
            # Standoff is parsed out of the description, and falls back to the
            # voltage coded into the part name. When neither states one the
            # export is the last resort -- a part with no Vrwm is dropped by the
            # tolerance filter before it can be scored. Only for TVS/ESD parts:
            # this key holds Vz on a zener, which is a different column.
            if (safe_strip(comp_specs.get("Voltage - Reverse Standoff (Typ)")) == "-"
                    and "zener" not in str(comp_specs.get("Source File", "")).lower()):
                digi_vrwm = fetch_digi_vrwm(
                    evvo_digi_df, EVVO_DIGI_FILE,
                    comp_specs.get("Device Name"), part_input)
                if digi_vrwm:
                    comp_specs["Voltage - Reverse Standoff (Typ)"] = digi_vrwm
            return comp_specs
        # Found in EVVO's sheets but refused (Schottky). Stop here rather than
        # falling through to DigiKey, which would cross it anyway.
        if evvo_has_part(part_input, evvo_df_map):
            return None

    is_goodark_part = any(k in competitor_name for k in ["goodark", "good-ark", "good ark", "gsc"])
    if is_goodark_part:
        comp_specs = fetch_goodark_specs_from_excel(part_input, goodark_df_map)
        if comp_specs:
            # Configuration says "Array" without a count, so the export fills it
            # in where it has the part; otherwise the sheet's own value stands.
            if "Channels" in comp_specs:
                comp_specs["Channels"] = fetch_digi_channels(
                    goodark_digi_df, comp_specs.get("Direction"), GOODARK_DIGI_FILE,
                    comp_specs.get("Device Name"), part_input) or comp_specs.get("Channels", "1")
            return comp_specs

    is_central_part = any(k in competitor_name for k in ["central", "centralsemi", "csi"])
    if is_central_part:
        comp_specs = fetch_central_specs(part_input, central_zener_df)
        if comp_specs:
            return comp_specs
        # Found in Central's sheet but rejected (non surface-mount). Stop here
        # rather than falling through to DigiKey, which would cross it anyway.
        if central_has_part(part_input, central_zener_df):
            return None

    is_galaxy_part = any(k in competitor_name for k in ["galaxy", "gme"])

    is_anbon_part = any(k in competitor_name for k in ["anbon", "anb"])
    if is_anbon_part:
        comp_specs = fetch_anbon_specs(part_input, anbon_esd_df, anbon_zener_df, anbon_tvs_df)
        if comp_specs:
            return comp_specs
        
    is_onsemi_part = any(k in competitor_name for k in ["onsemi", "on semi"])

    if is_onsemi_part:
        comp_specs = fetch_onsemi_specs_from_excel(
            part_input,
            onsemi_df_map
        )

        if comp_specs:
            return comp_specs

        return None

    is_diodes_part = any(k in competitor_name for k in ["diodes", "diodes inc"])
    is_nexperia_part = "nexperia" in competitor_name
    is_jjm_part = any(k in competitor_name for k in ["jjm", "jiejie", "jie jie"])
    is_littelfuse_part = any(k in competitor_name for k in ["lf", "littelfuse"])
    is_jiangsu_part = any(k in competitor_name for k in ["jiangsu", "jsu"])
    is_vishay_part = "vishay" in competitor_name
    is_panjit_part = "panjit" in competitor_name
    is_yangjie_part = any(k in competitor_name for k in ["yangjie", "yj"])   # ~line 1248
    is_mcc_part = "mcc" in competitor_name
    is_leshan_part = any(k in competitor_name for k in ["lrc", "leshan"])
    is_yenyo_part = "yenyo" in competitor_name
    is_nichtek_part = "nichtek" in competitor_name
    is_inpaq_part = "inpaq" in competitor_name

    if is_diodes_part or is_nexperia_part or is_jjm_part or is_nichtek_part or is_inpaq_part or is_littelfuse_part or is_jiangsu_part or is_yenyo_part or is_vishay_part or is_panjit_part or is_yangjie_part or is_leshan_part or is_mcc_part or is_galaxy_part:
        df_map = {}
        part_col_keywords = []
        fetch_func = None

        if is_diodes_part:
            df_map, part_col_keywords, fetch_func = (diodes_df_map, ["part number", "mfr part"], fetch_diodes_specs_from_excel)
        elif is_nexperia_part:
            df_map, part_col_keywords, fetch_func = (nexperia_df_map, ["type number", "part number", "mfr part"], fetch_nexperia_specs_from_excel)
        elif is_littelfuse_part:
            df_map, part_col_keywords, fetch_func = (littelfuse_df_map, ["part number", "mfr part"], fetch_littelfuse_specs_from_excel)
        elif is_jiangsu_part:
            df_map, part_col_keywords, fetch_func = (jiangsu_df_map, ["part number", "mfr part"], fetch_jiangsu_specs_from_excel)
        elif is_jjm_part:
            df_map, part_col_keywords, fetch_func = (jjm_df_map, ["product name", "part number", "device"], fetch_jjm_specs_from_excel)
        elif is_vishay_part:
            df_map, part_col_keywords, fetch_func = (vishay_df_map, ["part number", "mfr part"], fetch_vishay_specs_from_excel)
        elif is_yenyo_part:
            df_map, part_col_keywords, fetch_func = (yenyo_df_map, ["part number", "product", "device"], fetch_yenyo_specs_from_excel)
        elif is_panjit_part:
            df_map, part_col_keywords, fetch_func = (
                panjit_df_map,
                ["part number"],
                fetch_panjit_specs_from_excel,
            )
        elif is_yangjie_part:
            df_map, part_col_keywords, fetch_func = (yangjie_df_map, ["产品名称", "part number", "product"], fetch_yangjie_specs_from_excel)
        elif is_mcc_part:
            df_map, part_col_keywords, fetch_func = (mcc_df_map, ["product", "part number"], fetch_mcc_specs_from_excel)
        elif is_galaxy_part:
            df_map, part_col_keywords, fetch_func = (galaxy_df_map, ["product", "part number"], fetch_galaxy_specs_from_excel)
        elif is_leshan_part:
            df_map, part_col_keywords, fetch_func = (leshan_df_map, ["device", "part number"], fetch_leshan_specs_from_excel)
        elif is_nichtek_part:
            df_map, part_col_keywords, fetch_func = (nichtek_df_map, ["p/n", "part number"], fetch_nichtek_specs_from_excel)
        elif is_inpaq_part:
            df_map, part_col_keywords, fetch_func = (
                inpaq_df_map,
                ["part no", "part number"],
                fetch_inpaq_specs_from_excel,
            )

        non_fuzzy_methods = [
            lambda ni, df, pc: _find_exact_match(ni, df, pc, raw_input=part_input),
            _find_partial_match,
            _find_reverse_partial_match,
        ]
        fuzzy_methods = [
            lambda ni, df, pc: _find_fuzzy_match(ni, df, pc, min_len=6),
            lambda ni, df, pc: _find_fuzzy_match(ni, df, pc, min_len=4)
        ]

        for method in non_fuzzy_methods:
            for df_name, df in df_map.items():
                part_col = _find_part_number_column(df, part_col_keywords)
                if part_col:
                    exact_part_name = method(norm_input, df, part_col)
                    if exact_part_name:
                        comp_specs = fetch_func(exact_part_name, df_map)
                        break
            if comp_specs:
                break

        if not comp_specs:
            # Fall back to fuzzy
            for method in fuzzy_methods:
                for df_name, df in df_map.items():
                    part_col = _find_part_number_column(df, part_col_keywords)
                    if part_col:
                        exact_part_name = method(norm_input, df, part_col)
                        if exact_part_name:
                            comp_specs = fetch_func(exact_part_name, df_map)
                            break
                if comp_specs:
                    break

    else:
        single_file_checks = [
            ("aos", aos_specs_df, ["product", "part number", "mfr part"], fetch_aos_specs_from_excel),
            ("amazing", amazing_specs_df, ["part number", "mfr part"], fetch_amazing_specs_from_excel),
            ("semtech", semtech_specs_df, ["parts", "mfr part"], fetch_semtech_specs_from_excel),
            ("stm", stm_specs_df, ["manufacturer part number", "part number", "mfr part"], fetch_stm_specs_from_excel),
            ("eaton", eaton_specs_df, ["eaton part number", "part number", "mfr part"], fetch_eaton_specs_from_excel),
            ("infineon", infineon_specs_df, ["opn"], fetch_infineon_specs_from_excel)
        ]

        found_competitor = False
        for comp_key, df, keywords, fetch_func in single_file_checks:
            if comp_key in competitor_name:
                found_competitor = True
                part_col = _find_part_number_column(df, keywords)
                if part_col:
                    # Try exact/partial/reverse before fuzzy
                    exact_part_name = (
                        _find_exact_match(norm_input, df, part_col) or
                        _find_partial_match(norm_input, df, part_col) or
                        _find_reverse_partial_match(norm_input, df, part_col)
                    )
                    if exact_part_name:
                        comp_specs = fetch_func(exact_part_name, df)
                    else:
                        exact_part_name = (
                            _find_fuzzy_match(norm_input, df, part_col, min_len=6) or
                            _find_fuzzy_match(norm_input, df, part_col, min_len=4)
                        )
                        if exact_part_name:
                            comp_specs = fetch_func(exact_part_name, df)
                break

        # Eaton lists leaded (axial/radial) bodies alongside its surface-mount
        # catalogue. fetch_eaton_specs_from_excel returns None for those, and a
        # leaded part has no TI equivalent - so stop here rather than falling
        # through to DigiKey, which would cross it anyway.
        if "eaton" in competitor_name and comp_specs is None and exact_part_name:
            return None

        if not found_competitor:
            print(f"Competitor '{competitor_name}' is not in the local databases.")
            return None

    if comp_specs:
        print(f"Found match. The exact part name is '{exact_part_name}'. Fetching specs...")
        # AZ9xxx parts from Amazing are always automotive grade
        if str(exact_part_name).upper().startswith("AZ9"):
            comp_specs["Grade"] = "Automotive"

    
        # Amazing's Excel carries no directionality, so recover it from DigiKey.
        # The rest of the specs still come from the Excel - only Direction is
        # taken from DigiKey, and only when DigiKey actually states one.
        if "amazing" in competitor_name and str(comp_specs.get("Source File", "")).lower().find("digikey") < 0:
            dk_direction = fetch_digikey_direction(
                exact_part_name or part_input, competitor_name, amazing_digi_df, None,
                primary_label=AMAZING_DIGI_FILE
            )
            if dk_direction:
                prev_direction = str(comp_specs.get("Direction", "-")).strip()
                if prev_direction and prev_direction not in ("-", "nan", "none") \
                        and prev_direction.lower() != dk_direction.lower():
                    print(f"Direction - overriding Excel value '{prev_direction}' with DigiKey's '{dk_direction}'.")
                comp_specs["Direction"] = dk_direction
            if "Channels" in comp_specs:
                comp_specs["Channels"] = fetch_digi_channels(
                    amazing_digi_df, comp_specs.get("Direction"), AMAZING_DIGI_FILE,
                    exact_part_name, part_input) or comp_specs.get("Channels", "1")
    
        # Nexperia grade: anything out of an automotive catalogue is automotive
        # whatever its part number looks like. In the standard catalogues the
        # automotive parts are the ones carrying a "-Q" suffix.
        if is_nexperia_part:
            if "auto" in str(comp_specs.get("Source File", "")).lower():
                comp_specs["Grade"] = "Automotive"
            else:
                comp_specs["Grade"] = (
                    "Automotive" if is_nexperia_automotive(exact_part_name or part_input)
                    else "Commercial"
                )
        
        # Littelfuse's own sheets rarely carry a channel count, so the DigiKey
        # export backs it up where it has the part. Only the count is taken from
        # that export - every other spec still comes from Littelfuse's sheets.
        if is_littelfuse_part and "digikey" not in str(comp_specs.get("Source File", "")).lower():
            dk_channels = fetch_digi_channels(
                littelfuse_digi_df, comp_specs.get("Direction"), LITTELFUSE_DIGI_FILE,
                comp_specs.get("Device Name"), exact_part_name, part_input)
            if dk_channels:
                comp_specs["Channels"] = dk_channels
    else:
        print(f"Could not find '{part_input}' in the local databases.")

    return comp_specs


# ============================================================================
# Parts that are never crossed
#
# Part numbers starting with any of these prefixes have no TI equivalent -
# 1N is the JEDEC registration series, the rest are legacy zener and
# high-power TVS families. Nor does anything opening with a power rating in
# kilowatts (1.5KE, 5KP, 15KPA), which is a leaded high-power body.
# They are blocked here, at the two entry points to the TI search, which
# covers every competitor and every mode. Their specs still look up
# normally; only the cross is refused.
# ============================================================================

# A Nexperia automotive part is marked by a "-Q" suffix on the part number
# (BZB84-B9V1-Q). Matches "-Q" at the end of the name or followed by a
# separator, so packing suffixes like "-Q,115" still read as automotive, while
# TI's own "-Q1" naming is not what this is looking at.
NEXPERIA_AUTO_SUFFIX_RE = re.compile(r"-Q(?![A-Z0-9])", re.IGNORECASE)


def is_nexperia_automotive(part_name):
    """Whether a Nexperia part number carries the automotive '-Q' suffix."""
    return bool(NEXPERIA_AUTO_SUFFIX_RE.search(str(part_name or "").strip()))


def is_automotive_grade(grade_text):
    """
    Whether a competitor's Grade string means automotive. Tested for the
    negation first - "Non-Automotive" contains "automotive", so a plain
    substring check reads every commercial part as automotive.
    """
    text = str(grade_text).lower()
    if "non-automotive" in text or "non automotive" in text or "not automotive" in text:
        return False
    return "automotive" in text or "aec-q" in text


NO_CROSS_PREFIXES = (
    "1N", "2Z", "Z1", "1Z", "2E", "MRD", "MZ", "HZ", "MT", "2E", "3E",
    "P4", "SA", "BZM", "GS3", "ZMM", "BZV", "GLL", "GLZ", "CLL", "TLZ", "TZM", "TZX", "Z4", "ZM4", "ZMC", "ZGL", "ZMY", "ZPY", "BYZ", "NZX", "P6K"
)

# A leading kilowatt rating. The decimal point is stripped before this runs,
# so "1.5KE440A" arrives as "15KE440A" and matches the same way.
NO_CROSS_KILOWATT = re.compile(r'^\d+K')


def is_no_cross_part(part_name):
    normalized = re.sub(r'[\W_]+', '', str(part_name).upper())
    if NO_CROSS_KILOWATT.match(normalized):
        return True
    return normalized.startswith(NO_CROSS_PREFIXES)


def find_ti_zener_alternatives(competitor_specs, ti_zener_specs_df):
    comp_device_name = competitor_specs.get("Device Name", "-")
    if is_no_cross_part(comp_device_name):
        print(f"'{comp_device_name}' starts with a no-cross prefix - not crossed.")
        return []

    # SMx (SMA/SMB/SMC/SMD/SMF/SM8) and DO- bodies are never crossed. Both the
    # normalized package and the canonical code are checked: normalization
    # rewrites some of these into a TI name, and only one of the two may
    # actually spell the body out.
    if is_blocked_competitor_package(competitor_specs.get("Package")) or \
            is_blocked_competitor_package(competitor_specs.get("Canonical Package")):
        print(f"'{comp_device_name}' is an SMx / DO- package - not crossed.")
        return []

    if ti_zener_specs_df is None or ti_zener_specs_df.empty:
        return []

    print("\n--- Starting TI Zener Diode Search ---")

    SCORE_PKG_MATCH = 4000
    SCORE_VZ_EXACT = 2000
    SCORE_VZ_CLOSE = 1000
    SCORE_TOLERANCE_MATCH = 500
    SCORE_POWER_MATCH = 300
    SCORE_GRADE_MATCH = 500
    VZ_BAND_PERCENT = 30      # matches the +/-30% Vz window parts are kept within

    comp_is_automotive = is_automotive_grade(competitor_specs.get("Grade", "-"))

    comp_pkg_alias = competitor_specs.get("Package", "")
    is_structured_pkg = isinstance(comp_pkg_alias, dict)
    valid_comp_pkgs_for_match = []
    if not is_structured_pkg:
        valid_comp_pkgs_for_match = [normalize_package(p) for p in str(comp_pkg_alias).split(',') if p.strip()]

    comp_vz_str = competitor_specs.get("Voltage - Reverse Standoff (Typ)", None)
    comp_vz = to_numeric_val(comp_vz_str, default_if_error=None)
    comp_tol_str = competitor_specs.get("Tolerance", "")
    comp_tol = float('inf')
    comp_tol_num_match = re.search(r'([\d.]+)', comp_tol_str)
    if comp_tol_num_match:
        matched_str = comp_tol_num_match.group(1)
        if matched_str != '.':
            try:
                comp_tol = float(matched_str)
            except ValueError:
                pass
    comp_pd_str = competitor_specs.get("Power Dissipation (Pd)", "")
    comp_pd = to_numeric_val(comp_pd_str, default_if_error=None)

    if comp_vz is None:
        print("Competitor Zener Voltage (Vz) is not available. Cannot find alternatives.")
        return []

    part_num_col = next((c for c in ti_zener_specs_df.columns if 'product or part number' in c.lower()), None)
    pkg_col = next((c for c in ti_zener_specs_df.columns if 'package name' in c.lower()), None)
    vz_col = next((c for c in ti_zener_specs_df.columns if 'vz (nom) (v)' in c.lower()), None)
    tol_col = next((c for c in ti_zener_specs_df.columns if 'tolerance' in c.lower()), None)
    pd_col = next((c for c in ti_zener_specs_df.columns if 'pd (max) (w)' in c.lower()), None)

    if not all([part_num_col, pkg_col, vz_col, tol_col, pd_col]):
        print("! Warning: One or more required columns not found in the TI Zener database.")
        return []

    _, package_matches_df, _ = find_package_matches(comp_pkg_alias, ti_zener_specs_df)
    if package_matches_df.empty:
        print("Warning: No TI Zener parts with a matching package found. Scoring on parameters only.")

    potential_alternatives = []
    target_df = ti_zener_specs_df.copy()
    target_df[vz_col] = pd.to_numeric(target_df[vz_col], errors='coerce')

    for index, ti_row in target_df.iterrows():
        score = 0
        ti_vz = ti_row.get(vz_col)
        if pd.isna(ti_vz) or not (comp_vz * 0.70 <= ti_vz <= comp_vz * 1.30 and ti_vz * 0.70 <= comp_vz <= ti_vz * 1.30):
            continue

        is_pkg_match = index in package_matches_df.index
        if is_pkg_match:
            score += SCORE_PKG_MATCH

        # Vz is scored on a slope rather than in two buckets, so how close a
        # part actually sits decides the ranking. Within 1% still earns the full
        # SCORE_VZ_EXACT; past that it falls off linearly to SCORE_VZ_CLOSE at
        # the edge of the +/-30% band the loop above already enforces.
        vz_diff_percent = abs(ti_vz - comp_vz) / comp_vz * 100
        if vz_diff_percent <= 1:
            score += SCORE_VZ_EXACT
        else:
            _vz_fraction = min((vz_diff_percent - 1) / (VZ_BAND_PERCENT - 1), 1.0)
            score += SCORE_VZ_EXACT - _vz_fraction * (SCORE_VZ_EXACT - SCORE_VZ_CLOSE)

        ti_tol_str = str(ti_row.get(tol_col, ''))
        ti_tol_num_match = re.search(r'([\d.]+)', ti_tol_str)
        if ti_tol_num_match:
            ti_tol = float(ti_tol_num_match.group(1))
            if ti_tol <= comp_tol:
                score += SCORE_TOLERANCE_MATCH

        if comp_pd is not None:
            ti_pd = pd.to_numeric(ti_row.get(pd_col), errors='coerce')
            if pd.notna(ti_pd) and (comp_pd * 0.90 <= ti_pd <= comp_pd * 1.10):
                score += SCORE_POWER_MATCH

        # An automotive competitor part crosses to TI's automotive (-Q1) part
        # and a commercial one to the commercial part. Without this the two are
        # identical on every other term and the tie breaks arbitrarily.
        ti_part_name = str(ti_row[part_num_col]).strip()
        ti_is_automotive = ti_part_name.upper().endswith("-Q1")
        if ti_is_automotive == comp_is_automotive:
            score += SCORE_GRADE_MATCH

        if score > 0:
            potential_alternatives.append({
                "part_number": ti_row[part_num_col],
                "score": score,
                "is_package_match": is_pkg_match,
                "vz": ti_vz,
                "vz_diff": vz_diff_percent,
                "tolerance": ti_row.get(tol_col, '-'),
                "pd": ti_row.get(pd_col, '-'),
                "is_automotive": ti_is_automotive
            })

    if not potential_alternatives:
        return []

    sorted_alternatives = sorted(
        potential_alternatives,
        key=lambda x: (x["score"], x["is_package_match"], x["is_automotive"] == comp_is_automotive, -x["vz_diff"]),
        reverse=True,
    )
    print("Top potential Zener alternatives:")
    for i, alt in enumerate(sorted_alternatives[:5]):
        print(f"{i+1}. {alt['part_number']} - Score: {round(alt['score'], 1)}, PkgMatch: {alt['is_package_match']}, Auto: {alt['is_automotive']}, Vz: {alt['vz']} ({alt['vz_diff']:.1f}% off), Tol: {alt['tolerance']}, Pd: {alt['pd']}")
    return sorted_alternatives[:3]


# ESD (IEC 61000-4-2) scoring. Unlike the other pass/fail parameters, a higher
# rating keeps earning points: meeting the competitor's number is worth the
# tier's base, and headroom above it adds up to the tier's headroom award,
# reached at the tier's cap multiple of the competitor's rating.
#
# Two tiers, keyed on the COMPETITOR's rating. A part already rated 20 kV or
# better was chosen for its ESD performance, so beating it counts for more --
# and since TI's own ceiling is not far above 20 kV, the high tier reaches its
# full award at a smaller multiple than doubling.
SCORE_ESD_BASE = 250
SCORE_ESD_HEADROOM = 750
ESD_HEADROOM_CAP = 2.0

ESD_HIGH_TIER_THRESHOLD = 20.0    # competitor kV at or above which the boost applies
SCORE_ESD_BASE_HIGH = 500
SCORE_ESD_HEADROOM_HIGH = 1500
ESD_HEADROOM_CAP_HIGH = 1.5


# On top of the relative score above, a flat bonus for the TI part's OWN rating.
# The tiers above measure a part against the competitor, which cannot separate a
# 30 kV part from a 22 kV one once both have cleared the headroom cap - this
# does. Highest threshold first; the first one met wins.
#
# The TOP tier only counts when the competitor is itself a high-ESD part
# (>= ESD_HIGH_TIER_THRESHOLD). Against a lower-rated competitor the extra
# headroom is not what the design called for, so a top-tier part is worth no
# more than the tier below it.
ESD_ABSOLUTE_TIERS = (
    (30.0, 1000),   # TI rated 30 kV or better
    (25.0,  500),   # 25 kV or better
)

# Top-of-range match. A competitor rated at the ceiling of the ESD range can
# only be answered by a TI part that is also there - nothing below it even
# scores, since a part under the competitor's rating earns no ESD points at
# all. When both sides are at the top, that match is worth more than the tiers
# above can express on their own, so it earns this on top of them.
ESD_TOP_MATCH_THRESHOLD = 30.0
SCORE_ESD_TOP_MATCH = 1000


def _esd_kv(esd_value):
    """
    ESD ratings in kV, whatever the source used. to_numeric_val expands a "kV"
    string to volts (30 kV -> 30000) while a bare number stays as written, so
    both spellings reach the tier thresholds - which are in kV. Real ratings top
    out around 30 kV, so anything past 1000 is volts.
    """
    try:
        value = float(esd_value)
    except (TypeError, ValueError):
        return 0.0
    return value / 1000.0 if value > 1000 else value


def _esd_tier(comp_esd):
    """(base score, headroom award, headroom cap multiple) for a competitor rating."""
    if _esd_kv(comp_esd) >= ESD_HIGH_TIER_THRESHOLD:
        return SCORE_ESD_BASE_HIGH, SCORE_ESD_HEADROOM_HIGH, ESD_HEADROOM_CAP_HIGH
    return SCORE_ESD_BASE, SCORE_ESD_HEADROOM, ESD_HEADROOM_CAP


def _esd_absolute_bonus(ti_esd, comp_esd):
    """
    Flat bonus for the TI part's own ESD rating. The top tier is only in play
    when the competitor is a high-ESD part in its own right; below that it is
    dropped, so a top-tier part earns the same as the tier beneath it.
    """
    ti_kv = _esd_kv(ti_esd)
    tiers = ESD_ABSOLUTE_TIERS
    if _esd_kv(comp_esd) < ESD_HIGH_TIER_THRESHOLD:
        tiers = ESD_ABSOLUTE_TIERS[1:]
    for threshold, bonus in tiers:
        if ti_kv >= threshold:
            return bonus
    return 0


def _esd_top_match_bonus(ti_esd, comp_esd):
    """
    Extra award when the competitor is at the top of the ESD range and the TI
    part matches it there. Stacks on the tier score and the absolute bonus.
    """
    if (_esd_kv(comp_esd) >= ESD_TOP_MATCH_THRESHOLD
            and _esd_kv(ti_esd) >= ESD_TOP_MATCH_THRESHOLD):
        return SCORE_ESD_TOP_MATCH
    return 0

# A direction match is what makes a part a drop-in, so it outweighs a raw
# parameter win: a bidirectional TI part with a better ESD rating should not
# take the cross from a unidirectional part that matches the competitor.
SCORE_DIRECTION_MATCH = 4000

# Tie-breaker only. Two parts can sit the same distance from the competitor on
# opposite sides - 12 V and 18 V are both 20% off a 15 V competitor - so the
# Vrwm tiers below cannot separate them. Deliberately far smaller than the
# 500-point gap between those tiers, so it can never promote a worse Vrwm
# match over a better one.
SCORE_VRWM_HEADROOM = 50

MMBZ_VARIANT_RE = re.compile(r'^MMBZ\d*(V[A-Z]{2})')

# Interface families both TI and its competitors spell out in the part name.
# Longest first, so ESD2CANFD24 is read as CANFD rather than CAN, and ETHERNET
# is not cut short by anything. Kept to protocol names -- a generic token like
# ESD or TVS appears in most of the catalogue and would match everything.
INTERFACE_FAMILY_TOKENS = (
    "FLEXRAY", "ETHERNET", "CANFD", "RS485", "RS232",
    "HDMI", "SATA", "MIPI", "USB", "CAN", "LIN",
)


def _interface_family(part_name):
    """
    The interface a part is named for, or None: PESD1LIN -> LIN,
    ESD1LIN24DYFR -> LIN, TSD24CDYFR -> None.
    """
    flat = re.sub(r'[\W_]+', '', str(part_name or "")).upper()
    for token in INTERFACE_FAMILY_TOKENS:
        if token in flat:
            return token
    return None

def _mmbz_variant(part_name):
    """
    The clamp-family token out of an MMBZ part number, or None for anything
    that is not one: MMBZ27VALQ -> VAL, MMBZ27VCLDBZRQ1 -> VCL.

    The families are not interchangeable, so a VAL competitor has to take the
    VAL part and a VCL the VCL, whatever the parametrics say.
    """
    m = MMBZ_VARIANT_RE.match(re.sub(r'[\W_]+', '', str(part_name).upper()))
    return m.group(1) if m else None

def _row_is_package_match(comp_pkg_alias, ti_row_df):
    """
    Re-runs the stage 1 / stage 2 package match against a single TI row.

    Grade correction swaps an automotive TI part for its commercial sibling,
    but the two are separate rows with separate 'Package name' cells and do not
    have to ship in the same bodies. The commercial row therefore has to earn
    the package match on its own cell instead of inheriting the automotive
    row's. find_package_matches is reused rather than reimplemented so
    structured and unstructured competitor packages follow the identical rules.
    """
    if ti_row_df is None or ti_row_df.empty:
        return False
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            _stage1, _stage2, pin_matched_df = find_package_matches(comp_pkg_alias, ti_row_df)
    except Exception:
        return False
    return not pin_matched_df.empty

def find_ti_alternatives(competitor_specs, ti_specs_df):
    comp_device_name = competitor_specs.get("Device Name", "-")
    if is_no_cross_part(comp_device_name):
        print(f"'{comp_device_name}' starts with a no-cross prefix - not crossed.")
        return []

    # SMx (SMA/SMB/SMC/SMD/SMF/SM8) and DO- bodies are never crossed. Both the
    # normalized package and the canonical code are checked: normalization
    # rewrites some of these into a TI name, and only one of the two may
    # actually spell the body out.
    if is_blocked_competitor_package(competitor_specs.get("Package")) or \
            is_blocked_competitor_package(competitor_specs.get("Canonical Package")):
        print(f"'{comp_device_name}' is an SMx / DO- package - not crossed.")
        return []
    
    if ti_specs_df is None or ti_specs_df.empty:
        return []
    ti_specs_df.columns = ti_specs_df.columns.str.strip()

    comp_grade = competitor_specs.get("Grade", "-").lower()
    comp_is_automotive = is_automotive_grade(comp_grade)
    comp_pkg_alias = competitor_specs.get("Package", "")
    is_structured_pkg = isinstance(comp_pkg_alias, dict)
    valid_comp_pkgs_for_match = []
    if not is_structured_pkg:
        valid_comp_pkgs_for_match = [p.strip() for p in str(comp_pkg_alias).split(',') if p.strip()]
    comp_v_clamp_str = competitor_specs.get("Voltage - Clamping (Max) @ Ipp", "-")
    comp_v_clamp = to_numeric_val(comp_v_clamp_str if comp_v_clamp_str and comp_v_clamp_str != "-" else "inf")
    comp_cap_str = competitor_specs.get("Capacitance", "-")
    comp_cap = to_numeric_val(comp_cap_str if comp_cap_str and comp_cap_str != "-" else "inf")
    comp_vrw_str = competitor_specs.get("Voltage - Reverse Standoff (Typ)", "-")
    comp_vrw_numeric = to_numeric_val(comp_vrw_str, default_if_error=None)
    comp_direction = competitor_specs.get("Direction", "-").lower()
    comp_channels_str = competitor_specs.get("Channels", "1")
    comp_surge_str = competitor_specs.get("IEC 61000-4-5", "-")
    comp_surge = to_numeric_val(comp_surge_str, default_if_error=0)
    comp_esd_str = competitor_specs.get("IEC 61000-4-2", "-")
    comp_esd = to_numeric_val(comp_esd_str, default_if_error=0)
    comp_ppp_str = competitor_specs.get("Power Dissipation (Pd)", "-")
    comp_ppp = to_numeric_val(comp_ppp_str, default_if_error=0)

    pkg_display_for_log = comp_pkg_alias if is_structured_pkg else valid_comp_pkgs_for_match
    print(f"Grade: {'Automotive' if comp_is_automotive else 'Non-Automotive'}, Normalized Competitor Pkgs: {pkg_display_for_log}")
    base_ti_df_for_search = ti_specs_df.copy()

    package_matches_df = pd.DataFrame()
    vrwm_matches_df = pd.DataFrame()

    stage1_package_matches_df, final_package_matches_df, pin_matched_df = find_package_matches(comp_pkg_alias, base_ti_df_for_search)
    pin_match_indices = set(pin_matched_df.index)
    final_package_match_indices = set(final_package_matches_df.index)
    VRWM_TOLERANCE_PERCENT = 15
    if comp_vrw_numeric is not None and 'Vrwm (V)' in base_ti_df_for_search.columns:
        vrwm_matched_indices = base_ti_df_for_search['Vrwm (V)'].apply(
            lambda ti_vrwm: is_similar_voltage(str(ti_vrwm), comp_vrw_str, tolerance_percent=VRWM_TOLERANCE_PERCENT, two_sided=False)
        )
        vrwm_matches_df = base_ti_df_for_search[vrwm_matched_indices].copy()
        print(f"Found {len(vrwm_matches_df)} TI parts with Vrwm within {VRWM_TOLERANCE_PERCENT}% tolerance.")

    all_candidates_df = pd.concat([stage1_package_matches_df, vrwm_matches_df]).drop_duplicates()
    target_df_for_params = all_candidates_df if not all_candidates_df.empty else base_ti_df_for_search
    if all_candidates_df.empty:
        print(f"FindTI: Stage 2 - No package or Vrwm matches. Checking all {len(target_df_for_params)} parts.")
    else:
        print(f"FindTI: Stage 2 - Parametrically checking {len(target_df_for_params)} package/Vrwm-matched parts.")

    potential_alternatives = []
    VCLAMP_TOLERANCE = 3.1

    # Capacitance tiers, keyed on the competitor's own capacitance. The lower it
    # is, the more it matters: a low-cap part is chosen for being low-cap, so a
    # TI part that gives that up is not an alternative at all.
    #   (upper bound of comp_cap, max headroom over comp_cap, weight)
    # Above the last tier capacitance is a "don't care" - a lower TI value still
    # scores, but a higher one no longer disqualifies.
    CAP_TIERS = (
        (1.0,  3.00, 6000),   # <=1 pF   : TI may exceed by 200%
        (3.6,  1.40, 4000),   # 1-2 pF   : by 40%
        (10.0, 1.50, 2500),   # 2-10 pF  : by 50%
    )
    CAP_DONT_CARE_WEIGHT = 1000

    def _cap_tier(cap):
        """(headroom multiplier or None for don't-care, score weight)."""
        for upper, headroom, weight in CAP_TIERS:
            if cap <= upper:
                return headroom, weight
        return None, CAP_DONT_CARE_WEIGHT

    # Counts of which hard filter rejected each part, so an empty result can be
    # explained instead of guessed at. A part can fail more than one.
    reject_counts = {"clamping": 0, "capacitance": 0, "Vrwm": 0,
                     "direction mismatch (scored, not rejected)": 0}

    for index, ti_row in target_df_for_params.iterrows():
        ti_part_num = safe_strip(ti_row.get("Product or Part number"))
        if ti_part_num.upper().startswith(('UC', 'SN')):
            continue
        ti_pkg_excel_original = safe_strip(ti_row.get("Package name"))
        ti_v_clamp_str = str(ti_row.get("Clamping voltage (V)", "inf"))
        ti_v_clamp = to_numeric_val(ti_v_clamp_str)
        vclamp_ok = ti_v_clamp <= (comp_v_clamp * VCLAMP_TOLERANCE) if comp_v_clamp != float('inf') else True

        ti_cap_str = str(ti_row.get("IO capacitance (typ) (pF)", "inf"))
        ti_cap = to_numeric_val(ti_cap_str)
        if comp_cap == float('inf'):
            cap_ok = True
        else:
            cap_headroom, _cap_weight = _cap_tier(comp_cap)
            # No headroom limit above 10 pF; below it, TI may not exceed the
            # competitor by more than the tier allows. Lower is always fine.
            # Epsilon so a value sitting exactly on the limit is not lost to
            # float error (2.0 * 1.40 does not land exactly on 2.8).
            cap_ok = True if cap_headroom is None else (ti_cap <= comp_cap * cap_headroom + 1e-9)

        ti_vrw_str_excel = str(ti_row.get("Vrwm (V)", "-"))
        vrw_filter_passed = False
        if comp_vrw_numeric is None:
            vrw_filter_passed = True
        else:
            vrw_filter_passed = is_similar_voltage(ti_vrw_str_excel, comp_vrw_str, tolerance_percent=VRWM_TOLERANCE_PERCENT, two_sided=False)

        ti_direction_raw = safe_strip(str(ti_row.get("Bi-/uni-directional", ""))).lower()
        ti_direction = "unidirectional" if "uni-directional" in ti_direction_raw else \
                       "bidirectional" if "bi-directional" in ti_direction_raw else "unknown"
        # Direction is a strong preference, not a gate. A bidirectional
        # competitor part still crosses to a unidirectional TI part, but scores
        # far below one that matches, so a matching part always outranks it.
        direction_match = (comp_direction in ["-", "unknown"] or ti_direction == "unknown" or comp_direction == ti_direction)
        direction_known_both = (comp_direction not in ["-", "unknown"] and ti_direction != "unknown")
        direction_conflict = direction_known_both and comp_direction != ti_direction

        num_channels_ti_raw = ti_row.get("Number of channels")
        channels_match = False
        try:
            ti_channels_val = int(float(safe_strip(str(num_channels_ti_raw))))
            comp_channels_val = int(float(comp_channels_str))
            channels_match = (ti_channels_val == comp_channels_val)
        except (ValueError, TypeError):
            channels_match = False

        ti_is_automotive = ti_part_num.endswith("-Q1")
        grade_match = (comp_is_automotive == ti_is_automotive)

        score = 0
        is_pkg_and_pin_match = index in pin_match_indices
        is_pkg_match_for_score = index in final_package_match_indices
        if is_pkg_and_pin_match:
            score += 6000   # package + pin count match
        elif is_pkg_match_for_score:
            score += 2000   # package match, wrong/unknown pin count
        if grade_match: score += 500
        if channels_match: score += 3000
        if direction_conflict:
            pass                          # crossable, but earns nothing here
        elif direction_known_both:
            score += SCORE_DIRECTION_MATCH                 # both stated and they agree
        elif direction_match:
            score += 250                  # one side unknown - nothing to compare

        ti_surge_str = str(ti_row.get("IEC 61000-4-5 (A)", "0"))
        ti_surge = to_numeric_val(ti_surge_str, default_if_error=0)
        ti_esd_str = str(ti_row.get("IEC 61000-4-2 contact (k±V)", "0"))
        ti_esd = to_numeric_val(ti_esd_str, default_if_error=0)
        if ti_surge >= comp_surge and comp_surge > 0: score += 500
        if ti_esd >= comp_esd and comp_esd > 0:
            # Scaled: the further above the competitor's rating, the more it is
            # worth, flattening out at the tier's cap multiple. The tier itself
            # is set by how high the competitor's own rating is.
            _esd_base, _esd_weight, _esd_cap = _esd_tier(comp_esd)
            _esd_headroom = (ti_esd - comp_esd) / comp_esd
            _esd_fraction = min(_esd_headroom / (_esd_cap - 1.0), 1.0) if _esd_cap > 1.0 else 0.0
            score += _esd_base + _esd_fraction * _esd_weight
            score += _esd_absolute_bonus(ti_esd, comp_esd)
            score += _esd_top_match_bonus(ti_esd, comp_esd)
        ti_ppp = to_numeric_val(str(ti_row.get("Peak pulse power (8/20 μs) (max) (W)", "0")), default_if_error=0)
        if ti_ppp >= comp_ppp and comp_ppp > 0: score += 500

        ti_vrw_numeric_current = to_numeric_val(ti_vrw_str_excel, default_if_error=None)
        exact_vrw_match = False
        if vrw_filter_passed and comp_vrw_numeric is not None and ti_vrw_numeric_current is not None:
            vrwm_diff_percent = abs(comp_vrw_numeric - ti_vrw_numeric_current) / comp_vrw_numeric * 100
            if vrwm_diff_percent <= 5:
                exact_vrw_match = True
                score += 2000
            elif vrwm_diff_percent <= 15:
                score += 1500
            else:
                score += 300
                # Same distance, opposite sides: the higher standoff wins. A TVS
                # sitting below the competitor's Vrwm starts conducting inside the
                # operating range, so it is the worse part even at equal error.
                if ti_vrw_numeric_current >= comp_vrw_numeric:
                    score += SCORE_VRWM_HEADROOM

        if comp_cap != float('inf') and ti_cap != float('inf'):
            _headroom, max_cap_score = _cap_tier(comp_cap)
            if ti_cap <= comp_cap:
                # Matching or beating the competitor is what earns the weight,
                # scaled by how much lower TI sits.
                improvement_ratio = (comp_cap - ti_cap) / comp_cap
                score += 200 + improvement_ratio * max_cap_score
            elif cap_ok:
                score += 100

        meets_min_criteria_relaxed = True
        if not vclamp_ok:
            meets_min_criteria_relaxed = False
            reject_counts["clamping"] += 1
        if not cap_ok:
            meets_min_criteria_relaxed = False
            reject_counts["capacitance"] += 1
        if direction_conflict:
            reject_counts["direction mismatch (scored, not rejected)"] += 1
        if comp_vrw_numeric is not None and not vrw_filter_passed:
            meets_min_criteria_relaxed = False
            reject_counts["Vrwm"] += 1
        if meets_min_criteria_relaxed and score > 0:
            potential_alternatives.append({
                "part_number": ti_part_num, "score": score, "v_clamp": ti_v_clamp,
                "capacitance": ti_cap, "vrw": ti_vrw_str_excel,
                "direction": ti_direction.capitalize(), "package": ti_pkg_excel_original,
                "is_package_match": is_pkg_and_pin_match, "pkg_only_match": is_pkg_match_for_score, "exact_vrw_match": exact_vrw_match,
                "is_automotive": ti_is_automotive, "esd": ti_esd
            })

    if potential_alternatives:
        def get_base_part_number(pn):
            return pn.replace("-Q1", "") if pn.upper().endswith("-Q1") else pn
        final_alternatives = []
        processed_base_names = set()
        sorted_for_dedup = sorted(potential_alternatives, key=lambda x: x['score'], reverse=True)
        for alt in sorted_for_dedup:
            base_name = get_base_part_number(alt["part_number"])
            if base_name in processed_base_names:
                continue
            is_wrong_grade = alt["is_automotive"] != comp_is_automotive
            if is_wrong_grade and not comp_is_automotive:
                commercial_part_name = base_name
                commercial_part_row = base_ti_df_for_search[base_ti_df_for_search['Product or Part number'].astype(str).str.strip().str.lower() == commercial_part_name.lower()]
                if not commercial_part_row.empty:
                    print(f"Grade Correction - Swapping automotive '{alt['part_number']}' for commercial '{commercial_part_name}'.")
                    correct_alt = {
                        "part_number": commercial_part_name,
                        "score": alt['score'] + 1,
                        "is_automotive": False,
                        "is_package_match": _row_is_package_match(comp_pkg_alias, commercial_part_row),
                        "exact_vrw_match": alt.get("exact_vrw_match", False),
                        "v_clamp": to_numeric_val(str(commercial_part_row.iloc[0].get("Clamping voltage (V)", "inf"))),
                        "capacitance": to_numeric_val(str(commercial_part_row.iloc[0].get("IO capacitance (typ) (pF)", "inf"))),
                        "vrw": alt.get("vrw", "-"),
                        "direction": alt.get("direction", "-"),
                        "package": safe_strip(commercial_part_row.iloc[0].get("Package name")),
                        "esd": to_numeric_val(str(commercial_part_row.iloc[0].get("IEC 61000-4-2 contact (k±V)", "0")), default_if_error=0)
                    }
                    final_alternatives.append(correct_alt)
                else:
                    final_alternatives.append(alt)
            else:
                final_alternatives.append(alt)
            processed_base_names.add(base_name)

        print(f"Deduplicated and grade-corrected {len(potential_alternatives)} candidates down to {len(final_alternatives)}.")
        potential_alternatives = final_alternatives

    if not potential_alternatives:
        print("Found 0 potential TI alternatives after filtering and scoring.")
        _not_considered = len(base_ti_df_for_search) - len(target_df_for_params)
        print(f"  Rejected by filter (a part can fail more than one): {reject_counts}")
        if _not_considered > 0:
            print(f"  A further {_not_considered} TI parts were never considered - no package "
                  f"match and Vrwm outside tolerance of {comp_vrw_str}.")
        return []
    
    comp_mmbz_variant = _mmbz_variant(comp_device_name)
    comp_interface_family = _interface_family(comp_device_name)
    # An SMA / SMB competitor crosses only into its own body, so anything
    # without a package match is dropped outright rather than ranked lower.
    if requires_exact_package_match(competitor_specs.get("Package")) or \
            requires_exact_package_match(competitor_specs.get("Canonical Package")):
        _kept = [a for a in potential_alternatives if a.get("is_package_match")]
        if len(_kept) != len(potential_alternatives):
            print(f"  SMx competitor body - dropped {len(potential_alternatives) - len(_kept)} "
                  f"alternative(s) with no package match.")
        potential_alternatives = _kept
        if not potential_alternatives:
            print("Found 0 potential TI alternatives after filtering and scoring.")
            return []
        
    sorted_alternatives = sorted(
        potential_alternatives,
        key=lambda alt: (
            # An automotive competitor crosses to an automotive TI part. This
            # outranks the score, so a commercial part cannot take the cross by
            # winning on capacitance or ESD. No effect on a commercial
            # competitor -- grade correction already handles that direction.
            alt["is_package_match"],
            (alt["is_automotive"] if comp_is_automotive else True),
            # A package match is the difference between a Q and a P, so every
            # Q outranks every P regardless of score. A part in the wrong body
            # is not a drop-in however well it scores parametrically.

            # MMBZ clamp families cross like-for-like: a VAL competitor takes
            # the VAL part, a VCL the VCL. Ranked above the score so a better
            # parametric fit in the wrong family cannot take the cross.
            (comp_mmbz_variant is not None
             and _mmbz_variant(alt["part_number"]) == comp_mmbz_variant),
            
            alt["score"],
            (comp_interface_family is not None
             and _interface_family(alt["part_number"]) == comp_interface_family),
            alt["exact_vrw_match"],
            alt["exact_vrw_match"],
            alt["capacitance"],
            -alt["v_clamp"],
        ),
        reverse=True
    )

    print(f"Found {len(sorted_alternatives)} potential TI alternatives after filtering and scoring.")
    if not sorted_alternatives:
        _not_considered = len(base_ti_df_for_search) - len(target_df_for_params)
        print(f"  Rejected by filter (a part can fail more than one): {reject_counts}")
        if _not_considered > 0:
            print(f"  A further {_not_considered} TI parts were never considered - no package "
                  f"match and Vrwm outside tolerance of {comp_vrw_str}.")
    print("Top potential alternatives before final selection:")
    for i, alt in enumerate(sorted_alternatives[:5]):
        print(f"{i+1}. {alt['part_number']} - Score: {round(alt['score'], 1)}, PkgMatch: {alt['is_package_match']}, Auto: {alt['is_automotive']}, Vcl: {alt['v_clamp']}, Cap: {alt['capacitance']}, ESD: {alt.get('esd', '-')}")

    return sorted_alternatives[:3]


def manage_data_files(force_reload=False):
    ti_file = "ti_specs.xlsx"
    ti_zener_file = "ti_zener_specs.xlsx"

    if force_reload:
        print("\n--- Force reloading TI data files ---")
        download_and_rename_ti_specs()
        download_and_rename_ti_zener_specs()
        print("\n--- TI data files reloaded. ---")
        return

    if not os.path.exists(ti_file):
        print(f"'{ti_file}' not found. Downloading...")
        download_and_rename_ti_specs()

    if not os.path.exists(ti_zener_file):
        print(f"'{ti_zener_file}' not found. Downloading...")
        download_and_rename_ti_zener_specs()


def _pin_from_name(norm_pkg, available_pins=None):
    """
    Infer a package's pin count from its name.
    `available_pins` lets generic names (SOT-5X3, SOT-9X3) pick the plausible
    value actually present in the row.
    """
    p = normalize_package(norm_pkg)
    avail = set(available_pins or [])

    # Explicit trailing pin count: SC70-6 -> SC706, SOT-23-6 -> SOT236
    m = re.match(r'^(SC70|SOT23)(\d+)$', p)
    if m:
        return int(m.group(2))

    # SOT-5X3 / SOT-9X3: middle digit is the pin count; X is a wildcard.
    m = re.match(r'^SOT([59])X3$', p)
    if m:
        # SOT-5X3 is built in 2-, 3-, 5- and 6-pin and SOT-9X3 in 3-, 5- and
        # 6-pin; neither name says which, so only commit when the row offers
        # exactly one of them.
        cands = [c for c in ((2, 3, 5, 6) if m.group(1) == '5' else (3, 5, 6)) if c in avail]
        if len(cands) == 1:
            return cands[0]
        return 3 if m.group(1) == '9' else None
    m = re.match(r'^SOT([59])(\d)3$', p)
    if m:
        # SOT-923 is the 3-lead member of the SOT-9X3 family.
        if p == 'SOT923':
            return 3
        return int(m.group(2))

    return _known_pin_count(p)


def _known_pin_count(norm_pkg):
    if norm_pkg.startswith("DFN1006"):   return 2
    if norm_pkg.startswith("DFN11103"):  return 3
    if norm_pkg.startswith("DFN2020"):   return 6
    if norm_pkg.startswith("DSBGA"):     return 4
    if norm_pkg.startswith("SC703"):     return 3
    if norm_pkg.startswith("SC706"):     return 6
    if norm_pkg.startswith("SOD323"):    return 2
    if norm_pkg.startswith("SOD523"):    return 2
    if norm_pkg.startswith("SOT233"):    return 3
    if norm_pkg.startswith("SOT234"):    return 4
    if norm_pkg.startswith("SOT235"):    return 5
    if norm_pkg.startswith("SOT236"):    return 6
    if norm_pkg.startswith("WQFN"):      return 12
    return None


# Authoritative pin counts for (GPN, package family) pairs that elimination
# cannot resolve. The TI "Pin count" column lists the distinct pin counts in a
# row, so when two packages can both claim the same value the row is ambiguous.
# Add entries here whenever a cross produces the wrong pin count / suffix.
# TI parts whose "USON" is the 1.6x1.6 variant (competitor DFN1616-6 equivalent).
TI_USON1616_PARTS = {"TPD4E001"}

TI_GPN_SUFFIX_OVERRIDES = {
    ("TVS2210", "DFN1006"): "YMZR",   # the only DFN1006 part not on DPYR
}


def _override_suffix(ti_gpn, pkg):
    """
    The suffix for a GPN that does not take its package's usual one.

    Consulted by BOTH suffix paths. The canonical path reads straight out of
    CANONICAL_SUFFIX_MAP, which is keyed by package and pin count alone and so
    cannot express a per-part exception -- ESD122 escapes that only by owning a
    distinct canonical code (DFN1006_3), while TVS2210 is an ordinary
    DFN1006_2 part that merely ships under a different suffix.
    """
    gpn = str(ti_gpn or "").upper().replace("-Q1", "").strip()
    norm = normalize_package(pkg)
    for (g, fam), suffix in TI_GPN_SUFFIX_OVERRIDES.items():
        if g in gpn and norm.startswith(normalize_package(fam)):
            return suffix
    return None

TI_PACKAGE_PIN_OVERRIDES = {
    ("TPD4E001",   "USON"): 6,    # the only 6-pin USON
}

# A handful of automotive parts carry a qualification letter between the GPN
# base and the package suffix -- TPD1E05U06 + Q + DPYR + Q1. This is NOT the
# general rule for -Q1 parts: every automotive GPN not listed here keeps the
# plain gpn_base + suffix + "Q1" form. Keyed on the GPN with -Q1 stripped.
TI_AUTOMOTIVE_QUALIFIERS = {
    "TPD4E02B04": "Q",   # TPD4E02B04QDQARQ1
    "TPD1E05U06": "Q",   # TPD1E05U06QDPYRQ1
    "TPD1E10B06": "Q",   # TPD1E10B06QDPYRQ1
    "TPD1E10B09": "Q",
    "TPD2E2U06":  "Q",
    "TPD4E05U06": "Q",
    "TPD4E001":   "Q",
    "TPD2E001":   "I",   # TPD2E001IDRLRQ1 -- the odd one out
}


def _automotive_opn(gpn_base, suffix):
    """
    gpn_base + suffix + "Q1", with a qualification letter inserted before the
    suffix for the parts listed in TI_AUTOMOTIVE_QUALIFIERS.
    """
    qualifier = TI_AUTOMOTIVE_QUALIFIERS.get(str(gpn_base).upper().strip(), "")
    return f"{gpn_base}{qualifier}{suffix}Q1"

def _override_pin(ti_gpn, norm_pkg):
    gpn = str(ti_gpn).upper().replace("-Q1", "").strip()
    p = normalize_package(norm_pkg)
    for (g, fam), pins in TI_PACKAGE_PIN_OVERRIDES.items():
        if g == gpn and p.startswith(normalize_package(fam)):
            return pins
    return None


def _resolve_pin_count(target_norm_pkg, all_packages_str, all_pins_str, ti_gpn):
    """
    Work out how many pins `target_norm_pkg` has for this TI part.

    The TI "Pin count" column lists the DISTINCT pin counts present in the row,
    not one value per package, so position is meaningless. Strategy:
      1. Read the count straight out of the package name where possible.
      2. Otherwise eliminate the counts claimed by named packages; whatever
         remains belongs to the unnamed ones.
    """
    packages = [normalize_package(p.strip())
                for p in str(all_packages_str).split(",") if p.strip()]
    try:
        pins_sorted = sorted({int(p.strip())
                              for p in str(all_pins_str).split(",") if p.strip()})
    except (ValueError, TypeError):
        pins_sorted = []

    target = normalize_package(target_norm_pkg)

    forced = _override_pin(ti_gpn, target)
    if forced is not None:
        return forced

    if not pins_sorted:
        return _pin_from_name(target)

    # Whole row shares one pin count.
    if len(pins_sorted) == 1 or len(packages) <= 1:
        return pins_sorted[0]

    # Part-specific exception.
    if "ESD122" in str(ti_gpn).upper() and target.startswith("DFN1006"):
        return 3

    # 1) Name tells us outright — trust it when the row actually offers it.
    direct = _pin_from_name(target, pins_sorted)
    if direct is not None and direct in pins_sorted:
        return direct

    # 2) Eliminate counts accounted for by packages we can name.
    named, unknown = {}, []
    for pkg in packages:
        pc = _pin_from_name(pkg, pins_sorted)
        if pc is not None and pc in pins_sorted:
            named[pkg] = pc
        else:
            unknown.append(pkg)

    leftover = [p for p in pins_sorted if p not in set(named.values())]

    if target in unknown:
        # Pin values may be SHARED by several packages, so "one leftover value"
        # only forces a conclusion when exactly one package is unaccounted for.
        # With two unknowns (say USON and WSON) and 6 left over, USON might still
        # be 5 alongside SOT-5X3 while WSON takes the 6 - unknowable from the row.
        if len(leftover) == 1 and len(unknown) == 1:
            return leftover[0]
        if len(unknown) > 1:
            return None
        if leftover:
            plausible = [c for c in leftover if _plausible_pin_for_family(target, c)]
            # Only commit when the choice is forced.
            if len(plausible) == 1 and len([u for u in unknown
                                            if _plausible_pin_for_family(u, plausible[0])]) == 1:
                return plausible[0]
            return None
        # Named packages consumed everything; fall back to the row's counts.
        if direct is not None:
            return direct
        return pins_sorted[-1]

    if direct is not None:
        return direct
    if len(leftover) == 1:
        return leftover[0]
    return pins_sorted[0]


def _plausible_pin_for_family(norm_pkg, pin):
    """Whether `pin` is a pin count this package family is actually made in."""
    p = normalize_package(norm_pkg)
    allowed = {
        "USON":    {6},
        "WSON":    {15},
        "X2SON":   {3, 4, 6},
        "VSSOP":   {8, 10},
        "DFN1006": {2, 3},
        "DFN0603": {2, 4, 6},
        "DFN2020": {3, 6},
        "DFN2510": {6, 10},
        "DFN1616": {6},
        "DFN1110": {3},
        "DFN3030": {6, 8},
        "UQFN":    {10, 14},
        "WQFN":    {12},
        "DSBGA":   {4},
        "SOT886":  {6},
        "SOT5X3":  {2, 3, 5, 6},
        "SOT9X3":  {3, 5, 6},
    }
    for fam, pins in allowed.items():
        if p.startswith(fam):
            return pin in pins
    return True


def _get_suffix(norm_pkg, pin_val, ti_gpn):
    override = _override_suffix(ti_gpn, norm_pkg)
    if override:
        return override
    if norm_pkg.startswith("DFN0603"):
        if pin_val == 6:   return "DPFR"
        if pin_val == 4:   return "DPWR"
        if pin_val == 2:   return "DPLR"
        if pin_val == 3:   return "DMYR"
        if pin_val == 4:   return "DPWR"
        if pin_val == 6:   return "DPFR"
    elif norm_pkg.startswith("DFN1006"):
        if "ESD122" in ti_gpn.upper(): return "DMXR"
        return "DPYR"
    elif norm_pkg.startswith("DFN11103"):
        return "DXAR"
    elif norm_pkg.startswith("DFN2020"):
        return "DRVR"
    elif norm_pkg.startswith("DFN2510"):
        if pin_val == 10:  return "DQAR"
        if pin_val == 6:   return "DRYR"
    elif norm_pkg.startswith("DFN3030"):
        if pin_val == 8:   return "DRBR"
        if pin_val == 6:   return "DRSR"
        return None
    elif norm_pkg.startswith("DSBGA"):
        return "YZFR"
    elif norm_pkg.startswith("SC703") or norm_pkg.startswith("SC706"):
        return "DCKR"
    elif norm_pkg.startswith("SOD323"):
        return "DYFR"
    elif norm_pkg.startswith("SOD523"):
        return "DYAR"
    elif norm_pkg.startswith("SOT233"):
        return "DBZR"
    elif norm_pkg.startswith("SOT234"):
        return "DZDR"
    elif norm_pkg.startswith("SOT235") or norm_pkg.startswith("SOT236"):
        return "DBVR"
    elif re.match(r"SOT5[X\d]3", norm_pkg):
        return "DRLR"
    elif re.match(r"SOT9[X\d]3", norm_pkg):
        return "DRTR"
    elif norm_pkg.startswith("UQFN"):
        if pin_val == 10:  return "RSER"
        if pin_val == 14:  return "RVZR"
    elif norm_pkg.startswith("USON"):
        if pin_val == 6:
            if ti_gpn and str(ti_gpn).upper().replace("-Q1", "").strip() in TI_USON1616_PARTS:
                return "DPKR"
            return "DRYR"
        return None
    elif norm_pkg.startswith("WQFN"):
        return "RSFR"
    elif norm_pkg.startswith("WSON"):
        if pin_val == 15:  return "DSMR"
    return None


def _generate_ti_opn(ti_gpn, ti_package_str, ti_pin_str, competitor_package_alias, competitor_canonical_pkg=None):
    if not ti_gpn or ti_gpn == "-" or not ti_package_str or ti_package_str == "-":
        return ti_gpn

    # If no canonical pkg passed, try to derive one from the raw package string
    if not competitor_canonical_pkg and competitor_package_alias:
        pkg_str = competitor_package_alias if isinstance(competitor_package_alias, str) else ""
        for p in pkg_str.split(','):
            derived = _classify_digikey_package(normalize_package(p.strip()), "")
            if derived:
                competitor_canonical_pkg = derived
                break

    gpn_base = ti_gpn
    is_automotive = False
    if gpn_base.upper().endswith("-Q1"):
        gpn_base = gpn_base[:-3]
        is_automotive = True

    ti_packages_list = [p.strip() for p in str(ti_package_str).split(",") if p.strip()]
    try:
        ti_pins_list = [int(p.strip()) for p in str(ti_pin_str).split(",") if p.strip()]
    except (ValueError, TypeError):
        ti_pins_list = []

    # --- Canonical-code matching path (preferred) ---
    if competitor_canonical_pkg:
        for ti_pkg in ti_packages_list:
            # Never pair by list position — resolve this package's own pin count.
            pin_for_pkg = _resolve_pin_count(
                normalize_package(ti_pkg), ti_package_str, ti_pin_str, ti_gpn
            )

            ti_canonical = _classify_ti_package(ti_pkg, pin_for_pkg, ti_gpn)
            if ti_canonical and ti_canonical == competitor_canonical_pkg:
                suffix = _override_suffix(ti_gpn, ti_pkg) or CANONICAL_SUFFIX_MAP.get(ti_canonical)
                if suffix:
                    final_opn = gpn_base + suffix
                    if is_automotive:
                        return _automotive_opn(gpn_base, suffix)
                    return gpn_base + suffix

    # --- Legacy fallback path ---
    comp_pkg_set = set()
    if isinstance(competitor_package_alias, dict):
        comp_pkg_set = normalize_structured_package(competitor_package_alias)
    elif isinstance(competitor_package_alias, str) and competitor_package_alias:
        comp_pkg_set = {normalize_package(p) for p in competitor_package_alias.split(",") if p.strip()}

    matched_pkg = None
    for ti_pkg in ti_packages_list:
        if normalize_package(ti_pkg) in comp_pkg_set:
            matched_pkg = ti_pkg
            break

    if not matched_pkg:
        # Prefer a package in the same family as the competitor's, even if the
        # pin count differs; otherwise fall back to the row's first package.
        comp_family = (competitor_canonical_pkg or "").rsplit("_", 1)[0]
        if comp_family:
            for ti_pkg in ti_packages_list:
                np = normalize_package(ti_pkg)
                if np.startswith(comp_family) or comp_family.startswith(np):
                    matched_pkg = ti_pkg
                    break
        if not matched_pkg:
            matched_pkg = ti_packages_list[0] if ti_packages_list else None

    if not matched_pkg:
        return ti_gpn

    ordered = [matched_pkg] + [p for p in ti_packages_list if p != matched_pkg]
    suffix, norm_pkg = None, normalize_package(matched_pkg)
    for cand in ordered:
        cand_norm = normalize_package(cand)
        cand_pin = _resolve_pin_count(cand_norm, ti_package_str, ti_pin_str, ti_gpn)
        cand_suffix = _get_suffix(cand_norm, cand_pin, ti_gpn)
        if cand_suffix:
            suffix, norm_pkg = cand_suffix, cand_norm
            break

    if not suffix:
        # Last resort: any suffix this part's packages can legitimately carry.
        for cand in ordered:
            cn = normalize_package(cand)
            for pin_guess in (2, 3, 4, 5, 6, 8, 10, 12, 14, 15):
                if not _plausible_pin_for_family(cn, pin_guess):
                    continue
                s = _get_suffix(cn, pin_guess, ti_gpn)
                if s:
                    suffix = s
                    break
            if suffix:
                break

    if suffix:
        if is_automotive:
            return _automotive_opn(gpn_base, suffix)
        return gpn_base + suffix

    return ti_gpn

def _canonical_to_ti_pkg(canonical_code):
    """
    Derives a TI-compatible package string from a canonical code for use in
    find_package_matches. Strips the pin count suffix.
    e.g. 'DFN2510_10' -> 'DFN2510', 'SOT886_6' -> 'SOT886', 'SOT233_3' -> 'SOT-23-3'
    """
    if not canonical_code:
        return None
    family = canonical_code.rsplit('_', 1)[0]
    # Map back to TI's exact package name conventions where they differ
    mapping = {
        "SOT233": "SOT-23-3",
        "SOT234": "SOT-23-4",
        "SOT235": "SOT-23-5",
        "SOT236": "SOT-23-6",
        "SC703":  "SC70-3",
        "SC706":  "SC70-6",
        "SOD323": "SOD323",
        "SOD523": "SOD523",
        "SOT886": "SOT886",
        "SOT523": "SOT-5X3",
        "SOT553": "SOT-5X3",
        "SOT5X3": "SOT-5X3",
        "SOT9X3": "SOT-9X3",
    }
    return mapping.get(family, family)




def _canonical_to_ti_pkg_with_pins(canonical_code):
    """
    Like _canonical_to_ti_pkg but appends '-{pins}' so find_package_matches
    can extract the pin count for unstructured stage 2 filtering.
    e.g. 'DFN2510_10' -> 'DFN2510-10'
    (no suffix for packages where pin count is fixed/unambiguous)
    """
    if not canonical_code:
        return None
    parts = canonical_code.rsplit('_', 1)
    base = _canonical_to_ti_pkg(canonical_code)
    if len(parts) == 2:
        pin_count = parts[1]
        # Families where TI writes the pin count into the package name, so the
        # competitor string has to carry it too or stage 1 cannot match.
        # DFN1110 is fixed at 3 pins but TI still spells it "DFN1110-3".
        variable_pin_families = {"DFN2020", "DFN2510", "DFN0603", "DFN1006", "UQFN", "WSON", "X2SON", "USON", "USON1616", "SOT5X3", "SOT9X3", "DFN3030", "DFN1616", "DFN1610", "DFN1110"}
        family = parts[0]
        if family in variable_pin_families:
            return f"{base}-{pin_count}"
    return base


def _check_capacitance_rules(comp_cap, ti_cap, rule_set):
    if comp_cap is None or ti_cap is None or comp_cap == float('inf') or ti_cap == float('inf'):
        return True
    if rule_set == 'S':
        # Capacitance is usually THE selection parameter on a data-line array,
        # so an S has to stay proportional to what the designer specified
        # rather than sharing a fixed bucket with parts several times its size.
        # The ratio meets the old 10 pF -> 15 pF ceiling exactly, so nothing
        # above 10 pF changes.
        if comp_cap <= 10.0: return ti_cap <= comp_cap * CAP_S_MAX_RATIO
        return True
    elif rule_set == 'Q':
        if comp_cap <= 0.5: return ti_cap <= 1.0
        if comp_cap <= 1.0: return ti_cap <= 5.0
        if comp_cap <= 2.0: return ti_cap <= 10.0
        if comp_cap <= 10.0: return ti_cap <= 20.0
        return True
    return False


def _channel_count(value):
    """Channel count as an int, or None when the cell does not state one."""
    if value is None:
        return None
    text = str(value).strip()
    if text.lower() in ("", "-", "nan", "none"):
        return None
    match = re.search(r"\d+", text)
    if not match:
        return None
    return int(match.group(0)) or None


# How close two Vrwm figures must be to count as "exact" - float-noise slack
# only, in volts.
VRW_S_TOLERANCE_PERCENT = 10.0
# How much more capacitance a TI part may carry and still be a drop-in, as a
# multiple of the competitor's.
CAP_S_MAX_RATIO = 1.5

def _get_replacement_type(comp_specs, ti_alt_specs, is_zener, is_package_match):
    # No package match is what P means. Reaching this function at all means the
    # part already cleared the cross filters, so its specs are close enough.
    if not is_package_match:
        return 'P'

    if is_zener:
        comp_vz_str = comp_specs.get("Voltage - Reverse Standoff (Typ)")
        ti_vz_str = ti_alt_specs.get("Voltage - Reverse Standoff (Typ)")
        comp_tol_str = comp_specs.get("Tolerance", "inf")
        ti_tol_str = ti_alt_specs.get("Tolerance", "inf")
        comp_tol = to_numeric_val(comp_tol_str, default_if_error=float('inf'))
        ti_tol = to_numeric_val(ti_tol_str, default_if_error=float('inf'))

        comp_vz = to_numeric_val(comp_vz_str, default_if_error=None)
        ti_vz = to_numeric_val(ti_vz_str, default_if_error=None)

        vz_exact = (
            comp_vz is not None
            and ti_vz is not None
            and abs(comp_vz - ti_vz) <= 1e-6
        )

        if vz_exact and (ti_tol <= comp_tol + 0.1):
            return 'S'

        return 'Q'

    else:
        comp_vrw = comp_specs.get("Voltage - Reverse Standoff (Typ)")
        ti_vrw = ti_alt_specs.get("Voltage - Reverse Standoff (Typ)")
        comp_dir = comp_specs.get("Direction", "").lower()
        ti_dir = ti_alt_specs.get("Direction", "").lower()
        comp_cap = to_numeric_val(comp_specs.get("Capacitance"), float('inf'))
        ti_cap = to_numeric_val(ti_alt_specs.get("Capacitance"), float('inf'))
        comp_chan = _channel_count(comp_specs.get("Channels"))
        ti_chan = _channel_count(ti_alt_specs.get("Channels"))

        # S is decided by Vrwm, direction, channel count and capacitance. Vrwm
        # has to be exact -- a drop-in is the same standoff, not a near one.
        # ESD and surge are deliberately not part of it: both were tests a TI
        # part could fail by being BETTER than the competitor.
                # S is decided by Vrwm, direction, channel count and capacitance. Vrwm
        # only has to be close - a TI part within VRW_S_TOLERANCE_PERCENT of
        # the competitor's standoff still drops in. ESD and surge are
        # deliberately not part of it: both were tests a TI part could fail by
        # being BETTER than the competitor.
        comp_vrw_num = to_numeric_val(comp_vrw, default_if_error=None)
        ti_vrw_num = to_numeric_val(ti_vrw, default_if_error=None)

        s_vrw_ok = False
        if comp_vrw_num is not None and ti_vrw_num is not None:
            if comp_vrw_num > 0:
                s_vrw_ok = (abs(comp_vrw_num - ti_vrw_num) / comp_vrw_num * 100
                            <= VRW_S_TOLERANCE_PERCENT)
            else:
                s_vrw_ok = (comp_vrw_num == ti_vrw_num)
        s_dir_ok = (comp_dir == ti_dir)
        s_cap_ok = _check_capacitance_rules(comp_cap, ti_cap, 'S')
        s_chan_ok = (comp_chan is None or ti_chan is None or comp_chan == ti_chan)
        if s_vrw_ok and s_dir_ok and s_cap_ok and s_chan_ok:
            return 'S'

        # Package matches but at least one of those broke, so the cross is a Q.
        return 'Q'


LoadedData = namedtuple(
    "LoadedData",
    ["all_dfs", "ti_specs_df", "ti_zener_specs_df"],
)


def load_all_dfs(force_reload=False):
    """
    Loads every competitor / TI / DigiKey database and packs them in the
    order get_competitor_specs_leniently unpacks them.

    This is the ONLY place all_dfs is built. Everything that needs the
    databases -- main() below, categorize_crosses.load_context -- calls this
    instead of assembling its own tuple, so adding a competitor is a single
    edit here and no second copy can drift to the wrong length.
    """
    print("Loading databases (from cache if available)...")

    ti_specs_df = load_and_cache("ti_specs.pkl", TI_SPECS_DATABASE_FILE, force_reload,
        lambda: load_excel_data(TI_SPECS_DATABASE_FILE, header_row=10))

    ti_zener_specs_df = load_and_cache("ti_zener_specs.pkl", TI_ZENER_SPECS_DATABASE_FILE, force_reload,
        lambda: load_excel_data(TI_ZENER_SPECS_DATABASE_FILE, header_row=10))

    aos_specs_df = load_and_cache("aos_specs.pkl", AOS_SPECS_DATABASE_FILE, force_reload,
        lambda: load_excel_data(AOS_SPECS_DATABASE_FILE, header_row=6))

    diodes_dl_df = load_and_cache("diodes_dl.pkl", DIODES_DL_DATABASE_FILE, force_reload,
        lambda: load_excel_data(DIODES_DL_DATABASE_FILE, header_row=0))
    diodes_pl_df = load_and_cache("diodes_pl.pkl", DIODES_PL_DATABASE_FILE, force_reload,
        lambda: load_excel_data(DIODES_PL_DATABASE_FILE, header_row=0))
    diodes_zener_df = load_and_cache("diodes_zener.pkl", DIODES_ZENER_DATABASE_FILE, force_reload,
        lambda: load_excel_data(DIODES_ZENER_DATABASE_FILE, header_row=0))
    yenyo_esd_df = load_and_cache("yenyo_esd.pkl", YENYO_ESD_FILE, force_reload,
        lambda: load_excel_data(YENYO_ESD_FILE, header_row=0))
    yenyo_tvs_df = load_and_cache("yenyo_tvs.pkl", YENYO_TVS_FILE, force_reload,
        lambda: load_excel_data(YENYO_TVS_FILE, header_row=0))
    yenyo_zener_df = load_and_cache("yenyo_zener.pkl", YENYO_ZENER_FILE, force_reload,
        lambda: load_excel_data(YENYO_ZENER_FILE, header_row=0))
    nexperia_zener_df = load_and_cache("nexperia_zener.pkl", NEXPERIA_ZENER_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_ZENER_DATABASE_FILE, header_row=9))
    nexperia_esd_df = load_and_cache("nexperia_esd.pkl", NEXPERIA_ESD_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_ESD_DATABASE_FILE, header_row=9))
    nexperia_tvs_df = load_and_cache("nexperia_tvs.pkl", NEXPERIA_TVS_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_TVS_DATABASE_FILE, header_row=9))
    nexperia_auto_zener_df = load_and_cache("nexperia_auto_zener.pkl", NEXPERIA_AUTO_ZENER_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_AUTO_ZENER_DATABASE_FILE, header_row=9))
    nexperia_auto_esd_df = load_and_cache("nexperia_auto_esd.pkl", NEXPERIA_AUTO_ESD_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_AUTO_ESD_DATABASE_FILE, header_row=9))
    nexperia_auto_tvs_df = load_and_cache("nexperia_auto_tvs.pkl", NEXPERIA_AUTO_TVS_DATABASE_FILE, force_reload,
        lambda: load_excel_data(NEXPERIA_AUTO_TVS_DATABASE_FILE, header_row=9))
    littelfuse_tvs_array_df = load_and_cache("littelfuse_tvs_array.pkl", LITTELFUSE_TVS_ARRAY_FILE, force_reload,
        lambda: load_excel_data(LITTELFUSE_TVS_ARRAY_FILE))
    littelfuse_auto_tvs_df = load_and_cache("littelfuse_auto_tvs.pkl", LITTELFUSE_AUTO_TVS_FILE, force_reload,
        lambda: load_excel_data(LITTELFUSE_AUTO_TVS_FILE))
    littelfuse_tvs_df = load_and_cache("littelfuse_tvs.pkl", LITTELFUSE_TVS_FILE, force_reload,
        lambda: load_excel_data(LITTELFUSE_TVS_FILE))

    nichtek_esd_df = load_and_cache("nichtek_esd.pkl", NICHTEK_ESD_FILE, force_reload,
        lambda: load_excel_data(NICHTEK_ESD_FILE, header_row=0))
    nichtek_tvs_df = load_and_cache("nichtek_tvs.pkl", NICHTEK_TVS_FILE, force_reload,
        lambda: load_excel_data(NICHTEK_TVS_FILE, header_row=0))
    nichtek_zener_df = load_and_cache("nichtek_zener.pkl", NICHTEK_ZENER_FILE, force_reload,
        lambda: load_excel_data(NICHTEK_ZENER_FILE, header_row=0))

    mcc_esd_df = load_and_cache("mcc_esd.pkl", MCC_ESD_FILE, force_reload,
        lambda: load_excel_data(MCC_ESD_FILE, header_row=1))
    mcc_tvs_df = load_and_cache("mcc_tvs.pkl", MCC_TVS_FILE, force_reload,
        lambda: load_excel_data(MCC_TVS_FILE, header_row=1))
    mcc_zener_df = load_and_cache("mcc_zener.pkl", MCC_ZENER_FILE, force_reload,
        lambda: load_excel_data(MCC_ZENER_FILE, header_row=1))

    galaxy_esd_df = load_and_cache("galaxy_esd.pkl", GALAXY_ESD_FILE, force_reload,
        lambda: load_excel_data(GALAXY_ESD_FILE, header_row=0))
    galaxy_tvs_df = load_and_cache("galaxy_tvs.pkl", GALAXY_TVS_FILE, force_reload,
        lambda: load_excel_data(GALAXY_TVS_FILE, header_row=0))
    galaxy_zener_df = load_and_cache("galaxy_zener.pkl", GALAXY_ZENER_FILE, force_reload,
        lambda: load_excel_data(GALAXY_ZENER_FILE, header_row=0))

    leshan_esd_df = load_and_cache("leshan_esd.pkl", LESHAN_ESD_FILE, force_reload,
        lambda: load_excel_data(LESHAN_ESD_FILE, header_row=0))
    leshan_tvs_df = load_and_cache("leshan_tvs.pkl", LESHAN_TVS_FILE, force_reload,
        lambda: load_excel_data(LESHAN_TVS_FILE, header_row=0))
    leshan_zener_df = load_and_cache("leshan_zener.pkl", LESHAN_ZENER_FILE, force_reload,
        lambda: load_excel_data(LESHAN_ZENER_FILE, header_row=0))

    # INPAQ scraped sheets put their headers on row 1.
    inpaq_esd_df = load_and_cache(
        "inpaq_esd.pkl",
        INPAQ_ESD_FILE,
        force_reload,
        lambda: load_excel_data(INPAQ_ESD_FILE, header_row=0)
    )

    inpaq_auto_esd_df = load_and_cache(
        "inpaq_auto_esd.pkl",
        INPAQ_AUTO_ESD_FILE,
        force_reload,
        lambda: load_excel_data(INPAQ_AUTO_ESD_FILE, header_row=0)
    )

    inpaq_tvs_df = load_and_cache(
        "inpaq_tvs.pkl",
        INPAQ_TVS_FILE,
        force_reload,
        lambda: load_excel_data(INPAQ_TVS_FILE, header_row=0)
    )

    inpaq_auto_tvs_df = load_and_cache(
        "inpaq_auto_tvs.pkl",
        INPAQ_AUTO_TVS_FILE,
        force_reload,
        lambda: load_excel_data(INPAQ_AUTO_TVS_FILE, header_row=0)
    )

    inpaq_df_map = {
        "ESD": inpaq_esd_df,
        "Auto_ESD": inpaq_auto_esd_df,
        "TVS": inpaq_tvs_df,
        "Auto_TVS": inpaq_auto_tvs_df,
    }
    jiangsu_esd_df = load_and_cache(
        "jiangsu_esd.pkl",
        JIANGSU_ESD_FILE,
        force_reload,
        lambda: pd.read_excel(
            JIANGSU_ESD_FILE,
            header=0,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True},
        ),
    )

    jiangsu_tvs_df = load_and_cache(
        "jiangsu_tvs.pkl",
        JIANGSU_TVS_FILE,
        force_reload,
        lambda: pd.read_excel(
            JIANGSU_TVS_FILE,
            header=0,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True},
        ),
    )

    jiangsu_zener_df = load_and_cache(
        "jiangsu_zener.pkl",
        JIANGSU_ZENER_FILE,
        force_reload,
        lambda: pd.read_excel(
            JIANGSU_ZENER_FILE,
            header=0,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True},
        ),
    )

    semtech_specs_df = load_and_cache("semtech_specs.pkl", SEMTECH_SPECS_FILE, force_reload,
        lambda: load_excel_data(SEMTECH_SPECS_FILE))

    onsemi_esd_df = load_and_cache(
        "onsemi_esd.pkl",
        ONSEMI_ESD_FILE,
        force_reload,
        lambda: pd.read_csv(ONSEMI_ESD_FILE, encoding="utf-8-sig", low_memory=False)
    )

    onsemi_zener_df = load_and_cache(
        "onsemi_zener.pkl",
        ONSEMI_ZENER_FILE,
        force_reload,
        lambda: pd.read_csv(ONSEMI_ZENER_FILE, encoding="utf-8-sig", low_memory=False)
    )

    onsemi_df_map = {
        "ESD": onsemi_esd_df,
        "Zener": onsemi_zener_df,
    }

    # Vishay's ESD and Zener sheets put their headers on the first row. The TVS
    # sheet is left out on purpose - see VISHAY_TVS_FILE.
    vishay_esd_df = load_and_cache("vishay_esd.pkl", VISHAY_ESD_FILE, force_reload,
        lambda: load_excel_data(VISHAY_ESD_FILE, header_row=0))
    vishay_zener_df = load_and_cache("vishay_zener.pkl", VISHAY_ZENER_FILE, force_reload,
        lambda: load_excel_data(VISHAY_ZENER_FILE, header_row=0))
    yangjie_esd_df = load_and_cache("yangjie_esd.pkl", YANGJIE_ESD_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_ESD_FILE))
    yangjie_tvs_df = load_and_cache("yangjie_tvs.pkl", YANGJIE_TVS_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_TVS_FILE))
    yangjie_zener_df = load_and_cache("yangjie_zener.pkl", YANGJIE_ZENER_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_ZENER_FILE))
    yangjie_auto_esd_df = load_and_cache("yangjie_auto_esd.pkl", YANGJIE_AUTO_ESD_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_AUTO_ESD_FILE))
    yangjie_auto_tvs_df = load_and_cache("yangjie_auto_tvs.pkl", YANGJIE_AUTO_TVS_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_AUTO_TVS_FILE))
    yangjie_auto_zener_df = load_and_cache("yangjie_auto_zener.pkl", YANGJIE_AUTO_ZENER_FILE, force_reload,
        lambda: load_yangjie_xls(YANGJIE_AUTO_ZENER_FILE))


    jjm_zener_df = load_and_cache("jjm_zener.pkl", JJM_ZENER_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_ZENER_FILE, header_row=1))
    jjm_tvs_df = load_and_cache("jjm_tvs.pkl", JJM_TVS_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_TVS_FILE, header_row=1))
    jjm_esd_df = load_and_cache("jjm_esd.pkl", JJM_ESD_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_ESD_FILE, header_row=1))
    jjm_auto_zener_df = load_and_cache("jjm_auto_zener.pkl", JJM_AUTO_ZENER_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_AUTO_ZENER_FILE, header_row=1))
    jjm_auto_tvs_df = load_and_cache("jjm_auto_tvs.pkl", JJM_AUTO_TVS_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_AUTO_TVS_FILE, header_row=1))
    jjm_auto_esd_df = load_and_cache("jjm_auto_esd.pkl", JJM_AUTO_ESD_FILE, force_reload,
        lambda: load_jjm_xlsx(JJM_AUTO_ESD_FILE, header_row=1))
    
    stm_specs_df = load_and_cache("stm_specs.pkl", STM_SPECS_DATABASE_FILE, force_reload,
        lambda: pd.read_excel(STM_SPECS_DATABASE_FILE))
    if stm_specs_df.empty and os.path.exists(STM_SPECS_CSV_FILE):
        print(f"First-time setup: Found '{STM_SPECS_CSV_FILE}'. Converting...")
        if _unblock_csv_file(STM_SPECS_CSV_FILE):
            temp_df = pd.read_csv(STM_SPECS_CSV_FILE, on_bad_lines='skip', encoding='latin-1')
            temp_df.to_excel(STM_SPECS_DATABASE_FILE, index=False)
            os.remove(STM_SPECS_CSV_FILE)
            stm_specs_df = load_and_cache("stm_specs.pkl", STM_SPECS_DATABASE_FILE, True,
                lambda: pd.read_excel(STM_SPECS_DATABASE_FILE))

    
    panjit_esd_df = load_and_cache(
        "panjit_esd.pkl",
        PANJIT_ESD_FILE,
        force_reload,
        lambda: pd.read_excel(
            PANJIT_ESD_FILE,
            header=6,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True}
        )
    )

    panjit_tvs_df = load_and_cache(
        "panjit_tvs.pkl",
        PANJIT_TVS_FILE,
        force_reload,
        lambda: pd.read_excel(
            PANJIT_TVS_FILE,
            header=6,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True}
        )
    )

    panjit_zener_df = load_and_cache(
        "panjit_zener.pkl",
        PANJIT_ZENER_FILE,
        force_reload,
        lambda: pd.read_excel(
            PANJIT_ZENER_FILE,
            header=6,
            engine="xlrd",
            engine_kwargs={"ignore_workbook_corruption": True}
        )
    )
    panjit_df_map = {
        "ESD": panjit_esd_df,
        "TVS": panjit_tvs_df,
        "Zener": panjit_zener_df,
    }

    amazing_specs_df = load_and_cache("amazing_specs.pkl", AMAZING_SPECS_DATABASE_FILE, force_reload,
        lambda: pd.read_excel(AMAZING_SPECS_DATABASE_FILE, sheet_name='Amazing-Parametric', header=0, engine='openpyxl'))

    anbon_esd_df   = load_and_cache("anbon_esd.pkl",   ANBON_ESD_FILE,   force_reload, lambda: load_excel_data(ANBON_ESD_FILE))
    anbon_zener_df = load_and_cache("anbon_zener.pkl", ANBON_ZENER_FILE, force_reload, lambda: load_excel_data(ANBON_ZENER_FILE))
    anbon_tvs_df   = load_and_cache("anbon_tvs.pkl",   ANBON_TVS_FILE,   force_reload, lambda: load_excel_data(ANBON_TVS_FILE))

    anbon_esd_df   = load_and_cache("anbon_esd.pkl",   ANBON_ESD_FILE,   force_reload, lambda: load_excel_data(ANBON_ESD_FILE))
    anbon_zener_df = load_and_cache("anbon_zener.pkl", ANBON_ZENER_FILE, force_reload, lambda: load_excel_data(ANBON_ZENER_FILE))
    anbon_tvs_df   = load_and_cache("anbon_tvs.pkl",   ANBON_TVS_FILE,   force_reload, lambda: load_excel_data(ANBON_TVS_FILE))

    central_zener_df = load_and_cache("central_zener.pkl", CENTRAL_ZENER_FILE, force_reload,
        lambda: load_excel_data(CENTRAL_ZENER_FILE))

    # Comchip sheets put their headers on the first row.
    comchip_esd_df   = load_and_cache("comchip_esd.pkl",   COMCHIP_ESD_FILE,   force_reload,
        lambda: load_excel_data(COMCHIP_ESD_FILE, header_row=0))
    comchip_tvs_df   = load_and_cache("comchip_tvs.pkl",   COMCHIP_TVS_FILE,   force_reload,
        lambda: load_excel_data(COMCHIP_TVS_FILE, header_row=0))
    comchip_zener_df = load_and_cache("comchip_zener.pkl", COMCHIP_ZENER_FILE, force_reload,
        lambda: load_excel_data(COMCHIP_ZENER_FILE, header_row=0))

    # Diotec sheets put their headers on the first row.
    diotec_esd_df   = load_and_cache("diotec_esd.pkl",   DIOTEC_ESD_FILE,   force_reload,
        lambda: load_excel_data(DIOTEC_ESD_FILE, header_row=0))
    diotec_tvs_df   = load_and_cache("diotec_tvs.pkl",   DIOTEC_TVS_FILE,   force_reload,
        lambda: load_excel_data(DIOTEC_TVS_FILE, header_row=0))
    diotec_zener_df = load_and_cache("diotec_zener.pkl", DIOTEC_ZENER_FILE, force_reload,
        lambda: load_excel_data(DIOTEC_ZENER_FILE, header_row=0))

    # Eaton's parametric export puts its headers on the first row.
    eaton_specs_df = load_and_cache("eaton_specs.pkl", EATON_SPECS_FILE, force_reload,
        lambda: load_excel_data(EATON_SPECS_FILE, header_row=0))
    if eaton_specs_df.empty and os.path.exists(EATON_SPECS_CSV_FILE):
        print(f"First-time setup: Found '{EATON_SPECS_CSV_FILE}'. Converting...")
        if _unblock_csv_file(EATON_SPECS_CSV_FILE):
            temp_df = pd.read_csv(EATON_SPECS_CSV_FILE, on_bad_lines='skip', encoding='latin-1')
            temp_df.to_excel(EATON_SPECS_FILE, index=False)
            os.remove(EATON_SPECS_CSV_FILE)
            eaton_specs_df = load_and_cache("eaton_specs.pkl", EATON_SPECS_FILE, True,
                lambda: load_excel_data(EATON_SPECS_FILE, header_row=0))

    # EIC's scraped sheets put their headers on the first row.
    eic_tvs_df   = load_and_cache("eic_tvs.pkl",   EIC_TVS_FILE,   force_reload,
        lambda: load_excel_data(EIC_TVS_FILE, header_row=0))
    eic_zener_df = load_and_cache("eic_zener.pkl", EIC_ZENER_FILE, force_reload,
        lambda: load_excel_data(EIC_ZENER_FILE, header_row=0))

    # EVVO's scraped sheets put their headers on the first row.
    evvo_tvs_df   = load_and_cache("evvo_tvs.pkl",   EVVO_TVS_FILE,   force_reload,
        lambda: load_excel_data(EVVO_TVS_FILE, header_row=0))
    evvo_zener_df = load_and_cache("evvo_zener.pkl", EVVO_ZENER_FILE, force_reload,
        lambda: load_excel_data(EVVO_ZENER_FILE, header_row=0))

    # Good-Ark's scraped sheets put their headers on the first row.
    goodark_esd_df   = load_and_cache("goodark_esd.pkl",   GOODARK_ESD_FILE,   force_reload,
        lambda: load_excel_data(GOODARK_ESD_FILE, header_row=0))
    goodark_tvs_df   = load_and_cache("goodark_tvs.pkl",   GOODARK_TVS_FILE,   force_reload,
        lambda: load_excel_data(GOODARK_TVS_FILE, header_row=0))
    goodark_zener_df = load_and_cache("goodark_zener.pkl", GOODARK_ZENER_FILE, force_reload,
        lambda: load_excel_data(GOODARK_ZENER_FILE, header_row=0))

    # Infineon parametric export: downloaded by hand, headers on the first row.
    infineon_specs_df = load_and_cache("infineon_specs.pkl", INFINEON_SPECS_FILE, force_reload,
        lambda: load_excel_data(INFINEON_SPECS_FILE, header_row=0))

    # Per-competitor DigiKey exports: channel counts keyed on 'Mfr Part #',
    # split across a bidirectional and a unidirectional column.
    evvo_digi_df = load_digi_export(EVVO_DIGI_FILE, EVVO_DIGI_CSV,
        "evvo_digi.pkl", force_reload, "EVVO")
    goodark_digi_df = load_digi_export(GOODARK_DIGI_FILE, GOODARK_DIGI_CSV,
        "goodark_digi.pkl", force_reload, "Good-Ark")
    amazing_digi_df = load_digi_export(AMAZING_DIGI_FILE, AMAZING_DIGI_CSV,
        "amazing_digi.pkl", force_reload, "Amazing")
    littelfuse_digi_df = load_digi_export(LITTELFUSE_DIGI_FILE, LITTELFUSE_DIGI_CSV,
        "littelfuse_digi.pkl", force_reload, "Littelfuse")

    # Comchip DigiKey export: channel counts keyed on 'Mfr Part #'.
    comchip_digi_df = pd.DataFrame()
    if os.path.exists(COMCHIP_DIGI_CSV):
        try:
            comchip_digi_df = pd.read_csv(COMCHIP_DIGI_CSV, on_bad_lines='skip', encoding='utf-8-sig', low_memory=False)
            print(f"Loaded {len(comchip_digi_df)} rows from '{COMCHIP_DIGI_CSV}'.")
        except Exception as e:
            print(f"Warning: Could not load '{COMCHIP_DIGI_CSV}'. Error: {e}")
    else:
        print(f"Warning: '{COMCHIP_DIGI_CSV}' not found. Comchip parts will be treated as single channel.")

    # Pack competitor DFs for dispatcher
    all_dfs = (
        aos_specs_df, amazing_specs_df,
        {"Data Line": diodes_dl_df, "Power Line": diodes_pl_df, "Zener": diodes_zener_df},
        {"Zener": nexperia_zener_df, "ESD": nexperia_esd_df, "TVS": nexperia_tvs_df,
         "Auto_Zener": nexperia_auto_zener_df, "Auto_ESD": nexperia_auto_esd_df, "Auto_TVS": nexperia_auto_tvs_df},
        {"Zener": jjm_zener_df, "TVS": jjm_tvs_df, "ESD": jjm_esd_df,
         "Auto_Zener": jjm_auto_zener_df, "Auto_TVS": jjm_auto_tvs_df, "Auto_ESD": jjm_auto_esd_df},
        {"TVS": littelfuse_tvs_df, "TVS_Array": littelfuse_tvs_array_df, "Auto_TVS": littelfuse_auto_tvs_df},
        littelfuse_digi_df,
        {"ESD": jiangsu_esd_df, "TVS": jiangsu_tvs_df, "Zener": jiangsu_zener_df},
        {"ESD": vishay_esd_df, "Zener": vishay_zener_df},
        {"ESD": yangjie_esd_df, "TVS": yangjie_tvs_df, "Zener": yangjie_zener_df,
         "Auto_ESD": yangjie_auto_esd_df, "Auto_TVS": yangjie_auto_tvs_df, "Auto_Zener": yangjie_auto_zener_df},
        {"ESD": mcc_esd_df, "TVS": mcc_tvs_df, "Zener": mcc_zener_df},
        {"ESD": galaxy_esd_df, "TVS": galaxy_tvs_df, "Zener": galaxy_zener_df},
        {"ESD": leshan_esd_df, "TVS": leshan_tvs_df, "Zener": leshan_zener_df},
        {"ESD": nichtek_esd_df, "TVS": nichtek_tvs_df, "Zener": nichtek_zener_df},
        {"ESD": yenyo_esd_df, "TVS": yenyo_tvs_df, "Zener": yenyo_zener_df},
        inpaq_df_map,
        semtech_specs_df, stm_specs_df, panjit_df_map, onsemi_df_map,
        anbon_esd_df, anbon_zener_df, anbon_tvs_df, central_zener_df,
        comchip_esd_df, comchip_tvs_df, comchip_zener_df, comchip_digi_df,
        {"ESD": diotec_esd_df, "TVS": diotec_tvs_df, "Zener": diotec_zener_df},
        eaton_specs_df,
        {"TVS": eic_tvs_df, "Zener": eic_zener_df},
        {"TVS": evvo_tvs_df, "Zener": evvo_zener_df},
        evvo_digi_df,
        {"ESD": goodark_esd_df, "TVS": goodark_tvs_df, "Zener": goodark_zener_df},
        goodark_digi_df, amazing_digi_df,
        infineon_specs_df,
        ti_specs_df, ti_zener_specs_df
    )

    print("All databases loaded and ready.")

    return LoadedData(all_dfs, ti_specs_df, ti_zener_specs_df)


def main():
    start_time = time.time()

    reload_input = input("Type 'reload' to refresh TI data files, or press Enter to continue: ").strip().lower()
    force_reload = (reload_input == 'reload')
    manage_data_files(force_reload=force_reload)

    loaded = load_all_dfs(force_reload)
    all_dfs = loaded.all_dfs
    ti_specs_df = loaded.ti_specs_df
    ti_zener_specs_df = loaded.ti_zener_specs_df

    mode = input("Select mode: (1) Single Part Cross, (2) Batch File Cross, (3) Specs: ").strip()

    if mode == '3':
        comp_part_input = input("Enter the competitor's diode part number: ").strip()
        if not comp_part_input:
            print("No competitor part number entered.")
            return

        competitor_name = input("Enter the competitor's name: ").strip().lower()
        comp_specs = get_competitor_specs_leniently(comp_part_input, competitor_name, all_dfs)

        if not comp_specs:
            print(f"\nCould not find specifications for '{comp_part_input}'.")
            return

        is_zener_part = "zener" in comp_specs.get("Source File", "").lower()

        param_keys_tvs = [
            ("Device Name", "Device Name"), ("Direction", "Direction"), ("Package", "Package"),
            ("Voltage - Reverse Standoff (Typ)", "Vrw Standoff (Typ)"),
            ("Voltage - Clamping (Max) @ Ipp", "Vcl @ Ipp (Max)"),
            ("Capacitance", "Capacitance"), ("Channels", "Channels"),
            ("IEC 61000-4-5", "IEC 61000-4-5 (Surge)"), ("IEC 61000-4-2", "IEC 61000-4-2 (ESD)"),
            ("Power Dissipation (Pd)", "Peak Pulse Power"),
            ("Price ($/ku)", "Price ($/ku)")
        ]
        param_keys_zener = [
            ("Device Name", "Device Name"), ("Package", "Package"),
            ("Voltage - Reverse Standoff (Typ)", "Zener Voltage (Vz)"),
            ("Tolerance", "Tolerance"), ("Power Dissipation (Pd)", "Power (Pd)"),
            ("Price ($/ku)", "Price ($/ku)")
        ]

        display_keys = param_keys_zener if is_zener_part else param_keys_tvs
        table_data = [[display_name, comp_specs.get(key, "-")] for key, display_name in display_keys]

        print(f"\n--- Specifications for {comp_specs.get('Device Name', comp_part_input).upper()} ---")
        print(tabulate(table_data, headers=["Parameter", "Value"], tablefmt="grid"))

        end_time = time.time()
        print(f"\nTotal execution time: {end_time - start_time:.2f} seconds")
        return

    if mode == '2':
        is_gui_mode = "RUNNING_IN_GUI" in os.environ

        if is_gui_mode:
            print("Backend: GUI mode detected, reading file paths from input stream.")
            input_excel_path = input().strip()
            output_excel_path = input().strip()
        else:
            print("Backend: Command-line mode, opening file dialogs.")
            from tkinter import Tk, filedialog
            root = Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            input_excel_path = filedialog.askopenfilename(
                title="Select the Competitor Parts Excel File",
                filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
            if not input_excel_path:
                print("No input file selected. Aborting.")
                return
            output_excel_path = filedialog.asksaveasfilename(
                title="Save Output As...",
                filetypes=[("Excel Files", "*.xlsx")],
                defaultextension=".xlsx",
                initialfile="ti_cross_results.xlsx"
            )
            if not output_excel_path:
                print("No output file location selected. Aborting.")
                return

        try:
            batch_df = pd.read_excel(input_excel_path)
            if "Competitor Parts" not in batch_df.columns or "Competitor Name" not in batch_df.columns:
                print(f"ERROR: Input file must have 'Competitor Parts' AND 'Competitor Name' columns.")
                return
        except FileNotFoundError:
            print(f"ERROR: Input file not found at '{input_excel_path}'")
            return
        except Exception as e:
            print(f"ERROR: Could not read the Excel file. {e}")
            return

        all_alternatives_padded = []

        try:
            from tqdm import tqdm
            iterator = tqdm(batch_df.iterrows(), desc="Processing parts", unit="part", file=sys.stdout, total=len(batch_df))
        except ImportError:
            print("Processing parts... (for a progress bar, run: pip install tqdm)")
            iterator = batch_df.iterrows()

        for index, row in iterator:
            part_str = str(row["Competitor Parts"]).strip()
            competitor_name_batch = str(row["Competitor Name"]).strip().lower()

            if not part_str or pd.isna(part_str) or part_str.lower() == 'nan' or not competitor_name_batch:
                all_alternatives_padded.append(["-", "-", "-"])
                continue

            comp_specs = get_competitor_specs_leniently(part_str, competitor_name_batch, all_dfs)

            if not comp_specs:
                all_alternatives_padded.append(["Specs Not Found", "-", "-"])
                continue

            is_zener_part_batch = "zener" in comp_specs.get("Source File", "").lower()
            original_display_package = comp_specs.get("Package", "-")

            if is_zener_part_batch:
                comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
                ti_alternatives_list = find_ti_zener_alternatives(comp_specs, ti_zener_specs_df)
            else:
                comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
                ti_alternatives_list = find_ti_alternatives(comp_specs, ti_specs_df)

            generated_opns = []
            if ti_alternatives_list:
                for alt_dict in ti_alternatives_list:
                    gpn = alt_dict['part_number']
                    if is_zener_part_batch:
                        ti_alt_specs = fetch_ti_zener_specs_from_excel(gpn, ti_zener_specs_df, [])
                    else:
                        ti_alt_specs = fetch_ti_specs_from_excel(gpn, ti_specs_df, [])

                    if ti_alt_specs:
                        ti_package = ti_alt_specs.get("Package", "-")
                        competitor_pkg_alias = comp_specs.get("Package", "")
                        src_df = ti_zener_specs_df if is_zener_part_batch else ti_specs_df
                        ti_pin_str = "-"
                        if src_df is not None and not src_df.empty:
                            pn_col = next((c for c in src_df.columns if "product or part number" in c.lower()), None)
                            pin_col = next((c for c in src_df.columns if "pin count" in c.lower()), None)
                            if pn_col and pin_col:
                                match = src_df[src_df[pn_col].astype(str).str.strip() == gpn.strip()]
                                if not match.empty:
                                    ti_pin_str = str(match.iloc[0][pin_col])
                        opn = _generate_ti_opn(gpn, ti_package, ti_pin_str, competitor_pkg_alias, comp_specs.get("Canonical Package"))
                        generated_opns.append(opn)
                    else:
                        generated_opns.append(gpn)

            padded_list = (generated_opns + ["-", "-", "-"])[:3]
            all_alternatives_padded.append(padded_list)

        batch_df["Competitor Name"] = batch_df["Competitor Name"].apply(canonical_competitor_name)
        batch_df["TI Alternate 1"] = [alt[0] for alt in all_alternatives_padded]
        batch_df["TI Alternate 2"] = [alt[1] for alt in all_alternatives_padded]
        batch_df["TI Alternate 3"] = [alt[2] for alt in all_alternatives_padded]

        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                batch_df.to_excel(writer, index=False, sheet_name='Results')
                worksheet = writer.sheets['Results']
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value)) for cell in column_cells)
                    column_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[column_letter].width = length + 2
            print(f"\nSuccessfully processed batch file. Output saved to '{output_excel_path}'")
        except Exception as e:
            print(f"\nERROR: Could not save the output Excel file. {e}")

        end_time = time.time()
        print(f"\nTotal execution time: {end_time - start_time:.2f} seconds")
        return

    elif mode == '4':
        is_gui_mode = "RUNNING_IN_GUI" in os.environ

        if is_gui_mode:
            print("Backend: GUI mode detected, reading file paths from input stream.")
            input_excel_path = input().strip()
            output_excel_path = input().strip()
        else:
            print("Backend: Command-line mode, opening file dialogs.")
            root = Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            input_excel_path = filedialog.askopenfilename(
                title="Select the Competitor Parts Excel File",
                filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
            if not input_excel_path:
                print("No input file selected. Aborting.")
                return
            output_excel_path = filedialog.asksaveasfilename(
                title="Save CrossRef Output As...",
                filetypes=[("Excel Files", "*.xlsx")],
                defaultextension=".xlsx",
                initialfile="ti_crossref_results.xlsx"
            )
            if not output_excel_path:
                print("No output file location selected. Aborting.")
                return

        try:
            batch_df = pd.read_excel(input_excel_path)
            if "Competitor Parts" not in batch_df.columns or "Competitor Name" not in batch_df.columns:
                print(f"ERROR: Input file must have 'Competitor Parts' AND 'Competitor Name' columns.")
                return
        except FileNotFoundError:
            print(f"ERROR: Input file not found at '{input_excel_path}'")
            return
        except Exception as e:
            print(f"ERROR: Could not read the Excel file. {e}")
            return

        crossref_results = []

        try:
            from tqdm import tqdm
            iterator = tqdm(batch_df.iterrows(), desc="Processing parts for CrossRef", unit="part", file=sys.stdout, total=len(batch_df))
        except ImportError:
            print("Processing parts for CrossRef... (for a progress bar, run: pip install tqdm)")
            iterator = batch_df.iterrows()

        for index, row in iterator:
            part_str = str(row["Competitor Parts"]).strip()
            competitor_name = str(row["Competitor Name"]).strip().lower()

            ti_opn = "-"
            top_ti_alt_gpn = "-"
            replacement_type = "P"
            is_competitor_zener = False

            if not part_str or pd.isna(part_str) or part_str.lower() == 'nan' or not competitor_name:
                pass
            else:
                comp_specs = get_competitor_specs_leniently(part_str, competitor_name, all_dfs)

                if comp_specs:
                    is_competitor_zener = "zener" in comp_specs.get("Source File", "").lower()
                    original_display_package = comp_specs.get("Package", "-")

                    ti_alternatives_list = []
                    if is_competitor_zener:
                        comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
                        ti_alternatives_list = find_ti_zener_alternatives(comp_specs, ti_zener_specs_df)
                    else:
                        comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
                        ti_alternatives_list = find_ti_alternatives(comp_specs, ti_specs_df)

                    if ti_alternatives_list:
                        top_alt_dict = ti_alternatives_list[0]
                        top_ti_alt_gpn = top_alt_dict['part_number']
                        is_pkg_match = top_alt_dict.get('is_package_match', False)

                        if is_competitor_zener:
                            ti_alt_full_specs = fetch_ti_zener_specs_from_excel(top_ti_alt_gpn, ti_zener_specs_df, [])
                        else:
                            ti_alt_full_specs = fetch_ti_specs_from_excel(top_ti_alt_gpn, ti_specs_df, [])

                        if ti_alt_full_specs:
                            replacement_type = _get_replacement_type(comp_specs, ti_alt_full_specs, is_competitor_zener, is_pkg_match)
                            ti_package = ti_alt_full_specs.get("Package", "-")
                            competitor_pkg_alias = comp_specs.get("Package", "")
                            src_df = ti_zener_specs_df if is_competitor_zener else ti_specs_df
                            ti_pin_str = "-"
                            if src_df is not None and not src_df.empty:
                                pn_col = next((c for c in src_df.columns if "product or part number" in c.lower()), None)
                                pin_col = next((c for c in src_df.columns if "pin count" in c.lower()), None)
                                if pn_col and pin_col:
                                    match = src_df[src_df[pn_col].astype(str).str.strip() == top_ti_alt_gpn.strip()]
                                    if not match.empty:
                                        ti_pin_str = str(match.iloc[0][pin_col])
                            ti_opn = _generate_ti_opn(top_ti_alt_gpn, ti_package, ti_pin_str, competitor_pkg_alias, comp_specs.get("Canonical Package"))
                        else:
                            ti_opn = top_ti_alt_gpn

            crossref_results.append({
                "COMPETITOR_NAME": competitor_name.upper(),
                "COMP_GENERIC_PART_NUMBER": part_str,
                "COMP_ORDERABLE_PART_NUMBER": part_str,
                "TI_GPN": top_ti_alt_gpn,
                "GPN_REPLACEMENT_TYPE": replacement_type,
                "TI_OPN": ti_opn,
                "OPN_REPLACEMENT_TYPE": "P"
            })

        output_df = pd.DataFrame(crossref_results)

        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='CrossRef_Results')
                worksheet = writer.sheets['CrossRef_Results']
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value)) for cell in column_cells)
                    column_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[column_letter].width = length + 2
            print(f"\nSuccessfully processed batch file. Output saved to '{output_excel_path}'")
        except Exception as e:
            print(f"\nERROR: Could not save the output Excel file. {e}")

        end_time = time.time()
        print(f"\nTotal execution time: {end_time - start_time:.2f} seconds")
        return

    if mode == '1':
        comp_part_input = input("Enter the competitor's diode part number: ").strip()
    else:
        print("Invalid mode selected. Please enter '1', '2', '3', or '4'.")
        return

    if not comp_part_input:
        print("No competitor part number entered.")
        return

    competitor_name = input("Enter the competitor's name (for display only): ").strip().lower()
    comp_specs = get_competitor_specs_leniently(comp_part_input, competitor_name, all_dfs)

    if not comp_specs:
        print(f"\nCould not find specifications for '{comp_part_input}'. Aborting.")
        return

    if is_automotive_grade(comp_specs.get("Grade", "")):
        print("--> Competitor part identified as Automotive Grade.")
    else:
        print("--> Competitor part identified as Commercial/Standard Grade.")

    is_zener_part = "zener" in comp_specs.get("Source File", "").lower()

    if is_zener_part:
        print("\nCompetitor identified as a Zener Diode. Starting Zener cross-reference...")
        param_keys_display_order = [
            ("Device Name", "Device Name"),
            ("Package", "Package"),
            ("Voltage - Reverse Standoff (Typ)", "Zener Voltage (Vz)"),
            ("Tolerance", "Tolerance"),
            ("Power Dissipation (Pd)", "Power (Pd)"),
            ("Price ($/ku)", "Price ($/ku)"),
        ]
        original_display_package = comp_specs.get("Package", "-")
        comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
        print(f"\nOriginal Pkg: '{original_display_package}' -> Normalized: '{comp_specs['Package']}'")
        ti_alternative_opns = find_ti_zener_alternatives(comp_specs, ti_zener_specs_df)
        comp_specs["Package"] = _canonical_to_ti_pkg(comp_specs.get("Canonical Package")) or original_display_package
    else:
        print("\nCompetitor identified as a standard Diode/TVS. Starting standard cross-reference...")
        param_keys_display_order = [
            ("Device Name", "Device Name"),
            ("Direction", "Direction"),
            ("Package", "Package"),
            ("Voltage - Reverse Standoff (Typ)", "Vrw Standoff (Typ)"),
            ("Voltage - Clamping (Max) @ Ipp", "Vcl @ Ipp (Max)"),
            ("Capacitance", "Capacitance"),
            ("Channels", "Channels"),
            ("IEC 61000-4-5", "IEC 61000-4-5 (Surge/EFT)"),
            ("IEC 61000-4-2", "IEC 61000-4-2 (ESD)"),
            ("Power Dissipation (Pd)", "Peak Pulse Power"),
            ("Price ($/ku)", "Price ($/ku)")
        ]
        original_display_package = comp_specs.get("Package", "-")
        comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")) or normalize_package(original_display_package)
        print(f"\nOriginal Pkg: '{original_display_package}' -> Normalized: '{comp_specs['Package']}'")
        ti_alternative_opns = find_ti_alternatives(comp_specs, ti_specs_df)
        comp_specs["Package"] = _canonical_to_ti_pkg(comp_specs.get("Canonical Package")) or original_display_package

    ti_alternatives_specs_list = []
    ti_alternative_codes = []
    if ti_alternative_opns:
        print(f"\nFetching TI specs for top {len(ti_alternative_opns)} alternatives...")
        for alt_dict in ti_alternative_opns:
            part_number_str = alt_dict['part_number']
            if is_zener_part:
                ti_alt_specs = fetch_ti_zener_specs_from_excel(part_number_str, ti_zener_specs_df, param_keys_display_order)
            else:
                if ti_specs_df is not None:
                    ti_alt_specs = fetch_ti_specs_from_excel(part_number_str, ti_specs_df, param_keys_display_order)
                else:
                    ti_alt_specs = None
            if ti_alt_specs:
                src_df = ti_zener_specs_df if is_zener_part else ti_specs_df
                ti_pin_str = "-"
                if src_df is not None and not src_df.empty:
                    pn_col = next((c for c in src_df.columns if "product or part number" in c.lower()), None)
                    pin_col = next((c for c in src_df.columns if "pin count" in c.lower()), None)
                    ppp_col = next((c for c in src_df.columns if "peak pulse" in c.lower() or "pd (max)" in c.lower()), None)
                    if pn_col:
                        match = src_df[src_df[pn_col].astype(str).str.strip() == part_number_str.strip()]
                        if not match.empty:
                            if pin_col:
                                ti_pin_str = str(match.iloc[0][pin_col])
                            if ppp_col:
                                ppp_val = str(match.iloc[0][ppp_col])
                                ppp_clean = ppp_val.strip() if ppp_val not in ("-", "nan", "") else "-"
                                if ppp_clean != "-":
                                    # Ensure W unit
                                    if not ppp_clean.lower().endswith('w'):
                                        ppp_clean = ppp_clean + " W"
                                ti_alt_specs["Power Dissipation (Pd)"] = ppp_clean
                opn = _generate_ti_opn(part_number_str, ti_alt_specs.get("Package", "-"), ti_pin_str, comp_specs.get("Package", ""), comp_specs.get("Canonical Package"))
                ti_alt_specs["OPN"] = opn
                ti_alternatives_specs_list.append(ti_alt_specs)
                # S/Q/P for this pairing, so a single-part cross shows the same
                # grade the batch export would put in replacementCodeOPN.
                ti_alternative_codes.append(
                    _get_replacement_type(comp_specs, ti_alt_specs, is_zener_part,
                                          alt_dict.get("is_package_match", False)) or "-"
                )

    # An S cross beats a higher-scoring Q or P. The list is already in score
    # order, so the first S in it is the best-scoring one; move it to the front
    # so the primary column of the table is the strongest cross available.
    if ti_alternative_codes and ti_alternative_codes[0] != "S" and "S" in ti_alternative_codes:
        s_index = ti_alternative_codes.index("S")
        promoted_name = ti_alternatives_specs_list[s_index].get(
            "OPN", ti_alternatives_specs_list[s_index].get("Device Name", "?"))
        displaced_code = ti_alternative_codes[0]
        for sequence in (ti_alternatives_specs_list, ti_alternative_codes):
            sequence.insert(0, sequence.pop(s_index))
        print(f"--> Promoted '{promoted_name}' to TI Alt 1: it is an S cross, "
              f"the higher-scored part was only a {displaced_code}.")

    if "Direction" not in comp_specs:
        comp_specs["Direction"] = "-"

    print("\n--- Competitor Specs ---")
    for key, value in comp_specs.items():
        if key in [k[0] for k in param_keys_display_order]:
            print(f"  {key}: {value}")
    print("---------------------------------------------")

    if comp_specs.get("Package", "").endswith(", No"):
        comp_specs["Package"] = comp_specs["Package"][:-4]

    comp_canonical = comp_specs.get("Canonical Package")
    if not comp_canonical:
        # Try raw package string first (before normalization strips hyphens/pin markers)
        raw_pkg = comp_specs.get("Package", "")
        for p in str(raw_pkg).split(','):
            derived = _classify_digikey_package(p.strip(), "")
            if derived:
                comp_canonical = derived
                break
    comp_pkg_norm = normalize_package(comp_specs.get("Package", ""))

    for ti_spec in ti_alternatives_specs_list:
        original_ti_pkg_str = ti_spec.get("Package", "-")
        if original_ti_pkg_str == "-":
            continue
        package_list = [p.strip() for p in original_ti_pkg_str.split(',') if p.strip()]
        if not package_list:
            continue

        # Get pin counts for this TI part to classify packages properly
        ti_pin_str = ti_spec.get("Pin count", "-")
        try:
            ti_pins_list = [int(p.strip()) for p in str(ti_pin_str).split(',') if p.strip()]
        except (ValueError, TypeError):
            ti_pins_list = []

        chosen_pkg = None

        # Prefer package whose canonical code matches competitor's
        if comp_canonical:
            for i, pkg in enumerate(package_list):
                pin_for_pkg = ti_pins_list[0] if len(ti_pins_list) == 1 else (ti_pins_list[i] if i < len(ti_pins_list) else None)
                if _classify_ti_package(pkg, pin_for_pkg, ti_spec.get("Device Name", "")) == comp_canonical:
                    chosen_pkg = pkg
                    break

        # Fall back: match by normalized name
        if not chosen_pkg:
            for pkg in package_list:
                if normalize_package(pkg) == comp_pkg_norm:
                    chosen_pkg = pkg
                    break

        # No match — use first package
        if not chosen_pkg:
            chosen_pkg = package_list[0]

        normalized_chosen = normalize_package(chosen_pkg)
        ti_spec["Package"] = PACKAGE_DISPLAY_MAP.get(normalized_chosen, chosen_pkg)

    comp_name_header_val = comp_specs.get('Device Name', comp_part_input)
    headers = ["Parameter", f"Competitor ({comp_name_header_val.upper()})"]

    for i, ti_alt_specs in enumerate(ti_alternatives_specs_list):
        ti_alt_name = ti_alt_specs.get("OPN", ti_alt_specs.get("Device Name", f"Alt {i+1}"))
        headers.append(f"TI Alt {i+1} ({ti_alt_name})")

    for i in range(len(ti_alternatives_specs_list), 3):
        headers.append(f"TI Alt {i+1} (N/A)")

    table_data = []
    for key_name, display_name in param_keys_display_order:
        comp_value = comp_specs.get(key_name, "-")
        if key_name == "Device Name":
            comp_value = str(comp_value).upper()
        row_data = [display_name, comp_value]
        for ti_alt_specs in ti_alternatives_specs_list:
            if key_name == "Device Name":
                row_data.append(ti_alt_specs.get("OPN", ti_alt_specs.get(key_name, "-")))
            else:
                row_data.append(ti_alt_specs.get(key_name, "-"))
        for _ in range(len(ti_alternatives_specs_list), 3):
            row_data.append("-")
        table_data.append(row_data)

    # Cross grade per alternative, on its own row at the bottom of the table.
    if ti_alternative_codes:
        code_row = ["Replacement Code", "-"] + ti_alternative_codes
        code_row += ["-"] * (3 - len(ti_alternative_codes))
        table_data.append(code_row)

    if not ti_alternative_opns:
        print("\nNo suitable TI alternatives found based on the criteria.")
    else:
        print("\n--- Comparison Table ---")
        print(tabulate(table_data, headers=headers, tablefmt="grid"))

    end_time = time.time()
    print(f"\nTotal execution time: {end_time - start_time:.2f} seconds")

if __name__ == "__main__":
    main()