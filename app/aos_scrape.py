# --- START OF FILE aos_scrape.py ---

import re
import pandas as pd
from parsing import safe_strip


def _flat(text):
    """Header comparison form: lowercase, alphanumerics only.

    Strips the separators that vary between exports, including the literal
    '<br>' that shows up inside 'AEC-Q101<br>Qualified'.
    """
    return re.sub(r'[^a-z0-9]+', '', str(text).lower())


def _find_cols(columns, prefixes=(), contains=()):
    """
    Every column matching by name, most specific first: prefix matches in the
    order given, then substring matches. Duplicates removed.

    Prefix-before-substring matters in this sheet: 'IPP_Max (A)' and
    'VC@IPP_Max (V)' both contain 'ipp_max', and 'VRRM (V)' and
    'IR@VRWM_Max (uA)' both look like a standoff-voltage column to a naive
    substring search. Anchoring at the start of the header keeps the
    measurement columns from being mistaken for the conditions they are
    measured at.

    Returning a list rather than one column lets a field fall through to its
    next-best source when the preferred column is blank for a given part.
    """
    flat_cols = [(_flat(c), c) for c in columns]
    found = []
    for prefix in prefixes:
        for flat, col in flat_cols:
            if flat.startswith(prefix) and col not in found:
                found.append(col)
    for keyword in contains:
        for flat, col in flat_cols:
            if keyword in flat and col not in found:
                found.append(col)
    return found


def _find_col(columns, prefixes=(), contains=()):
    """First column matching by name, or None."""
    hits = _find_cols(columns, prefixes, contains)
    return hits[0] if hits else None


def _first_cell(row, columns):
    """First non-empty value across an ordered list of candidate columns."""
    for col in columns:
        val = _cell(row, col)
        if val is not None:
            return val
    return None


def _cell(row, col):
    """Raw value at a column, or None when missing/blank."""
    if col is None or col not in row.index:
        return None
    val = row[col]
    if pd.isna(val):
        return None
    if isinstance(val, str) and not val.strip():
        return None
    return val


def _num(val):
    """Leading number in a cell as a float, or None."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    m = re.search(r'[-+]?\d*\.?\d+', str(val))
    return float(m.group(0)) if m else None


def fetch_aos_specs_from_excel(aos_opn, aos_specs_df):
    """
    Fetches specifications for an AOS part from a pre-loaded pandas DataFrame.

    Column layout handled (headers are matched loosely, so minor renames and
    unit-suffix changes are fine):
        Part Number | Package | AEC-Q101 Qualified | Ppp (W) | VRRM (V) |
        VBR_Min (V) | VC@IPP_Max (V) | IPP_Max (A) | IR@VRWM_Max (uA) |
        CJ_Typ (pF)
    The older layout (Product / Protected Lines / Directional / VRWM max /
    VCL max / Lightning / (ESD) Contact / Cj typ) is still recognised.
    """
    specs_result = {
        "Device Name": aos_opn,
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Channels": "-",
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
        "Grade": "-",
    }

    if aos_opn is None or aos_specs_df is None or aos_specs_df.empty:
        return None

    aos_specs_df.columns = aos_specs_df.columns.str.strip()
    cols = list(aos_specs_df.columns)

    part_num_col = _find_col(
        cols,
        prefixes=("partnumber", "partno", "product"),
        contains=("partnumber", "partno", "product", "mfrpart"),
    )
    if not part_num_col:
        print(f"AOS Scraper: no part number column found. Headers are: {cols}")
        return None

    target = str(aos_opn).strip().lower()
    part_col_vals = aos_specs_df[part_num_col].astype(str).str.strip().str.lower()
    part_row_series = aos_specs_df[part_col_vals == target]
    if part_row_series.empty:
        # Fall back to a separator-insensitive comparison ('ASD05MA-Q1' vs 'ASD05MA Q1')
        target_flat = _flat(aos_opn)
        part_row_series = aos_specs_df[
            aos_specs_df[part_num_col].astype(str).map(_flat) == target_flat
        ]

    if part_row_series.empty:
        print(f"AOS Part '{aos_opn}' not found in the AOS Excel database.")
        return None

    part_row = part_row_series.iloc[0]
    print(f"Found specs for AOS Part '{aos_opn}' in Excel.")

    # --- Resolve every column once, against this sheet's actual headers ---
    c_package  = _find_col(cols, prefixes=("package",), contains=("package",))
    c_aec      = _find_col(cols, prefixes=("aec",), contains=("aecq101", "aec", "qualified"))
    c_ppp      = _find_col(cols, prefixes=("ppp", "pppm"), contains=("peakpulsepower",))
    c_vrwm     = _find_col(cols, prefixes=("vrrm", "vrwm"), contains=("standoff",))
    c_vcl      = _find_col(cols, prefixes=("vcipp", "vcl", "vc"), contains=("vcipp", "clamping"))
    # IEC 61000-4-5 comes from the 'IEC61000-4-5 (Lightning)' column when the
    # sheet has one; IPP_Max is the same rating under its raw name, used only
    # as a fallback.
    c_ipp      = (_find_cols(cols, prefixes=("iec6100045", "lightning"),
                             contains=("iec6100045", "lightning"))
                  + _find_cols(cols, prefixes=("ippmax", "ipp")))
    c_cj       = _find_col(cols, prefixes=("cjtyp", "cj"), contains=("iognd", "capacitance"))
    c_esd      = _find_col(cols, contains=("esdcontact", "contactdischarge", "iec6100042"))
    c_channels = _find_col(cols, contains=("protectedlines", "channel", "numberoflines"))
    c_dir      = _find_col(cols, contains=("directional", "direction", "unibi"))

    pkg = safe_strip(_cell(part_row, c_package))
    if pkg:
        specs_result["Package"] = pkg

    aec_raw = _cell(part_row, c_aec)
    if aec_raw is not None:
        aec_txt = str(aec_raw).strip().upper()
        specs_result["Grade"] = "Automotive" if aec_txt.startswith("Y") else "Commercial"

    ch = _num(_cell(part_row, c_channels))
    if ch is not None:
        specs_result["Channels"] = str(int(ch))

    dir_text = str(_cell(part_row, c_dir) or "").lower()
    if 'uni' in dir_text:
        specs_result["Direction"] = "Unidirectional"
    elif 'bi' in dir_text:
        specs_result["Direction"] = "Bidirectional"

    vrwm = _num(_cell(part_row, c_vrwm))
    if vrwm is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm:g} V"

    vcl = _num(_cell(part_row, c_vcl))
    if vcl is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vcl:g} V"

    ppp = _num(_cell(part_row, c_ppp))
    if ppp is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp:g} W"

    ipp_raw = _first_cell(part_row, c_ipp)
    if not c_ipp:
        print(f"AOS Scraper: no IEC 61000-4-5 / IPP_Max column found. Headers are: {cols}")
    ipp = _num(ipp_raw)
    if ipp is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp:g} A (8/20µs)"
    elif ipp_raw is not None:
        specs_result["IEC 61000-4-5"] = safe_strip(ipp_raw)

    esd_raw = _cell(part_row, c_esd)
    if esd_raw is not None:
        cleaned = str(esd_raw).lower().replace('±', '').replace('kv', '').strip()
        specs_result["IEC 61000-4-2"] = f"±{cleaned} kV" if cleaned else safe_strip(esd_raw)

    cap_raw = _cell(part_row, c_cj)
    cap = _num(cap_raw)
    if cap is not None:
        specs_result["Capacitance"] = f"{cap:.2f} pF"
    elif cap_raw is not None:
        specs_result["Capacitance"] = safe_strip(cap_raw)

    return specs_result