import pandas as pd
import re
from parsing import safe_strip


def _get_col(row, *keywords):
    """Case-insensitive partial column match, returns first hit or '-'."""
    for col in row.index:
        col_l = str(col).lower()
        for kw in keywords:
            if kw.lower() in col_l:
                val = row[col]
                return str(val).strip() if pd.notna(val) else "-"
    return "-"


def _num(raw, unit=""):
    """Extract leading number and append unit with space."""
    if not raw or raw == "-":
        return "-"
    m = re.search(r"([\d.]+)", str(raw))
    if m:
        v = m.group(1)
        return f"{v} {unit}".strip() if unit else v
    return "-"


def _get_ipp_col(row):
    """
    The IPP_Max (A) column - the surge current the IEC 61000-4-5 rating is
    built from.

    Matched on a column *starting* with Ipp rather than merely containing it:
    "Vc@Ipp_Max (V)" contains the same text and sits earlier in these sheets,
    so a substring match returns the clamping voltage instead of the current.
    """
    cols = [(re.sub(r'[^a-z0-9]+', '', str(c).lower()), c) for c in row.index]
    for prefix in ("ippmax", "ipp"):
        for flat, col in cols:
            if flat.startswith(prefix):
                val = row[col]
                return str(val).strip() if pd.notna(val) else "-"
    return "-"


def _anbon_direction(part_number, uni_bi_col_val=""):
    """
    Determine direction for Anbon parts:
    - Parts starting with 'AS' are bidirectional
    - Parts containing 'C' just before end / -Q1 suffix are bidirectional
      e.g. ESD5Z12C-Q1, GBLC03C-Q1, SM15C-Q1
    - Fall back to UNI/BI column value if present
    - Default: Unidirectional
    """
    pn = str(part_number).upper()
    if pn.startswith("AS"):
        return "Bidirectional"
    base = re.sub(r'-Q\d.*$', '', pn)   # strip -Q1 suffix
    if re.search(r'C(?:\d*|-|$)', base) and re.search(r'C', base[-6:]):
        return "Bidirectional"
    col = str(uni_bi_col_val).upper()
    if "BI" in col:
        return "Bidirectional"
    if "UNI" in col:
        return "Unidirectional"
    return "Unidirectional"


def _anbon_channels(part_number, direction="-"):
    """
    Channel count for Anbon parts - the parametric tables do not carry it.
      - anything starting with SR is 2-channel, whatever its direction
      - SM parts are 2-channel when bidirectional
    Everything else is treated as single channel.
    """
    pn = str(part_number).upper().strip()
    if pn.startswith("SR"):
        return "2"
    if pn.startswith("SM") and "BI" in str(direction).upper():
        return "2"
    return "1"


def _parse_anbon_esd_row(row, part_number):
    auto = "Y" in str(_get_col(row, "aec-q101", "aec_q101", "aec")).upper()
    pkg_raw = safe_strip(_get_col(row, "package"))
    ipp_raw = _get_ipp_col(row)
    iec45 = f"{_num(ipp_raw)} A (8/20µs)" if _num(ipp_raw) != "-" else "-"
    direction = _anbon_direction(part_number)
    return {
        "Device Name": part_number,
        "Source File": "Anbon ESD",
        "Package": pkg_raw,
        "Grade": "Automotive" if auto else "Commercial",
        "Voltage - Reverse Standoff (Typ)": _num(_get_col(row, "vrrm"), "V"),
        "Voltage - Clamping (Max) @ Ipp": _num(_get_col(row, "vc@ipp", "vc@"), "V"),
        "Power Dissipation (Pd)": _num(_get_col(row, "ppp", "pppm"), "W"),
        "Capacitance": _num(_get_col(row, "cj_typ", "cj typ", "cj"), "pF"),
        "Channels": _anbon_channels(part_number, direction),
        "Direction": direction,
        "IEC 61000-4-5": iec45,
        "IEC 61000-4-2": "-",
        "Price ($/ku)": "-",
    }


def _parse_anbon_tvs_row(row, part_number):
    auto = "Y" in str(_get_col(row, "aec-q101", "aec_q101", "aec")).upper()
    pkg_raw = safe_strip(_get_col(row, "package"))
    uni_bi_val = _get_col(row, "uni/bi", "uni_bi")
    ipp_raw = _get_ipp_col(row)
    iec45 = f"{_num(ipp_raw)} A (8/20µs)" if _num(ipp_raw) != "-" else "-"
    direction = _anbon_direction(part_number, uni_bi_val)
    return {
        "Device Name": part_number,
        "Source File": "Anbon TVS",
        "Package": pkg_raw,
        "Grade": "Automotive" if auto else "Commercial",
        "Voltage - Reverse Standoff (Typ)": _num(_get_col(row, "vrwm"), "V"),
        "Voltage - Clamping (Max) @ Ipp": _num(_get_col(row, "vc@ipp", "vc@"), "V"),
        "Power Dissipation (Pd)": _num(_get_col(row, "pppm", "ppp"), "W"),
        "Capacitance": "-",
        "Channels": _anbon_channels(part_number, direction),
        "Direction": direction,
        "IEC 61000-4-5": iec45,
        "IEC 61000-4-2": "-",
        "Price ($/ku)": "-",
    }


def _parse_anbon_zener_row(row, part_number):
    auto = "Y" in str(_get_col(row, "aec-q101", "aec_q101", "aec")).upper()
    pkg_raw = safe_strip(_get_col(row, "package"))
    vz_raw = _get_col(row, "vz@izt_n", "vz@izt n", "vz@izt_nom")
    if vz_raw == "-":
        vz_min = re.search(r"([\d.]+)", _get_col(row, "vz@izt_min", "vz@izt min"))
        vz_max = re.search(r"([\d.]+)", _get_col(row, "vz@izt_max", "vz@izt max"))
        if vz_min and vz_max:
            vz_raw = str((float(vz_min.group(1)) + float(vz_max.group(1))) / 2)
    vz = _num(vz_raw, "V") if vz_raw != "-" else "-"
    tol_raw = _get_col(row, "vz (%)", "vz(%)", "tolerance", "vz %")
    tol = f"±{_num(tol_raw)}%" if _num(tol_raw) != "-" else "-"
    pd_raw = _get_col(row, "pd (mw)", "pd(mw)", "pd")
    pd_num = _num(pd_raw)
    if pd_num != "-":
        try:
            pd_val = f"{float(pd_num) / 1000:g} W"
        except ValueError:
            pd_val = pd_num + " mW"
    else:
        pd_val = "-"
    return {
        "Device Name": part_number,
        "Source File": "DigiKey Zener",
        "Package": pkg_raw,
        "Grade": "Automotive" if auto else "Commercial",
        "Voltage - Reverse Standoff (Typ)": vz,
        "Tolerance": tol,
        "Power Dissipation (Pd)": pd_val,
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Capacitance": "-",
        "Channels": _anbon_channels(part_number),
        "Direction": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Price ($/ku)": "-",
    }


def _find_part(norm_input, df, part_col):
    from applesauce import (
        _find_exact_match, _find_partial_match,
        _find_reverse_partial_match, _find_fuzzy_match,
    )
    return (
        _find_exact_match(norm_input, df, part_col) or
        _find_partial_match(norm_input, df, part_col) or
        _find_reverse_partial_match(norm_input, df, part_col) or
        _find_fuzzy_match(norm_input, df, part_col, min_len=6) or
        _find_fuzzy_match(norm_input, df, part_col, min_len=4)
    )


def fetch_anbon_specs(part_input, anbon_esd_df, anbon_zener_df, anbon_tvs_df):
    """Search all three Anbon spec files. Returns specs dict or None."""
    norm_input = re.sub(r'[\W_]+', '', str(part_input).upper())

    PART_KW = ["part number", "part no", "partno"]

    for df, label, parse_fn in [
        (anbon_esd_df,   "ESD",   _parse_anbon_esd_row),
        (anbon_tvs_df,   "TVS",   _parse_anbon_tvs_row),
        (anbon_zener_df, "Zener", _parse_anbon_zener_row),
    ]:
        if df is None or df.empty:
            continue
        part_col = next((c for c in df.columns if any(kw in c.lower() for kw in PART_KW)), None)
        if not part_col:
            continue
        match = _find_part(norm_input, df, part_col)
        if match:
            row = df[df[part_col].astype(str) == str(match)].iloc[0]
            print(f"Found '{match}' in Anbon {label} database.")
            return parse_fn(row, match)
    return None