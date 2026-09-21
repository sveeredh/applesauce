import pandas as pd
import re
import xlrd
from parsing import safe_strip


# A VBR/VZ window wider than this around the nominal is a series span rather
# than the part's own tolerance, so it is not read as a tolerance.
MAX_PLAUSIBLE_TOLERANCE_PCT = 20.0

def load_yangjie_xls(path, header_row=0):
    """
    Yangjie's .xls files are not written by Excel, and xlrd's OLE2 directory
    check rejects most of them ("Workbook corruption: seen[2] == 4"). The data
    itself is fine, so the book is opened with that check disabled and the
    sheet is turned into a DataFrame directly - pandas sniffs the file format
    before the engine runs, so it cannot pass the flag through.
    """
    book = xlrd.open_workbook(path, ignore_workbook_corruption=True)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    if not rows or header_row >= len(rows):
        return pd.DataFrame()
    header = [str(h).strip() for h in rows[header_row]]
    return pd.DataFrame(rows[header_row + 1:], columns=header)

def _norm_header(col_header):
    """Header text with spaces, underscores and case removed, for matching."""
    return re.sub(r'[\s_]+', '', str(col_header)).upper()


def _get_header_by_prefix(row, *prefixes):
    """
    First column header whose normalized text starts with one of the prefixes.

    Yangjie repeats parameter names inside other headers - "IR@VRWM(uA)"
    contains "VRWM", "VC@IPP(V)" contains "IPP" - so a plain substring search
    returns whichever column happens to sit first. Anchoring at the start of
    the header keeps each parameter on its own column.
    """
    for prefix in prefixes:
        target = _norm_header(prefix)
        for col_header in row.index:
            if _norm_header(col_header).startswith(target):
                return col_header
    return None


def _get_header_by_tokens(row, *token_sets):
    """
    First column header containing EVERY token in one of the given token sets.
    Tells apart headers that share a word ("VBR _Min(V)" vs "VBR _Max(V)",
    "VZ@IZT_Nom(A)" vs "VZ@IZT_Max(A)").
    """
    for tokens in token_sets:
        for col_header in row.index:
            header = _norm_header(col_header)
            if all(_norm_header(token) in header for token in tokens):
                return col_header
    return None


def _numeric_or_none(value):
    """
    First number in a cell, or None if the cell is blank or non-numeric.

    Yangjie writes an unstated parameter as "/", which has to read as missing
    rather than as a zero.
    """
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    text = str(value).strip()
    if text in ("", "/", "-", "N/A", "NA"):
        return None
    numeric_match = re.search(r"[\d.]+", text)
    if not numeric_match:
        return None
    try:
        return float(numeric_match.group(0))
    except (ValueError, TypeError):
        return None


def _value_by_prefix(row, *prefixes):
    """Cell value for the first header matching one of the prefixes."""
    header = _get_header_by_prefix(row, *prefixes)
    return row.get(header) if header is not None else None


def _yangjie_grade(found_df_name):
    """Grade comes from which of the two files the row was found in."""
    return "Automotive" if "auto" in str(found_df_name).lower() else "Commercial"


def _yangjie_direction(part_row):
    """The Type column states Bi or Uni outright."""
    type_text = safe_strip(_value_by_prefix(part_row, "Type")).lower()
    if type_text.startswith("uni"):
        return "Unidirectional"
    if type_text.startswith("bi"):
        return "Bidirectional"
    return "-"


def _yangjie_package(part_row):
    """
    Package name lives in the Chinese header 封装名称. The English fallbacks
    cover a sheet that has been re-headered.
    """
    return safe_strip(_value_by_prefix(part_row, "封装名称", "Package", "PKG"))


def _parse_yangjie_esd_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Yangjie ESD datasheet (yangjie_esd_specs /
    yangjie_auto_esd_specs).

    Columns: 产品名称, 封装名称, 规格书, Type, VRWM(V), VBR _Min(V),
    VBR _Max(V), IPP(A), VC@IPP(V), Cj _TYP(PF), IR@VRWM(uA), Tj (C), Status.

    The sheet states no ESD contact rating, no peak pulse power and no channel
    count, so those stay at their defaults.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yangjie_grade(found_df_name),
        "Direction": _yangjie_direction(part_row),
        "Channels": "1",
        "Package": _yangjie_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    vrwm_val = _numeric_or_none(_value_by_prefix(part_row, "VRWM"))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "VC@IPP", "VC@"))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "IPP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20µs)"

    cap_val = _numeric_or_none(_value_by_prefix(part_row, "Cj"))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR@"))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def _parse_yangjie_tvs_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Yangjie TVS datasheet (yangjie_tvs_specs /
    yangjie_auto_tvs_specs).

    Same layout as the ESD sheet plus PPk (W), and without Cj - so these rows
    carry a peak pulse power but no capacitance.

    Columns: 产品名称, 封装名称, 规格书, Type, PPk (W), VRWM(V), VBR _Min(V),
    VBR _Max(V), IPP(A), VC@IPP(V), IR@VRWM(uA), Tj (C), Status.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yangjie_grade(found_df_name),
        "Direction": _yangjie_direction(part_row),
        "Channels": "1",
        "Package": _yangjie_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    ppk_val = _numeric_or_none(_value_by_prefix(part_row, "PPk"))
    if ppk_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppk_val:g} W"

    vrwm_val = _numeric_or_none(_value_by_prefix(part_row, "VRWM"))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "VC@IPP", "VC@"))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "IPP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20µs)"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR@"))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def _parse_yangjie_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Yangjie Zener datasheet (yangjie_zener_specs /
    yangjie_auto_zener_specs).

    Columns: 产品名称, 封装名称, 规格书, PD(mW), VZ@IZT_Min(A), VZ@IZT_Nom(A),
    VZ@IZT_Max(A), IZT(mA), Zzt@Izt, IZk(mA), Zzt@Izk, IR@VR(uA), VR(V),
    Status.

    The sheet has no tolerance column, so the tolerance comes out of the
    VZ min/nom/max window.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yangjie_grade(found_df_name),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",  # not present in the Zener file headers
        "Package": _yangjie_package(part_row),
    }

    # Power: PD is stated in mW.
    pd_val = _numeric_or_none(_value_by_prefix(part_row, "PD"))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val / 1000.0:g} W"

    # Vz. The three VZ@IZT columns share a prefix, so each is picked by the
    # Min/Nom/Max token rather than by position.
    vz_nom = _numeric_or_none(part_row.get(_get_header_by_tokens(part_row, ("VZ@IZT", "NOM"))))
    vz_min = _numeric_or_none(part_row.get(_get_header_by_tokens(part_row, ("VZ@IZT", "MIN"))))
    vz_max = _numeric_or_none(part_row.get(_get_header_by_tokens(part_row, ("VZ@IZT", "MAX"))))

    if vz_nom is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"
    elif vz_min is not None and vz_max is not None:
        vz_nom = (vz_min + vz_max) / 2
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"

    # Tolerance from the window around the nominal, guarded so a series-wide
    # span is not mistaken for one part's tolerance.
    if (vz_nom is not None and vz_nom > 0
            and vz_min is not None and vz_max is not None
            and vz_min <= vz_nom <= vz_max):
        spread_pct = max(vz_nom - vz_min, vz_max - vz_nom) / vz_nom * 100
        if 0 < spread_pct <= MAX_PLAUSIBLE_TOLERANCE_PCT:
            specs_result["Tolerance"] = f"±{spread_pct:.3g}%"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR@"))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def fetch_yangjie_specs_from_excel(part_number, yangjie_dfs):
    """
    Searches across Yangjie DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS vs. ESD).

    The automotive and standard files hold different part numbers, so the file
    a row is found in is what states its grade.
    """
    if part_number is None or not yangjie_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = [
        "Auto_Zener", "Zener",
        "Auto_ESD", "ESD",
        "Auto_TVS", "TVS",
    ]

    all_df_keys = list(yangjie_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["产品名称", "part number", "mfr part", "product"]

    for df_name in ordered_search_keys:
        df = yangjie_dfs[df_name]
        if df is None or df.empty:
            continue

        part_num_col = None
        for col in df.columns:
            for keyword in part_num_col_keywords:
                if keyword.lower() in str(col).lower():
                    part_num_col = col
                    break
            if part_num_col:
                break

        if not part_num_col:
            continue

        part_row_series = df[df[part_num_col].astype(str).str.strip().str.lower() == part_number.lower()]

        if not part_row_series.empty:
            part_row = part_row_series.iloc[0]
            found_df_name = df_name
            break

    if part_row is None:
        print(f"Part '{part_number}' not found in any Yangjie database.")
        return None

    print(f"Found specs for Part '{part_number}' in Yangjie database '{found_df_name}'.")

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_yangjie_zener_row(part_row, found_df_name, part_number)
    elif "tvs" in name_lower:
        return _parse_yangjie_tvs_row(part_row, found_df_name, part_number)
    else:
        return _parse_yangjie_esd_row(part_row, found_df_name, part_number)