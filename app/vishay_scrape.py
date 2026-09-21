import pandas as pd
import re
from parsing import safe_strip

# A VZ min/max window wider than this around the nominal is the series' span
# rather than the part's own tolerance, so it is not read as a tolerance.
MAX_PLAUSIBLE_TOLERANCE_PCT = 20.0


def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    The search is case-insensitive. Returns the value from the matched column.
    """
    for col_header in row.index:
        for keyword in keywords:
            if keyword.lower() in str(col_header).lower():
                return row[col_header]
    return None # Return None if no matching column is found


def _numeric_or_none(value):
    """First number in a cell, or None if the cell is blank/non-numeric."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    numeric_match = re.search(r"[\d.]+", str(value))
    if not numeric_match:
        return None
    try:
        return float(numeric_match.group(0))
    except (ValueError, TypeError):
        return None


def _parse_vishay_esd_row(part_row, found_df_name, part_number):
    """
    Parses a row from the Vishay ESD datasheet.

    Columns: Part Number, Data Sheet, Package, Reverse Working Voltage (V),
    Capacitance (pF), Leakage Current (uA), Peak Pulse Current 8/20 us (A),
    Peak Pulse Power 8/20 us (W), Protectable Lines, ESD Immunity IEC
    61000-4-2 (kV), Max. Operating Junction Temperature (C), AEC-Q101
    Qualified.

    The sheet states no directionality and no clamping voltage, so those two
    stay at "-".
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": "Commercial",
        "Direction": (
            "Bidirectional"
            if (
                str(part_number).upper().startswith(("VCAN", "VCUT"))
                or str(part_number).upper().endswith(("SD0", "DD1"))
            )
            else "Unidirectional"
        ),
        "Channels": "1",
        "Package": "-",
        "Package Info": {},
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    # Package
    pkg_name = safe_strip(_get_value_by_keywords(part_row, ["Package"]))
    specs_result["Package"] = pkg_name
    specs_result["Package Info"] = {"name": pkg_name, "version": "-"}

    # Grade
    aec_text = safe_strip(_get_value_by_keywords(part_row, ["AEC-Q101"])).lower()
    if aec_text.startswith("yes"):
        specs_result["Grade"] = "Automotive"

    # Channels
    lines_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Protectable Lines"]))
    if lines_val is not None and lines_val > 0:
        specs_result["Channels"] = str(int(lines_val))

    # VRWM
    vrwm_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Reverse Working Voltage"]))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    # Capacitance
    cap_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Capacitance"]))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    # Leakage current (Ir)
    ir_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Leakage Current"]))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    # Surge (Ipp). The keyword has to name the current column outright - a bare
    # "Peak Pulse" would match the power column sitting next to it.
    surge_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Peak Pulse Current"]))
    if surge_val is not None:
        specs_result["IEC 61000-4-5"] = f"{surge_val:g} A (8/20µs)"

    # Peak pulse power
    ppp_val = _numeric_or_none(_get_value_by_keywords(part_row, ["Peak Pulse Power"]))
    if ppp_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp_val:g} W"

    # ESD immunity (already stated in kV in this sheet)
    esd_val = _numeric_or_none(_get_value_by_keywords(part_row, ["ESD Immunity", "IEC 61000-4-2"]))
    if esd_val is not None:
        specs_result["IEC 61000-4-2"] = f"±{esd_val:g} kV"

    return specs_result


def _parse_vishay_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from the Vishay Zener datasheet.

    Columns: Series, Data Sheet, Part Number, VZ Min (V), VZ Max (V), VZ Nom
    (V), IZT (mA), PTOT (mW), Package, VZ Specification, AEC-Q101 Qualified.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": "Commercial",
        "Voltage - Reverse Standoff (Typ)": "-", # This will hold Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-", # Not present in Zener file headers
        "Capacitance": "-", # Not present in Zener file headers
        "Package": "-",
        "Package Info": {},
    }

    # Package
    pkg_name = safe_strip(_get_value_by_keywords(part_row, ["Package"]))
    specs_result["Package"] = pkg_name
    specs_result["Package Info"] = {"name": pkg_name, "version": "-"}

    # Grade
    aec_text = safe_strip(_get_value_by_keywords(part_row, ["AEC-Q101"])).lower()
    if aec_text.startswith("yes"):
        specs_result["Grade"] = "Automotive"

    # Vz (mapped to standard key for compatibility with the main script). Each
    # keyword names its column outright - a bare "VZ" would match whichever of
    # the three VZ columns comes first.
    vz_nom = _numeric_or_none(_get_value_by_keywords(part_row, ["VZ Nom"]))
    vz_min = _numeric_or_none(_get_value_by_keywords(part_row, ["VZ Min"]))
    vz_max = _numeric_or_none(_get_value_by_keywords(part_row, ["VZ Max"]))

    if vz_nom is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"
    elif vz_min is not None and vz_max is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{(vz_min + vz_max) / 2:g} V"

    # Tolerance: the sheet has no tolerance column, so it comes out of the VZ
    # window around the nominal. Some rows carry the whole series' span in
    # those two cells (a 30 V part listed as 3.9 V - 75 V), so the window is
    # only read as a tolerance when it is narrow enough to be one.
    if (vz_nom is not None and vz_nom > 0
            and vz_min is not None and vz_max is not None
            and vz_min <= vz_nom <= vz_max):
        spread_pct = max(vz_nom - vz_min, vz_max - vz_nom) / vz_nom * 100
        if 0 < spread_pct <= MAX_PLAUSIBLE_TOLERANCE_PCT:
            specs_result["Tolerance"] = f"±{spread_pct:.3g}%"

    # Power dissipation (Ptot is in mW, convert to W)
    ptot_val = _numeric_or_none(_get_value_by_keywords(part_row, ["PTOT"]))
    if ptot_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ptot_val / 1000.0:g} W"

    return specs_result


def fetch_vishay_specs_from_excel(part_number, vishay_dfs):
    """
    Searches across Vishay DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. ESD).

    Only the ESD and Zener sheets are searched. The TVS sheet is deliberately
    not part of the map, so no Vishay TVS part reaches a cross from here.
    """
    if part_number is None or not vishay_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["ESD", "Zener"]

    all_df_keys = list(vishay_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["part number", "mfr part"]

    for df_name in ordered_search_keys:
        df = vishay_dfs[df_name]
        if df is None or df.empty:
            continue

        part_num_col = None
        for col in df.columns:
            for keyword in part_num_col_keywords:
                if keyword in str(col).lower():
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
        print(f"Part '{part_number}' not found in any Vishay database.")
        return None

    print(f"Found specs for Part '{part_number}' in Vishay database '{found_df_name}'.")

    if "zener" in found_df_name.lower():
        return _parse_vishay_zener_row(part_row, found_df_name, part_number)
    else:
        return _parse_vishay_esd_row(part_row, found_df_name, part_number)