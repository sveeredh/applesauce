import pandas as pd
import re
from parsing import safe_strip

def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    The search is case-insensitive. Returns the value from the matched column.
    """
    for col_header in row.index:
        for keyword in keywords:
            # Exact match is needed for short names like 'IR' vs 'VRWM'
            if keyword.lower() == str(col_header).lower().strip():
                return row[col_header]
            # Fallback to 'contains' for longer headers with units
            if keyword.lower() in str(col_header).lower():
                return row[col_header]
    return None

def _parse_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Jiangsu Zener datasheet.
    """
    specs_result = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Commercial", # Assuming Commercial grade as no auto file exists
        "Voltage - Reverse Standoff (Typ)": "-", # VZTyp for comparison
        "Power Dissipation (Pd)": "-",
        "Package": "-",
    }

    # Package
    specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))
    

    # Vz (mapped to standard key for compatibility)
    vz_val = _get_value_by_keywords(part_row, ["VZTyp"])
    if pd.notna(vz_val):
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(vz_val):g} V"

    # Power Dissipation (PD is in mW, convert to W)
    pd_val_mw = _get_value_by_keywords(part_row, ["PD"])
    if pd.notna(pd_val_mw):
        pd_in_watts = float(pd_val_mw) / 1000.0
        specs_result["Power Dissipation (Pd)"] = f"{pd_in_watts:g} W"

    return specs_result

def _parse_tvs_esd_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Jiangsu TVS or ESD datasheet.
    """
    specs_result = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Commercial",
        "Package": "-",
        "Direction": _get_jiangsu_direction(
            part_number,
            found_df_name
        ), # Jiangsu datasheets do not specify, default to Uni
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-", # Not present in headers, defaults to "-"
        "IEC 61000-4-2": "-", # Not present in headers, defaults to "-"
        "Capacitance": "-",
        "Channels": "1", # Default to 1 channel
        "Power Dissipation (Pd)": "-"
    }

    # Package
    specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))

    

    # Peak Pulse Power (Ppp)
    ppp_val = _get_value_by_keywords(part_row, ["Ppp"])
    if pd.notna(ppp_val) and str(ppp_val).strip() != "":
        try:
            specs_result["Power Dissipation (Pd)"] = f"{float(ppp_val):g} W"
        except (ValueError, TypeError):
            specs_result["Power Dissipation (Pd)"] = safe_strip(ppp_val)
    # VRWM
    vrwm_val = _get_value_by_keywords(part_row, ["VRWM"])
    if pd.notna(vrwm_val):
        # --- FIX: Add try-except for robust float conversion ---
        try:
            specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(vrwm_val):g} V"
        except (ValueError, TypeError):
            specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vrwm_val)

    # Clamping Voltage (VC)
    vcl_val = _get_value_by_keywords(part_row, ["VC"])
    if pd.notna(vcl_val):
        # --- FIX: Add try-except for robust float conversion ---
        try:
            specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(vcl_val):g} V"
        except (ValueError, TypeError):
            specs_result["Voltage - Clamping (Max) @ Ipp"] = safe_strip(vcl_val)

    # Capacitance (CJ)
    cap_val = _get_value_by_keywords(part_row, ["CJ"])
    if pd.notna(cap_val) and str(cap_val).strip() != "":
        # --- FIX: Add try-except for robust float conversion ---
        try:
            specs_result["Capacitance"] = f"{float(cap_val):.2f} pF"
        except (ValueError, TypeError):
            specs_result["Capacitance"] = safe_strip(cap_val)

    # Surge (Ipp)
    ipp_val = _get_value_by_keywords(part_row, ["IPP"])
    if pd.notna(ipp_val) and str(ipp_val).strip() != "":
        # --- FIX: Add try-except for robust float conversion ---
        try:
            specs_result["IEC 61000-4-5"] = f"{float(ipp_val):g} A (8/20µs)"
        except (ValueError, TypeError):
            specs_result["IEC 61000-4-5"] = safe_strip(ipp_val)
    
    esd_val = _get_value_by_keywords(part_row, ["vesd"])
    if pd.notna(esd_val):
        try:
            specs_result["IEC 61000-4-2"] = f"±{float(esd_val):g} kV"
        except (ValueError, TypeError):
            specs_result["IEC 61000-4-2"] = safe_strip(esd_val)

    return specs_result


def _get_jiangsu_direction(part_number, found_df_name):
    name = str(part_number).strip().upper()

    # TVS naming
    if "tvs" in str(found_df_name).lower():
        if name.endswith("CA"):
            return "Bidirectional"
        if name.endswith("A"):
            return "Unidirectional"

        return "Unidirectional"

    # ESD naming:
    # ESD...B... or CESD...B... = bidirectional
    if "ESD" in name and "B" in name:
        return "Bidirectional"

    return "Unidirectional"

def fetch_jiangsu_specs_from_excel(part_number, jiangsu_dfs):
    """
    Searches across Jiangsu DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS/ESD).
    """
    if part_number is None or not jiangsu_dfs:
        return None

    # Search priority: Zener > TVS > ESD
    search_priority = ["Zener", "TVS", "ESD"]
    
    all_df_keys = list(jiangsu_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    for df_name in ordered_search_keys:
        df = jiangsu_dfs[df_name]
        if df is None or df.empty:
            continue
        
        # Clean column headers by removing extra spaces and newlines
        df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]
        
        part_num_col = next((col for col in df.columns if 'part number' in col.lower()), None)
        if not part_num_col:
            continue

        # Search for a case-insensitive match at the start of the string
        part_row_series = df[df[part_num_col].astype(str).str.strip().str.lower() == part_number.lower()]
        
        if not part_row_series.empty:
            part_row = part_row_series.iloc[0]
            print(f"Found specs for Part '{part_number}' in Jiangsu database '{df_name}'.")
            
            # Dispatch to the correct parser
            if "zener" in df_name.lower():
                return _parse_zener_row(part_row, df_name, part_number)
            else:
                return _parse_tvs_esd_row(part_row, df_name, part_number)

    print(f"Part '{part_number}' not found in any Jiangsu database.")
    return None