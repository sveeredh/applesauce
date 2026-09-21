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
            if keyword.lower() in str(col_header).lower():
                return row[col_header]
    return None # Return None if no matching column is found

# --- File: nexperia_scrape.py ---

# ADD THIS ENTIRE NEW FUNCTION
# --- File: nexperia_scrape.py ---

# REPLACE the function formerly named _parse_nexperia_digikey_zener_row with this one:
def _parse_nexperia_digikey_tvs_row(part_row, source_df_name, part_number):
    """
    Parses a row from a Digi-Key formatted Nexperia TVS datasheet (backup file).
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": source_df_name,
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "IEC 61000-4-5": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Channels": "1",
        "Package Info": {}
    }
    
    # Grade (Assume Automotive if AEC is mentioned in any column)
    if any('aec-q' in str(val).lower() for val in part_row.values):
        specs_result["Grade"] = "Automotive"

    # Package
    pkg_supplier = safe_strip(_get_value_by_keywords(part_row, ["Supplier Device Package"]))
    specs_result["Package"] = pkg_supplier if pkg_supplier != "-" else safe_strip(_get_value_by_keywords(part_row, ["Package / Case"]))
    specs_result["Package Info"] = {"name": specs_result["Package"], "version": "-"}


    # Direction and Channels
    uni_channels_raw = _get_value_by_keywords(part_row, ["Unidirectional Channels"])
    bi_channels_raw = _get_value_by_keywords(part_row, ["Bidirectional Channels"])

    try:
        # Give priority to the "Bidirectional Channels" column
        if bi_channels_raw and bi_channels_raw != "-" and int(float(bi_channels_raw)) > 0:
            specs_result["Direction"] = "Bidirectional"
            specs_result["Channels"] = str(int(float(bi_channels_raw)))
        # Fall back to the "Unidirectional Channels" column
        elif uni_channels_raw and uni_channels_raw != "-" and int(float(uni_channels_raw)) > 0:
            specs_result["Direction"] = "Unidirectional"
            specs_result["Channels"] = str(int(float(uni_channels_raw)))
    except (ValueError, TypeError):
        # If values are not clean numbers, the default values will be used.
        pass
    
    # Key TVS Parameters
    # --- AFTER ---
    # Key TVS Parameters
    vr_val_raw = _get_value_by_keywords(part_row, ["Voltage - Reverse"])
    if pd.notna(vr_val_raw):
        # Use regex to find the first numeric part of the string, ignoring "(Max)", etc.
        numeric_match = re.search(r"([\d.]+)", str(vr_val_raw))
        if numeric_match:
            try:
                # Extract only the number and format it
                specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(numeric_match.group(1)):g} V"
            except (ValueError, TypeError):
                # Fallback in case of an unexpected format
                specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vr_val_raw)
        else:
            # If no number is found, use the original string
            specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vr_val_raw)
    clamping_val_raw = _get_value_by_keywords(part_row, ["Voltage - Clamping"])
    if pd.notna(clamping_val_raw):
        specs_result["Voltage - Clamping (Max) @ Ipp"] = safe_strip(clamping_val_raw)

    surge_val_raw = _get_value_by_keywords(part_row, ["Current - Peak Pulse"])
    if pd.notna(surge_val_raw):
        numeric_match = re.search(r"([\d.]+)", str(surge_val_raw))
        if numeric_match:
            specs_result["IEC 61000-4-5"] = f"{float(numeric_match.group(1)):g} A (8/20µs)"

    cap_raw = _get_value_by_keywords(part_row, ["Capacitance"])
    if pd.notna(cap_raw):
        specs_result["Capacitance"] = safe_strip(cap_raw)

    return specs_result

def _parse_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Nexperia Zener datasheet to extract Zener-specific parameters.
    """
    specs_result = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Automotive" if "auto" in found_df_name.lower() else "Non-Automotive",
        "Voltage - Reverse Standoff (Typ)": "-", # This will hold Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-", # Not present in Zener file headers
        "Capacitance": "-", # Not present in Zener file headers
        "Package": "-",
        "Package Info": {}
    }

    # Package Information
    pkg_name = safe_strip(_get_value_by_keywords(part_row, ["package name"]))
    pkg_version = safe_strip(_get_value_by_keywords(part_row, ["package version"]))
    specs_result["Package"] = pkg_name if pkg_name != "-" else pkg_version
    specs_result["Package Info"] = {"name": pkg_name, "version": pkg_version}

    # Vz (mapped to standard key for compatibility with the main script)
    vz_raw = _get_value_by_keywords(part_row, ["vz [nom]"])
    if pd.notna(vz_raw) and vz_raw != "-":
        # This regex finds the first group of digits (and an optional decimal point),
        # effectively ignoring any text like "(Max)" that follows.
        numeric_match = re.search(r"([\d.]+)", str(vz_raw))
        if numeric_match:
            # We extract only the matched numeric part (group 1) and format it.
            specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(numeric_match.group(1)):g} V"

    # Tolerance
    tol_val_str = safe_strip(_get_value_by_keywords(part_row, ["tolerance"]))
    if tol_val_str != "-":
        # Use regex to extract only the numeric part (e.g., '2' from '2%')
        numeric_match = re.search(r'(\d+(?:\.\d+)?)', tol_val_str)
        if numeric_match:
            try:
                numeric_val = float(numeric_match.group(1))
                # A cell with no "%" states the tolerance as a fraction
                # (0.01 = 1%). Only scaled when the value is below 1, so a bare
                # "5" still reads as 5% rather than 500%.
                if "%" not in tol_val_str and numeric_val < 1:
                    numeric_val *= 100
                specs_result["Tolerance"] = f"±{numeric_val:g}%"
            except (ValueError, TypeError):
                # Fallback if conversion somehow fails
                specs_result["Tolerance"] = tol_val_str
    
    # Power Dissipation (Ptot is in mW, convert to W)
    pd_val_mw = _get_value_by_keywords(part_row, ["ptot"])
    if pd.notna(pd_val_mw):
        pd_in_watts = float(pd_val_mw) / 1000.0
        specs_result["Power Dissipation (Pd)"] = f"{pd_in_watts:g} W"

    return specs_result

def _parse_tvs_row(part_row, found_df_name, part_number):
    """
    Parses a row from a standard TVS/ESD/Diode datasheet.
    """
    specs_result = { "Device Name": part_number, "IEC 61000-4-5": "-", "IEC 61000-4-2": "-", "Capacitance": "-", "Channels": "-", "Package": "-", "Direction": "-", "Voltage - Reverse Standoff (Typ)": "-", "Voltage - Clamping (Max) @ Ipp": "-", "Price ($/ku)": "-", "Grade": "-" }
    
    specs_result["Source File"] = found_df_name
    pkg_name = safe_strip(_get_value_by_keywords(part_row, ["package name"]))
    pkg_version = safe_strip(_get_value_by_keywords(part_row, ["package version"]))

    # Set a simple name for display in the final table (e.g., "DFN1006")
    specs_result["Package"] = pkg_name if pkg_name != "-" else pkg_version
    # Store the detailed dictionary for the new normalizer function
    specs_result["Package Info"] = {"name": pkg_name, "version": pkg_version}
    
    ch_val_raw = safe_strip(_get_value_by_keywords(part_row, ["nr of lines"]))
    try:
        specs_result["Channels"] = str(int(float(ch_val_raw)))
    except (ValueError, TypeError):
        specs_result["Channels"] = "1"

    config_text = safe_strip(_get_value_by_keywords(part_row, ["configuration"])).lower()
    if 'single' in config_text or 'unidi' in config_text:
        specs_result["Direction"] = "Unidirectional"
    elif 'dual' in config_text or 'bidi' in config_text:
        specs_result["Direction"] = "Bidirectional"
    
    specs_result["Grade"] = "Automotive" if "auto" in found_df_name.lower() else "Non-Automotive"

    vrwm_val = _get_value_by_keywords(part_row, ["vrwm"])
    if pd.notna(vrwm_val):
        # Proactively parse the string to handle non-numeric formats
        vrwm_str = str(vrwm_val)
        numeric_match = re.search(r'([\d.]+)', vrwm_str)
        if numeric_match:
            try:
                numeric_val = float(numeric_match.group(1))
                specs_result["Voltage - Reverse Standoff (Typ)"] = f"{numeric_val:g} V"
            except (ValueError, TypeError):
                specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vrwm_str) # Fallback
        else:
            specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vrwm_str) # Fallback

    # --- AFTER ---
    vrwm_val = _get_value_by_keywords(part_row, ["Reverse Standoff", "VRWM(V)"])
    if pd.notna(vrwm_val) and vrwm_val != "-":
        # Use regex to find the first numeric part of the string, ignoring "(Max)", etc.
        numeric_match = re.search(r"([\d.]+)", str(vrwm_val))
        if numeric_match:
            try:
                # Extract only the number and format it
                specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(numeric_match.group(1)):g} V"
            except (ValueError, TypeError):
                # Fallback in case of an unexpected format
                specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vrwm_val)



    vcl_val = _get_value_by_keywords(part_row, ["clamping voltage"])
    if pd.notna(vcl_val):
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(vcl_val):g} V"
    
    surge_val = _get_value_by_keywords(part_row, ["ippm"])
    if pd.notna(surge_val) and str(surge_val) != '-':
        try:
            specs_result["IEC 61000-4-5"] = f"{float(surge_val):g} A (8/20µs)"
        except (ValueError, TypeError):
            pass
    
    esd_val = _get_value_by_keywords(part_row, ["vesd"])
    if pd.notna(esd_val):
        try:
            specs_result["IEC 61000-4-2"] = f"±{float(esd_val):g} kV"
        except (ValueError, TypeError):
            specs_result["IEC 61000-4-2"] = safe_strip(esd_val)

    cap_val_str = safe_strip(_get_value_by_keywords(part_row, ["cd"]))
    if cap_val_str != "-":
        # Use regex to find the first numeric value in the string
        numeric_match = re.search(r'([\d.]+)', cap_val_str)
        if numeric_match:
            try:
                numeric_val = float(numeric_match.group(1))
                specs_result["Capacitance"] = f"{numeric_val:.2f} pF"
            except (ValueError, TypeError):
                # If conversion fails, use the original string
                specs_result["Capacitance"] = f"{cap_val_str} pF"

    return specs_result

def fetch_nexperia_specs_from_excel(part_number, nexperia_dfs):
    """
    Searches across Nexperia DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS).
    """
    if part_number is None or not nexperia_dfs:
        return None

    part_row = None
    found_df_name = ""

    # Prioritized search order: Automotive Zener first, then standard Zener, then others
    search_priority = [
        "Auto_Zener", "Zener", "Auto_ESD", "Auto_TVS"
        "ESD", "TVS"
    ]
    
    all_df_keys = list(nexperia_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    for df_name in ordered_search_keys:
        df = nexperia_dfs[df_name]
        if df is None or df.empty:
            continue
        
        # Determine the correct part number column based on the file type
        if "more" in df_name.lower():
            part_num_col_keywords = ["manufacturer part number", "part number", "mfr part"]
        else:
            part_num_col_keywords = ["type number"]

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
        print(f"Part '{part_number}' not found in any Nexperia database.")
        return None
    
    print(f"Found specs for Part '{part_number}' in Nexperia database '{found_df_name}'.")

    
    if "more" in found_df_name.lower():
        return _parse_nexperia_digikey_tvs_row(part_row, found_df_name, part_number)
    elif "zener" in found_df_name.lower():
        return _parse_zener_row(part_row, found_df_name, part_number)
    else:
        return _parse_tvs_row(part_row, found_df_name, part_number)