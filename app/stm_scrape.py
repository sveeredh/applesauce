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
            # Use a word boundary to ensure we match whole words, e.g., 'Package' not 'Packaging'
            if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', str(col_header).lower()):
                val = row[col_header]
                return val if pd.notna(val) else "-"
    return "-" # Return "-" if no matching column is found

def _parse_stm_row(part_row, source_df_name, part_number):
    """
    Parses a row from the STMicroelectronics CSV datasheet to extract key parameters.
    """
    specs_result = {
        "Device Name": part_number,
        "Source File": source_df_name,
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "IEC 61000-4-5": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-", # Not present in the STMicroelectronics CSV header
        "Capacitance": "-",
        "Channels": "1", # Default to 1
    }

    # Grade
    qualification_str = str(_get_value_by_keywords(part_row, ["Qualification"])).lower()
    if 'aec-q' in qualification_str:
        specs_result["Grade"] = "Automotive"

    # Package (Prioritize 'Supplier Device Package' as it's more specific)
    pkg_supplier = safe_strip(_get_value_by_keywords(part_row, ["Package / Case"]))
    if pkg_supplier != "-":
        specs_result["Package"] = pkg_supplier
    else:
        # Fallback to 'Package / Case'
        specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Supplier Package"]))

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

    # Voltage - Reverse Standoff (Typ)
    vr_val_raw = _get_value_by_keywords(part_row, ["Voltage - Reverse"])
    if pd.notna(vr_val_raw) and vr_val_raw != "-":
        numeric_match = re.search(r"([\d.]+)", str(vr_val_raw))
        if numeric_match:
            specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(numeric_match.group(1)):g} V"

    # Voltage - Clamping (Max) @ Ipp
    clamping_val_raw = _get_value_by_keywords(part_row, ["Voltage - Clamping"])
    if pd.notna(clamping_val_raw) and clamping_val_raw != "-":
        numeric_match = re.search(r"([\d.]+)", str(clamping_val_raw))
        if numeric_match:
            specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(numeric_match.group(1)):g} V"
        
    # Surge (Ipp)
    surge_val_raw = _get_value_by_keywords(part_row, ["Peak Pulse"])

    # Check for valid, non-placeholder input before processing
    if pd.notna(surge_val_raw) and surge_val_raw != "-":
        # Use regex to find the first valid number (integer or decimal) in the string
        numeric_match = re.search(r"([\d.]+)", str(surge_val_raw))
        
        if numeric_match:
            try:
                # Extract the number, convert to float for validation
                numeric_val = float(numeric_match.group(1))
                # Manually append the units after stripping everything else
                specs_result["IEC 61000-4-5"] = f"{numeric_val:g} A (8/20µs)"
            except (ValueError, IndexError):
                # If conversion fails for any reason, the default placeholder "-" remains
                pass
        
    # Capacitance
    cap_str = safe_strip(_get_value_by_keywords(part_row, ["Capacitance"]))
    cap_match = re.search(r'([\d.]+)\s*pF', cap_str, re.IGNORECASE)
    if cap_match:
        specs_result["Capacitance"] = f"{float(cap_match.group(1)):.2f} pF"

    return specs_result

def fetch_stm_specs_from_excel(part_number, stm_specs_df):

    if part_number is None or stm_specs_df is None or stm_specs_df.empty:
        return None

    # --- Find Manufacturer Part Number Column ---
    mfr_part_num_keywords = ["manufacturer part number", "part number", "mfr part #"]
    mfr_part_num_col = None
    for col in stm_specs_df.columns:
        for keyword in mfr_part_num_keywords:
            if keyword in str(col).lower():
                mfr_part_num_col = col
                break
        if mfr_part_num_col:
            break
    
    # --- First Attempt: Search by Manufacturer Part Number (Exact Match) ---
    if mfr_part_num_col:
        match_row_series = stm_specs_df[stm_specs_df[mfr_part_num_col].astype(str).str.strip().str.lower() == part_number.lower()]
        
        if not match_row_series.empty:
            part_row = match_row_series.iloc[0]
            matched_part_number = part_row[mfr_part_num_col]
            return _parse_stm_row(part_row, "stmicroelectronics_specs.csv", matched_part_number)

    # --- Fallback: Find and Search by DK Part Number (Contains Match) ---
    dk_part_num_col = None
    for col in stm_specs_df.columns:
        if "dk part" in str(col).lower():
            dk_part_num_col = col
            break
            
    if dk_part_num_col:
        
        # Create a regex pattern to find the part number as a whole word.
        # re.escape() handles cases where part numbers have special characters (e.g., '+', '.').
        # \b ensures we don't match substrings (e.g., '123' inside 'A1234').
        search_pattern = r'\b' + re.escape(part_number) + r'\b'

        # Use str.contains() with the regex pattern.
        # na=False ensures that empty/NaN cells are treated as non-matches.
        # case=False makes the search case-insensitive.
        match_row_series_dk = stm_specs_df[
            stm_specs_df[dk_part_num_col].astype(str).str.contains(search_pattern, case=False, na=False, regex=True)
        ]

        if not match_row_series_dk.empty:
            part_row = match_row_series_dk.iloc[0]
            # When found by DK P/N, we still display the manufacturer part number for clarity
            display_part_number = part_row.get(mfr_part_num_col, part_number)
            return _parse_stm_row(part_row, "stmicroelectronics_specs.csv", display_part_number)

    # If both searches fail, print the final message
    print(f"Part '{part_number}' not found in the local STMicroelectronics file by Mfr P/N or DK P/N.")
    return None