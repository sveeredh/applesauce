import pandas as pd
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

def fetch_semtech_specs_from_excel(semtech_opn, semtech_specs_df):
    """
    Fetches specifications for a Semtech TVS part from a pre-loaded pandas DataFrame
    using flexible, keyword-based column searching based on the provided header image.
    """
    # Initialize the standardized dictionary
    specs_result = {
        "Device Name": semtech_opn,
        "Source File": "Semtech TVS",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Channels": "-",
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Price ($/ku)": "-",
        "Grade": "-"
    }

    if semtech_opn is None or semtech_specs_df is None or semtech_specs_df.empty:
        return None

    semtech_specs_df.columns = semtech_specs_df.columns.str.strip()
    # The first column is named 'Parts' in the image
    part_num_col = next((col for col in semtech_specs_df.columns if 'parts' in col.lower()), None)
    
    if not part_num_col:
        print("Semtech Scraper: Could not find a 'Parts' column in the Excel file.")
        return None

    # Find the specific row for the part number
    part_row_series = semtech_specs_df[semtech_specs_df[part_num_col].astype(str).str.strip().str.lower() == semtech_opn.lower()]
    
    if part_row_series.empty:
        print(f"Semtech Part '{semtech_opn}' not found in the Semtech Excel database.")
        return None
    
    part_row = part_row_series.iloc[0]
    print(f"Found specs for Semtech Part '{semtech_opn}' in Excel.")

    # --- Map and format specs using the keyword helper function ---

    specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))

    # Map 'Number of Lines' to 'Channels'
    ch_val = _get_value_by_keywords(part_row, ["Number of Lines"])
    if pd.notna(ch_val):
        specs_result["Channels"] = str(int(float(ch_val)))

    # Map 'Configuration Type' to 'Direction'
    dir_text = safe_strip(_get_value_by_keywords(part_row, ["Configuration Type"])).lower()
    if 'uni' in dir_text:
        specs_result["Direction"] = "Unidirectional"
    elif 'bi' in dir_text:
        specs_result["Direction"] = "Bidirectional"

    # Map 'VRWM MAX (V)' to 'Voltage - Reverse Standoff (Typ)'
    vrwm_val = _get_value_by_keywords(part_row, ["VRWM MAX (V)"])
    if pd.notna(vrwm_val):
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(vrwm_val):g} V"

    # Map 'VClamp (V) at Ipp 8/20µsec' to 'Voltage - Clamping (Max) @ Ipp'
    vcl_val = _get_value_by_keywords(part_row, ["VClamp (V) at Ipp"])
    if pd.notna(vcl_val):
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(vcl_val):g} V"
    
    # Map 'IPP (A) at 8/20µsec' to 'IEC 61000-4-5'
    ippm_val = _get_value_by_keywords(part_row, ["IPP (A) at"])
    if pd.notna(ippm_val):
        try:
            specs_result["IEC 61000-4-5"] = f"{float(ippm_val):g} A (8/20µs)"
        except (ValueError, TypeError):
            specs_result["IEC 61000-4-5"] = safe_strip(ippm_val)

    # Map 'IEC61000-4-2 (ESD) contact' to 'IEC 61000-4-2'
    esd_val_raw = safe_strip(_get_value_by_keywords(part_row, ["IEC61000-4-2 (ESD) contact"]))
    if esd_val_raw != "-":
        cleaned_val = str(esd_val_raw).lower().replace('±', '').replace('kv', '').strip()
        if cleaned_val:
            specs_result["IEC 61000-4-2"] = f"±{cleaned_val} kV"
        else:
            specs_result["IEC 61000-4-2"] = safe_strip(esd_val_raw)
    
    # Map 'Cj (Typ) (pF)' to 'Capacitance'
    cap_val = _get_value_by_keywords(part_row, ["Cj (Typ) (pF)"])
    if pd.notna(cap_val):
        try:
            specs_result["Capacitance"] = f"{float(cap_val):.2f} pF"
        except (ValueError, TypeError):
            specs_result["Capacitance"] = safe_strip(cap_val)
    if semtech_opn.lower().endswith('q'):
        specs_result["Grade"] = "Automotive"

    return specs_result