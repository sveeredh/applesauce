import pandas as pd
from parsing import safe_strip

def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    The search is case-insensitive. Returns the value from the matched column.
    """
    if not hasattr(row, 'index'):
        return None
    for col_header in row.index:
        col_header_str = str(col_header).lower()
        for keyword in keywords:
            if keyword.lower() in col_header_str:
                return row[col_header]
    return None

def fetch_amazing_specs_from_excel(part_number, specs_df):
    """
    Fetches specifications for an Amazing IC part from a pre-loaded pandas DataFrame
    using flexible, keyword-based column searching.
    """
    specs_result = { "Device Name": part_number, "IEC 61000-4-5": "-", "IEC 61000-4-2": "-", "Capacitance": "-", "Channels": "-", "Package": "-", "Direction": "-", "Voltage - Reverse Standoff (Typ)": "-", "Voltage - Clamping (Max) @ Ipp": "-", "Price ($/ku)": "-", "Grade": "-" }
    
    if part_number is None or specs_df is None or specs_df.empty:
        return None
    
    specs_df.columns = specs_df.columns.str.strip()
    part_num_col_name = next((col for col in specs_df.columns if 'part number' in col.lower()), None)
    
    if not part_num_col_name:
        print("Amazing Scraper: Could not find a 'Part Number' column in the Excel file.")
        return None

    part_row_series = specs_df[specs_df[part_num_col_name].astype(str).str.strip().str.lower() == part_number.lower()]
    
    if part_row_series.empty:
        print(f"Part '{part_number}' not found in the Amazing IC Excel database.")
        return None
    
    part_row = part_row_series.iloc[0]
    print(f"Found specs for Part '{part_number}' in Excel.")

    
    specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))
    
    ch_val = _get_value_by_keywords(part_row, ["Channels"])
    if pd.notna(ch_val):
        try:
            specs_result["Channels"] = str(int(float(ch_val)))
        except (ValueError, TypeError):
            specs_result["Channels"] = safe_strip(ch_val)

    vrwm_val = _get_value_by_keywords(part_row, ["Vrwm (V)"])
    if pd.notna(vrwm_val):
        try:
            specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(vrwm_val):g} V"
        except (ValueError, TypeError):
            specs_result["Voltage - Reverse Standoff (Typ)"] = safe_strip(vrwm_val)

    vcl_val = _get_value_by_keywords(part_row, ["Vclamp (V)"])
    if pd.notna(vcl_val):
        try:
            specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(vcl_val):g} V"
        except (ValueError, TypeError):
            specs_result["Voltage - Clamping (Max) @ Ipp"] = safe_strip(vcl_val)
    
    ippm_val = _get_value_by_keywords(part_row, ["Surge Ipp", "(8/20)"])
    if pd.notna(ippm_val):
        try:
            specs_result["IEC 61000-4-5"] = f"{float(ippm_val):g} A (8/20µs)"
        except (ValueError, TypeError):
            specs_result["IEC 61000-4-5"] = safe_strip(ippm_val)

    # --- PRESERVING YOUR ORIGINAL ESD LOGIC ---
    esd_val_raw = safe_strip(_get_value_by_keywords(part_row, ["ESD Air/Contact"]))
    if esd_val_raw != "-" and isinstance(esd_val_raw, str):
        try:
            contact_val = esd_val_raw.split('/')[-1].strip()
            specs_result["IEC 61000-4-2"] = f"±{contact_val} kV"
        except IndexError:
            specs_result["IEC 61000-4-2"] = safe_strip(esd_val_raw)
    
    cap_val = _get_value_by_keywords(part_row, ["Capacitance (pF)"])
    if pd.notna(cap_val):
        try:
            specs_result["Capacitance"] = f"{float(cap_val):.2f} pF"
        except (ValueError, TypeError):
            specs_result["Capacitance"] = safe_strip(cap_val)

    return specs_result