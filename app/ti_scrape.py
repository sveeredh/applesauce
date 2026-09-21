# --- START OF FILE ti_scrape.py ---

import pandas as pd
from parsing import safe_strip

# --- pull excel file ---
def load_excel_data(filepath, sheet_name=0, header_row=0):
    """
    Loads an excel file into a pandas DataFrame.
    """
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row)
        return df
    except FileNotFoundError:
        print(f"ERROR: Excel file not found: {filepath}")
    except Exception as e:
        print(f"ERROR: Could not load Excel file {filepath}: {e}")
    return None

def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    The search is case-insensitive. Returns the value from the matched column.
    """
    # Ensure row.index is accessible and iterable
    if not hasattr(row, 'index'):
        return None

    for col_header in row.index:
        # Ensure col_header is a string for the 'in' check
        col_header_str = str(col_header).lower()
        for keyword in keywords:
            if keyword.lower() in col_header_str:
                return row[col_header]
    return None

# --- read specs ---
def fetch_ti_specs_from_excel(ti_opn, ti_specs_df, param_keys_config):
    """
    Fetches specifications for a TI part, preserving original user logic
    while using flexible keyword-based column searching.
    """
    ti_specs_result = {key_spec[0]: "-" for key_spec in param_keys_config}
    if ti_opn is None or ti_specs_df is None or ti_specs_df.empty:
        return ti_specs_result

    ti_specs_df.columns = ti_specs_df.columns.str.strip()
    
    # Find the part number column by keyword
    part_num_col_name = next((col for col in ti_specs_df.columns if 'product or part number' in col.lower()), None)
    if not part_num_col_name:
        print("ERROR: 'Product or Part number' column not found in TI specs Excel.")
        return ti_specs_result
        
    part_row_series = ti_specs_df[ti_specs_df[part_num_col_name].astype(str).str.strip().str.lower() == ti_opn.lower()]
    
    if part_row_series.empty:
        print(f"TI Part '{ti_opn}' not found in TI specs database.")
        return ti_specs_result
        
    part_row = part_row_series.iloc[0]
    print(f"Found specs for TI Part '{ti_opn}'.")

    # --- Map and format specs using keywords while preserving original logic ---
    
    device_name = safe_strip(_get_value_by_keywords(part_row, ["Product or Part number"]))
    ti_specs_result["Device Name"] = device_name
    
    if device_name.endswith("-Q1"):
        ti_specs_result["Grade"] = "Automotive"
    else:
        ti_specs_result["Grade"] = "-"

    ppp_val = _get_value_by_keywords(part_row, ["IEC 61000-4-5 (A)"])
    if pd.notna(ppp_val) and str(ppp_val).strip() != '':
        try:
            ti_specs_result["IEC 61000-4-5"] = f"{float(ppp_val):g} A (8/20µs)"
        except ValueError:
            ti_specs_result["IEC 61000-4-5"] = safe_strip(str(ppp_val))

    # --- PRESERVING YOUR ORIGINAL ESD LOGIC ---
    contact_kv = _get_value_by_keywords(part_row, ["IEC 61000-4-2 contact"])
    if pd.notna(contact_kv) and str(contact_kv).strip() != '': 
        try:
            # Your original logic: divide by 1000
            esd_in_kv = round(float(contact_kv) / 1000.0, None) 
            ti_specs_result["IEC 61000-4-2"] = f"±{esd_in_kv:g} kV"
        except (ValueError, TypeError):
             ti_specs_result["IEC 61000-4-2"] = safe_strip(str(contact_kv))


    cap_val = _get_value_by_keywords(part_row, ["IO capacitance (typ) (pF)"])
    if pd.notna(cap_val) and str(cap_val).strip() != '':
        try:
            ti_specs_result["Capacitance"] = f"{float(cap_val):.2f} pF"
        except ValueError:
             ti_specs_result["Capacitance"] = f"{safe_strip(str(cap_val))} pF"

    pin_count_val = _get_value_by_keywords(part_row, ["Pin count"])
    if pd.notna(pin_count_val):
        try:
            ti_specs_result["Pin count"] = int(float(pin_count_val))
        except (ValueError, TypeError):
            ti_specs_result["Pin count"] = safe_strip(pin_count_val)

    direction_text_ti = safe_strip(_get_value_by_keywords(part_row, ["-directional", "Directionality"]))
    if "uni-directional" in direction_text_ti.lower():
        ti_specs_result["Direction"] = "Unidirectional"
    elif "bi-directional" in direction_text_ti.lower():
        ti_specs_result["Direction"] = "Bidirectional"
    else:
        ti_specs_result["Direction"] = direction_text_ti

    num_channels_val = _get_value_by_keywords(part_row, ["Number of channels"])
    if pd.notna(num_channels_val) and safe_strip(str(num_channels_val)) != '-':
        try:
            ti_specs_result["Channels"] = str(int(float(num_channels_val)))
        except ValueError:
            ti_specs_result["Channels"] = safe_strip(str(num_channels_val))
    elif ti_specs_result.get("Direction", "-") in ["Unidirectional", "Bidirectional"]:
        ti_specs_result["Channels"] = "1" 
    else:
        ti_specs_result["Channels"] = "-"

    ti_specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package name"]))

    vrwm_val = _get_value_by_keywords(part_row, ["Vrwm (V)"])
    if pd.notna(vrwm_val) and str(vrwm_val).strip() != '':
        try:
            ti_specs_result["Voltage - Reverse Standoff (Typ)"] = f"{float(vrwm_val):g} V"
        except ValueError:
            ti_specs_result["Voltage - Reverse Standoff (Typ)"] = f"{safe_strip(str(vrwm_val))} V"

    vcl_val = _get_value_by_keywords(part_row, ["Clamping voltage (V)"])
    if pd.notna(vcl_val) and str(vcl_val).strip() != '':
        try:
            ti_specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{float(vcl_val):g} V"
        except ValueError:
            ti_specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{safe_strip(str(vcl_val))} V"

    price_val = _get_value_by_keywords(part_row, ["Price|Quantity (USD)"])
    if pd.notna(price_val) and str(price_val).strip() != '':
        try:
            ti_specs_result["Price ($/ku)"] = f"${float(price_val):.5f}"
        except ValueError:
            ti_specs_result["Price ($/ku)"] = safe_strip(str(price_val))
            
    return ti_specs_result