import pandas as pd
from parsing import safe_strip

def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    """
    if not hasattr(row, 'index'):
        return None
    for col_header in row.index:
        col_header_str = str(col_header).lower()
        for keyword in keywords:
            if keyword.lower() in col_header_str:
                return row[col_header]
    return None

def fetch_ti_zener_specs_from_excel(ti_opn, ti_zener_specs_df, param_keys_config):
    """
    Fetches Zener-specific specifications for a TI part from the TI Zener database,
    using the correct headers from the TI spec file.
    """
    ti_specs_result = {key_spec[0]: "-" for key_spec in param_keys_config}
    if ti_opn is None or ti_zener_specs_df is None or ti_zener_specs_df.empty:
        return ti_specs_result

    ti_zener_specs_df.columns = ti_zener_specs_df.columns.str.strip()
    
    part_num_col_name = next((col for col in ti_zener_specs_df.columns if 'product or part number' in col.lower()), None)
    if not part_num_col_name:
        print("ERROR: 'Product or Part number' column not found in TI Zener specs Excel.")
        return ti_specs_result
        
    part_row_series = ti_zener_specs_df[ti_zener_specs_df[part_num_col_name].astype(str).str.strip().str.lower() == ti_opn.lower()]
    
    if part_row_series.empty:
        print(f"TI Part '{ti_opn}' not found in TI Zener specs database.")
        return ti_specs_result
        
    part_row = part_row_series.iloc[0]
    print(f"Found Zener specs for TI Part '{ti_opn}'.")

    # Fetch Zener-specific parameters using corrected headers
    ti_specs_result["Device Name"] = safe_strip(_get_value_by_keywords(part_row, ["part number"]))
    ti_specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))
    
    price_val = _get_value_by_keywords(part_row, ["Price"])
    if pd.notna(price_val):
        try:
            ti_specs_result["Price ($/ku)"] = f"${float(price_val):.5f}"
        except (ValueError, TypeError):
            ti_specs_result["Price ($/ku)"] = safe_strip(price_val)


    pd_val = _get_value_by_keywords(part_row, ["pd (max)", "power"])
    if pd.notna(pd_val):
        ti_specs_result["Power Dissipation (Pd)"] = f"{safe_strip(pd_val)} W"

    tol_val = _get_value_by_keywords(part_row, ["Tolerance (±) (%)"])
    if pd.notna(tol_val):
        ti_specs_result["Tolerance"] = f"±{safe_strip(tol_val)}%"

    vz_val = _get_value_by_keywords(part_row, ["Vz (nom) (V)", "Voltage"])
    if pd.notna(vz_val):
        # This parameter is also aliased for comparison purposes in the main script
        ti_specs_result["Zener Voltage (Vz)"] = f"{safe_strip(vz_val)} V"
        ti_specs_result["Voltage - Reverse Standoff (Typ)"] = f"{safe_strip(vz_val)} V"


    return ti_specs_result