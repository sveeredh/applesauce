import pandas as pd
import re
from parsing import safe_strip


# Littelfuse array sheets carry no Vrwm and no channel count, so both are
# derived from the part number. The 4-digit code after the prefix indexes
# this table; anything not listed falls through to the structural rules.
_LF_CODE_VRWM = {
    "1001": 5.5, "4339": 4.5, "0115": 1.0, "4340": 4.0, "4337L": 1.5,
    "3118": 18.0, "3130": 28.0, "3022": 5.3, "3025": 2.5, "3530": 7.0,
    "4022": 15.0, "1230": 30.0, "1236": 36.0, "1248": 48.0, "4024": 36.0,
    "4315": 1.5, "1115": 15.0, "4208": 8.0, "1224": 22.0, "3000": 30.0,
    "1064": 60.0,
}
for _c in ("0502", "1105", "1210", "1205", "1003", "3041", "3045", "3052",
           "3102", "3400", "3522", "4020", "4337", "8008", "4322", "4338",
           "3422", "3423", "3213", "3030", "3031", "3006", "3010", "3011",
           "3019", "3021", "1326", "1533", "1250", "1305", "3402", "7538",
           "1007", "1008", "1012", "1015"):
    _LF_CODE_VRWM[_c] = 5.0
for _c in ("1002", "1005", "1006", "3051", "1020", "3001", "3002", "3003",
           "3004", "1026"):
    _LF_CODE_VRWM[_c] = 6.0
for _c in ("4021", "1112", "1312"):
    _LF_CODE_VRWM[_c] = 12.0
for _c in ("4023", "4324", "1124"):
    _LF_CODE_VRWM[_c] = 24.0
for _c in ("2525", "2555"):
    _LF_CODE_VRWM[_c] = 2.5
for _c in ("7520", "1333", "1103", "1233", "2502", "3205", "3222", "3304",
           "3312", "3384", "3374", "3401", "3420", "4042", "4045", "4065",
           "4203"):
    _LF_CODE_VRWM[_c] = 3.3
for _c in ("2504", "4044", "4060"):
    _LF_CODE_VRWM[_c] = 2.8
for _c in ("1120", "4320"):
    _LF_CODE_VRWM[_c] = 20.0
for _c in ("1122", "2200"):
    _LF_CODE_VRWM[_c] = 22.0

# Short codes that appear on SC and SM parts instead of a 4-digit code.
_LF_SHORT_VRWM = {"3V3": 3.3, "5V0": 5.0, "712": 12.0}

# Whole part numbers that follow none of the structural rules.
_LF_PART_OVERRIDES = {
    "SP03-3.3BTG": (3.3, 2),
    "SP03-6BTG":   (6.0, 2),
    "SP33R6-04UTG": (0.3, 4),
    "SP712-02HTG": (12.0, 2),
}

# Prefixes whose Vrwm is fixed regardless of the digits.
_LF_FIXED_VRWM = {"SESD": 7.0, "LC": 3.3}
# Prefixes whose channel count is fixed regardless of the digits.
_LF_FIXED_CHANNELS = {"LC": 2, "TPSMF": 2}

# Longest first, so SPHV is not read as SP and SDW is not read as SD.
_LF_PREFIXES = ("SRDA", "TPSMF", "SLVU", "SPHV", "SCHV", "AQHV", "SESD",
                "SDW", "SRV", "AQ", "SC", "SP", "SD", "SM", "SR", "LC")
def _get_value_by_keywords(row, keywords):
    """
    Searches a row's index (column headers) for the first match from a list of keywords.
    The search is case-insensitive. Returns the value from the matched column.
    """
    for col_header in row.index:
        for keyword in keywords:
            if keyword.lower() in str(col_header).lower():
                val = row[col_header]
                return val if pd.notna(val) else "-"
    return "-" # Return "-" if no matching column is found

def _lf_channels_from_pn(part_number):
    """Channel count from the part number. None when nothing matches."""
    pn = str(part_number).strip().upper()

    if pn in _LF_PART_OVERRIDES:
        return _LF_PART_OVERRIDES[pn][1]

    for prefix, channels in _LF_FIXED_CHANNELS.items():
        if pn.startswith(prefix):
            return channels

    # The number after the first hyphen: AQ1003-01LTG -> 1.
    hyphen_match = re.match(r'^[A-Z]+[^-]*-(\d+)', pn)
    if hyphen_match:
        return int(hyphen_match.group(1))

    # SESD0402Q2UG -> the digit after the Q.
    q_match = re.search(r'Q(\d+)', pn)
    if q_match:
        return int(q_match.group(1))

    # No hyphen at all (SLVU2.8HTG) means a single channel.
    return 1


def _lf_vrwm_from_pn(part_number):
    """Vrwm in volts from the part number. None when nothing matches."""
    pn = str(part_number).strip().upper()

    if pn in _LF_PART_OVERRIDES:
        return _LF_PART_OVERRIDES[pn][0]

    for prefix, vrwm in _LF_FIXED_VRWM.items():
        if pn.startswith(prefix):
            return vrwm

    prefix = next((p for p in _LF_PREFIXES if pn.startswith(p)), None)
    if prefix is None:
        return None

    body = pn[len(prefix):]
    core = body.split("-")[0]
    has_hyphen = "-" in body

    # These state the voltage directly after the prefix (SLVU2.8HTG, SM712).
    if prefix in ("SLVU", "TPSMF", "SCHV", "SD", "SDW", "SM"):
        short = re.match(r'^(3V3|5V0|712)(?=[A-Z\-]|$)', core)
        if short:
            return _LF_SHORT_VRWM[short.group(1)]
        number = re.match(r'^(\d+(?:\.\d+)?)', core)
        return float(number.group(1)) if number else None

    # SC and SM short codes.
    short = re.match(r'^(3V3|5V0|712)(?=[A-Z\-]|$)', core)
    if short:
        return _LF_SHORT_VRWM[short.group(1)]

    # No other SP7xx part can be crossed.
    if prefix in ("SP", "SPHV") and re.match(r'^7\d{2}(?=[A-Z\-]|$)', core):
        return None

    # Exactly two digits before letters or a hyphen: the digits are the
    # voltage (AQHV24-01ETG, AQ12CANA-02HTG).
    two_digit = re.match(r'^(\d{2})(?=[A-Z\-]|$)', core)
    if two_digit:
        return float(two_digit.group(1))

    # A four-digit code. With no hyphen an SP part splits as xxyy, where xx is
    # the voltage and yy the channel count (SP0502BAHT = 5 V, 2 channels).
    # Everywhere else the code indexes the table.
    code = re.match(r'^(\d{4})(L?)', core)
    if code:
        if prefix in ("SP", "SPHV") and not has_hyphen:
            return float(code.group(1)[:2])
        with_l = code.group(1) + code.group(2)
        if with_l in _LF_CODE_VRWM:
            return _LF_CODE_VRWM[with_l]
        return _LF_CODE_VRWM.get(code.group(1))

    return None

def _lf_grade(part_row, found_df_name):
    """
    The array sheet states qualification per part in an AEC-Q101 column, so
    the grade is read off the row. The other sheets are split by grade at the
    file level, so there the filename is what states it.
    """
    aec = safe_strip(_get_value_by_keywords(part_row, ["AEC-Q101", "AEC Q101", "AEC"]))
    if aec != "-":
        return "Automotive" if aec.strip().lower().startswith("y") else "Commercial"
    return "Automotive" if "auto" in found_df_name.lower() else "Commercial"

def _parse_littelfuse_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Littelfuse CSV datasheet to extract TVS/ESD parameters.
    """
    specs_result = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": _lf_grade(part_row, found_df_name),
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "IEC 61000-4-5": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Channels": "-",
    }


    # --- Map columns based on screenshot and expected values ---

    # Package (Simple, no version)
    specs_result["Package"] = safe_strip(_get_value_by_keywords(part_row, ["Package"]))

    # Standoff Voltage (VR)
    # Standoff Voltage (VR). The array sheets have no such column, so it is
    # derived from the part number instead.
    vr_val = _get_value_by_keywords(part_row, ["Vstandoff", "Operating"])
    if vr_val != "-":
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vr_val} V"
    elif "array" in found_df_name.lower():
        derived_vrwm = _lf_vrwm_from_pn(part_number)
        if derived_vrwm is not None:
            specs_result["Voltage - Reverse Standoff (Typ)"] = f"{derived_vrwm:g} V"


    # Directionality
    direction_str = safe_strip(_get_value_by_keywords(part_row, ["Uni / Bi-Directional", "Polarity"])).lower()
    if "uni" in direction_str:
        specs_result["Direction"] = "Unidirectional"
    elif "bi" in direction_str:
        specs_result["Direction"] = "Bidirectional"
    
    # Surge (Ipp)
    surge_val = _get_value_by_keywords(part_row, ["I PP 8x20µs (A)"])
    if surge_val != "-":
        specs_result["IEC 61000-4-5"] = f"{surge_val} A (8/20µs)"

    # esd
    esd_val_raw = _get_value_by_keywords(part_row, ["ESD Contact"])

    # Check for valid, non-placeholder input.
    if pd.notna(esd_val_raw) and esd_val_raw != "-":
        # Use regex to find the first valid number (integer or decimal).
        numeric_match = re.search(r"([\d.]+)", str(esd_val_raw))
        
        if numeric_match:
            try:
                # Extract the number, convert it to a float.
                numeric_val = float(numeric_match.group(1))
                # Now, build the final string with the clean number.
                specs_result["IEC 61000-4-2"] = f"±{numeric_val:g} kV"
            except (ValueError, IndexError):
                # If conversion fails for any reason, use a placeholder.
                specs_result["IEC 61000-4-2"] = "-"
        else:
            # If no number was found, it's not a valid value.
            specs_result["IEC 61000-4-2"] = "-"
    else:
        # If the initial value is a placeholder, keep it that way.
        specs_result["IEC 61000-4-2"] = "-"

    # --- Handle other potential columns not in the screenshot ---
    # These use the keyword search to be flexible across different CSV files.
    
    clamping_val = _get_value_by_keywords(part_row, ["Vc (V)", "Clamping"])
    if clamping_val != "-":
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamping_val} V"
    
    cap_val_raw = _get_value_by_keywords(part_row, ["Co Typ (pF)", "CI/O TYP (pF)"])

    # Check for valid, non-placeholder input before processing
    if pd.notna(cap_val_raw) and cap_val_raw != "-":
        # Convert raw value to string to handle both numbers and text formats
        s_val = str(cap_val_raw)
        
        # Regex to find the first valid number (integer or decimal) in the string
        numeric_match = re.search(r"([\d.]+)", s_val)
        
        if numeric_match:
            try:
                # Extract the number, convert to float, and format consistently
                numeric_val = float(numeric_match.group(1))
                specs_result["Capacitance"] = f"{numeric_val:.2f} pF"
            except (ValueError, IndexError):
                # If conversion fails for any reason, use a placeholder
                specs_result["Capacitance"] = "-"
        else:
            # If no number is found, it's not a valid capacitance value
            specs_result["Capacitance"] = "-"
    else:
        # If the initial value is a placeholder, keep it
        specs_result["Capacitance"] = "-"
        
    channels_val = _get_value_by_keywords(part_row, ["Channels"])
    if channels_val != "-":
        specs_result["Channels"] = str(int(float(channels_val)))
    elif "array" in found_df_name.lower():
        derived_channels = _lf_channels_from_pn(part_number)
        if derived_channels is not None:
            specs_result["Channels"] = str(derived_channels)

    return specs_result

def fetch_littelfuse_specs_from_excel(part_number, df_map):
    """
    Searches for a Littelfuse part number across multiple CSVs with a specific priority.
    """
    if part_number is None or not df_map:
        return None

    # Use the search priority requested by the user
    search_priority = [
        "Auto_TVS", "TVS",
        "TVS_Array"
    ]

    all_df_keys = list(df_map.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])
    
    for df_name in ordered_search_keys:
        df = df_map.get(df_name)
        if df is None or df.empty:
            continue
        
        part_num_col = next((col for col in df.columns if 'part number' in str(col).lower()), None)
        if not part_num_col:
            continue
        
        match_row_series = df[df[part_num_col].astype(str).str.strip().str.lower().str.startswith(part_number.lower())]
        
        if not match_row_series.empty:
            part_row = match_row_series.iloc[0]
            print(f"Found specs for Part '{part_number}' in Littelfuse database '{df_name}'.")
            # There are no Zeners for Littelfuse, so we directly call the standard parser.
            return _parse_littelfuse_row(part_row, df_name, part_number)

    print(f"Part '{part_number}' not found in any local Littelfuse files.")
    return None