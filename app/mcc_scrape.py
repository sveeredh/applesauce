import pandas as pd
import re
from parsing import safe_strip


def _norm_header(col_header):
    """Header text with spaces, non-breaking spaces and case removed."""
    return re.sub(r'[\s\u00a0_]+', '', str(col_header)).upper()


def _get_header_by_prefix(row, *prefixes):
    """
    First column header whose normalized text starts with one of the prefixes.
    Used where a short name would otherwise appear inside a longer header.
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
    MCC runs the parameter name and its symbol together with no separator
    ("BreakdownVoltageMin VBR(V)"), so Min/Max are told apart by token.
    """
    for tokens in token_sets:
        for col_header in row.index:
            header = _norm_header(col_header)
            if all(_norm_header(token) in header for token in tokens):
                return col_header
    return None


def _numeric_or_none(value):
    """First number in a cell, or None if the cell is blank or non-numeric."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    text = str(value).strip()
    if text in ("", "/", "-", "--", "N/A", "NA", "nan", "None"):
        return None
    numeric_match = re.search(r"[\d.]+", text)
    if not numeric_match:
        return None
    try:
        return float(numeric_match.group(0))
    except (ValueError, TypeError):
        return None


def _value_by_prefix(row, *prefixes):
    header = _get_header_by_prefix(row, *prefixes)
    return row.get(header) if header is not None else None


def _value_by_tokens(row, *token_sets):
    header = _get_header_by_tokens(row, *token_sets)
    return row.get(header) if header is not None else None


def _mcc_grade(part_row):
    """
    The Compliance column carries space-separated qualification letters
    ("A R H"), where A is the AEC-Q101 mark. It is read as a whole token so a
    stray A inside another word cannot promote a commercial part.
    """
    compliance = safe_strip(_value_by_prefix(part_row, "Compliance")).upper()
    if compliance == "-":
        return "Commercial"
    if "AEC" in compliance:
        return "Automotive"
    tokens = re.split(r'[\s,/]+', compliance)
    return "Automotive" if "A" in tokens else "Commercial"


def _mcc_direction(part_row):
    """Configuration states the directionality outright."""
    config = safe_strip(_value_by_prefix(part_row, "Configuration")).lower()
    config = re.sub(r'[^a-z]', '', config)
    if 'unidire' in config:
        return "Unidirectional"
    if 'bidire' in config:
        return "Bidirectional"
    return "-"


def _mcc_channels(part_row):
    """Number of Functions is the channel count."""
    channels = _numeric_or_none(_value_by_prefix(part_row, "Number of Functions"))
    if channels is not None and channels > 0:
        return str(int(channels))
    return "1"


def _mcc_package(part_row):
    return safe_strip(_value_by_prefix(part_row, "Package Type", "Package"))


def _mcc_esd_immunity(part_row):
    """
    The IEC 61000-4-2 column states air and contact together
    ("VESDIEC61000-4-2Air/Contact(kV)"). Where two figures are given the
    second is the contact rating, which is the one the cross compares on.
    """
    raw = safe_strip(_value_by_tokens(part_row, ("61000-4-2",), ("VESD",)))
    if raw == "-":
        return "-"
    numbers = re.findall(r"[\d.]+", raw)
    if not numbers:
        return "-"
    try:
        contact = float(numbers[1]) if len(numbers) > 1 else float(numbers[0])
    except (ValueError, TypeError):
        return "-"
    return f"±{contact:g} kV"


def _parse_mcc_protection_row(part_row, found_df_name, part_number):
    """
    Parses a row from an MCC ESD or TVS datasheet. Both sheets share the same
    header vocabulary, so one parser covers them.

    Columns: Product, Status, Compliance, Number of Functions, Configuration,
    Package Type, ReverseStandoffVoltageVRWM(V), Peak PulseCurrentIPP(A),
    Max.ClampingVoltageVC (V), JunctionCapacitanceCJ(pF),
    Peak PlusePowerDissipationPPPK (W), MaximumReverseLeakageIR (uA),
    BreakdownVoltageMin VBR(V), BreakdownVoltageMax VBR(V),
    JunctionTemperatureTj [max] (°C), VESDIEC61000-4-2Air/Contact(kV).

    A column the sheet does not carry simply stays at "-".
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _mcc_grade(part_row),
        "Direction": _mcc_direction(part_row),
        "Channels": _mcc_channels(part_row),
        "Package": _mcc_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": _mcc_esd_immunity(part_row),
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    vrwm_val = _numeric_or_none(_value_by_tokens(part_row, ("ReverseStandoff",), ("VRWM",)))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    ipp_val = _numeric_or_none(_value_by_tokens(part_row, ("PeakPulseCurrent",), ("IPP",)))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20µs)"

    vc_val = _numeric_or_none(_value_by_tokens(part_row, ("ClampingVoltage",)))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    cap_val = _numeric_or_none(_value_by_tokens(part_row, ("Capacitance",)))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    # "Peak PlusePowerDissipationPPPK (W)" - MCC's own spelling of Pulse.
    ppp_val = _numeric_or_none(_value_by_tokens(part_row, ("PowerDissipation",), ("PPPK",), ("PPPM",)))
    if ppp_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp_val:g} W"

    ir_val = _numeric_or_none(_value_by_tokens(part_row, ("ReverseLeakage",)))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def _parse_mcc_esd_row(part_row, found_df_name, part_number):
    """Parses a row from the MCC ESD datasheet."""
    return _parse_mcc_protection_row(part_row, found_df_name, part_number)


def _parse_mcc_tvs_row(part_row, found_df_name, part_number):
    """Parses a row from the MCC TVS datasheet."""
    return _parse_mcc_protection_row(part_row, found_df_name, part_number)


def _parse_mcc_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from the MCC Zener datasheet.

    Columns: Product, Status, Compliance, Number of Functions, Configuration,
    Package Type, PD(W), VZ[Nom](V), IZT(mA), IR (mA) [max]@VR, VR(V),
    ZZT(Omega)@IZT, ZZK(Omega)@IZT, IZK(mA), Tj [max](C).

    The sheet has no tolerance column and no VZ min/max pair, so Tolerance is
    left unstated rather than guessed at.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _mcc_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",  # not present in the Zener file headers
        "Package": _mcc_package(part_row),
    }

    # Vz. The token pair keeps this off VR(V), which also starts with V.
    vz_val = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "NOM")))
    if vz_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_val:g} V"

    # Power is already stated in watts on this sheet.
    pd_val = _numeric_or_none(_value_by_prefix(part_row, "PD("))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val:g} W"

    # Leakage is stated in mA here, against uA on the protection sheets.
    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR("))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val * 1000.0:g} uA"

    return specs_result


def fetch_mcc_specs_from_excel(part_number, mcc_dfs):
    """
    Searches across MCC DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS vs. ESD).
    """
    if part_number is None or not mcc_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["Zener", "ESD", "TVS"]

    all_df_keys = list(mcc_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["product", "part number", "mfr part"]

    for df_name in ordered_search_keys:
        df = mcc_dfs[df_name]
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
        print(f"Part '{part_number}' not found in any MCC database.")
        return None

    print(f"Found specs for Part '{part_number}' in MCC database '{found_df_name}'.")

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_mcc_zener_row(part_row, found_df_name, part_number)
    elif "tvs" in name_lower:
        return _parse_mcc_tvs_row(part_row, found_df_name, part_number)
    else:
        return _parse_mcc_esd_row(part_row, found_df_name, part_number)