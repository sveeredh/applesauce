import pandas as pd
import re
from parsing import safe_strip

# --- File: eaton_scrape.py ---
#
# Parses Eaton's protection-device parametric table. Column headers in the
# export are:
#
#   Eaton part number | Datasheet | Distributor inventory search | Website |
#   ECAD models | Product type | Size | Operating voltage | Bi/Uni |
#   VAC Operating voltage | Clamping voltage | Capacitance | ESD Withstand |
#   # of protection channels | Peak current | Discrete/Array | Spice models
#
# Header cells wrap inside Excel, so the text arrives as "Operating\nvoltage"
# or "Operatingvoltage" depending on the export. Every keyword lookup below
# therefore compares headers with all whitespace stripped out.


def _norm_header(text):
    """Collapse a header (or keyword) to lowercase with no whitespace, so
    'Operating voltage', 'Operating\\nvoltage' and 'Operatingvoltage' all
    compare equal."""
    return re.sub(r'[\s\u00a0]+', '', str(text)).lower()


def _get_value_by_keywords(row, keywords, exclude_keywords=None):
    """
    Searches a row's index (column headers) for the first match from a list of
    keywords. The search is case-insensitive and whitespace-insensitive.
    Any header containing one of `exclude_keywords` is skipped -- this is what
    keeps "Operating voltage" from being answered by "VAC Operating voltage".
    Returns the value from the matched column, or None.
    """
    excludes = [_norm_header(k) for k in (exclude_keywords or [])]
    for col_header in row.index:
        header = _norm_header(col_header)
        if any(bad in header for bad in excludes):
            continue
        for keyword in keywords:
            if _norm_header(keyword) in header:
                return row[col_header]
    return None  # Return None if no matching column is found


def _is_blank(value):
    """True for None, NaN, empty strings and Eaton's '-' placeholder."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() in ("", "-", "nan", "None", "N/A", "n/a", "--")


def _first_number(value):
    """First numeric token in a cell, as a float. None if there isn't one."""
    if _is_blank(value):
        return None
    match = re.search(r'[-+]?\d*\.?\d+', str(value).replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group(0))
    except (ValueError, TypeError):
        return None


# Eaton's "Size" column mixes surface-mount body codes (SMA, SMB, SMBF, SMC,
# SMT-4, SOD-123FL) with leaded through-hole styles ("Axial lead", "Radial
# lead"). TI's protection catalogue is entirely surface-mount, so a leaded part
# has no equivalent and is never crossed automatically. Add tokens here if
# Eaton introduces further through-hole styles.
LEADED_SIZE_PATTERN = re.compile(r'\b(axial|radial)\b', re.IGNORECASE)


def _is_leaded_package(size_value):
    """True if Eaton's Size cell names a leaded, non-surface-mount body."""
    if _is_blank(size_value):
        return False
    return bool(LEADED_SIZE_PATTERN.search(str(size_value)))


def _parse_eaton_row(part_row, source_df_name, part_number):
    """
    Parses a row from Eaton's parametric table into the standard spec dict the
    main script scores against.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": source_df_name,
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Capacitance": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Power Dissipation (Pd)": "-",
        "Channels": "1",
        "Tolerance": "-",
        "Price ($/ku)": "-",
        "Grade": "Commercial",
        "Package Info": {},
    }

    # Grade is decided solely by Eaton's "Product type" column: an entry like
    # "TVS diode - Automotive ESD" is automotive, "TVS Diode - Power" is not.
    # (Eaton's own capitalisation varies, hence the case-insensitive match.)
    product_type_raw = _get_value_by_keywords(part_row, ["Product type", "Product family"])
    if not _is_blank(product_type_raw) and re.search(r'automotive', str(product_type_raw), re.IGNORECASE):
        specs_result["Grade"] = "Automotive"

    # Package -- Eaton names the body style in "Size" (SMA / SMB / SMC /
    # SOT-23 / 0402 ...), which normalize_package already understands.
    size_raw = _get_value_by_keywords(part_row, ["Size", "Package"])
    if not _is_blank(size_raw):
        specs_result["Package"] = safe_strip(size_raw)
    specs_result["Package Info"] = {"name": specs_result["Package"], "version": "-"}

    # Direction -- the "Bi/Uni" column spells it out in full
    direction_raw = str(_get_value_by_keywords(part_row, ["Bi/Uni", "BiUni", "Direction"])).strip().lower()
    if direction_raw.startswith("uni"):
        specs_result["Direction"] = "Unidirectional"
    elif direction_raw.startswith("bi"):
        specs_result["Direction"] = "Bidirectional"

    # Channels -- "# of protection channels", falling back to Discrete/Array
    channels_val = _first_number(_get_value_by_keywords(part_row, ["# of protection channels", "protection channels", "channels"]))
    if channels_val is not None and channels_val > 0:
        specs_result["Channels"] = str(int(channels_val))
    else:
        disc_array = str(_get_value_by_keywords(part_row, ["Discrete/Array", "DiscreteArray"])).lower()
        if "discrete" in disc_array:
            specs_result["Channels"] = "1"

    # Reverse standoff -- Eaton's "Operating voltage" column, in volts. The
    # separate "VAC Operating voltage" column is an AC rating and is excluded.
    vrwm_val = _first_number(_get_value_by_keywords(
        part_row, ["Operating voltage", "Reverse Standoff", "VRWM"], exclude_keywords=["VAC"]))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    # Clamping voltage
    vcl_val = _first_number(_get_value_by_keywords(part_row, ["Clamping voltage", "Clamping"]))
    if vcl_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vcl_val:g} V"

    # Peak pulse current -> IEC 61000-4-5 surge rating
    ipp_raw = _get_value_by_keywords(part_row, ["Peak current", "Peak pulse current", "Ippm", "Ipp"])
    ipp_val = _first_number(ipp_raw)
    if ipp_val is not None:
        # Honour a stated mA unit; the column is otherwise amps.
        if re.search(r'\bma\b', str(ipp_raw), re.IGNORECASE):
            ipp_val = ipp_val / 1000.0
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20µs)"

    # ESD withstand -> IEC 61000-4-2. Eaton quotes kV; anything large enough to
    # be volts is converted so the comparison stays in kV.
    esd_raw = _get_value_by_keywords(part_row, ["ESD Withstand", "ESD"])
    esd_val = _first_number(esd_raw)
    if esd_val is not None:
        esd_text = str(esd_raw).lower()
        if "kv" not in esd_text and (esd_val > 100 or re.search(r'\d\s*v\b', esd_text)):
            esd_val = esd_val / 1000.0
        specs_result["IEC 61000-4-2"] = f"±{esd_val:g} kV"

    # Capacitance -- pF unless the cell says otherwise
    cap_raw = _get_value_by_keywords(part_row, ["Capacitance", "Cj", "Cd"])
    cap_val = _first_number(cap_raw)
    if cap_val is not None:
        cap_text = str(cap_raw).lower()
        if "nf" in cap_text:
            cap_val *= 1000.0
        elif "uf" in cap_text or "µf" in cap_text or "μf" in cap_text:
            cap_val *= 1000000.0
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    return specs_result


def fetch_eaton_specs_from_excel(part_number, eaton_df):
    """
    Finds a part in Eaton's parametric table and returns its specs.
    Mirrors the single-file competitor signature used by the main script:
    fetch_func(exact_part_name, df).
    """
    if part_number is None or eaton_df is None or eaton_df.empty:
        return None

    part_num_col_keywords = ["eaton part number", "part number", "mfr part", "manufacturer part number"]
    part_num_col = None
    for col in eaton_df.columns:
        header = _norm_header(col)
        for keyword in part_num_col_keywords:
            if _norm_header(keyword) in header:
                part_num_col = col
                break
        if part_num_col:
            break

    if part_num_col is None:
        print("Could not locate a part number column in the Eaton database.")
        return None

    part_row_series = eaton_df[
        eaton_df[part_num_col].astype(str).str.strip().str.lower() == str(part_number).strip().lower()
    ]

    if part_row_series.empty:
        print(f"Part '{part_number}' not found in the Eaton database.")
        return None

    part_row = part_row_series.iloc[0]

    # Leaded parts are rejected before anything else is read. Returning None
    # here (rather than a spec dict) is what keeps them out of the cross.
    size_raw = _get_value_by_keywords(part_row, ["Size", "Package"])
    if _is_leaded_package(size_raw):
        print(f"Part '{part_number}' has a leaded package ('{safe_strip(size_raw)}') "
              f"and cannot be crossed automatically.")
        return None

    # Eaton's "Product type" column carries the family name ("TVS Diode -
    # Power", "ESD Protection", "Zener Diode"). It is folded into the source
    # name so the main script can route zener parts to the zener comparison.
    product_type = safe_strip(_get_value_by_keywords(part_row, ["Product type", "Product family"]))
    source_df_name = "Eaton" if product_type == "-" else f"Eaton {product_type}"

    print(f"Found specs for Part '{part_number}' in Eaton database '{source_df_name}'.")

    return _parse_eaton_row(part_row, source_df_name, part_number)