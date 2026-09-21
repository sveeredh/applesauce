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

def _clean_numeric(value):
    """
    Returns the first numeric value from an Onsemi cell.
    Handles cells such as:
        '3.3,'
        '0.225,'
        '-,'
    """
    if value is None or pd.isna(value):
        return None

    match = re.search(r"[-+]?\d*\.?\d+", str(value))

    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


def _clean_text(value):
    if value is None or pd.isna(value):
        return "-"

    text = str(value).strip()

    # Onsemi exports have trailing commas in many cells.
    text = text.rstrip(",").strip()

    if text.lower() in ("", "-", "-,", "nan", "none"):
        return "-"

    return text


def _parse_onsemi_esd_row(part_row, part_number):
    """
    Parses onsemi_esd_specs.xlsx.

    Expected headers:
        Product Group
        Status
        Compliance
        Interface
        Number of Lines
        Direction
        C Max (pF)
        V(BR) Min (V)
        VRWM Max (V)
        IR Max (µA)
        PPK Max (W)
        Package Size (mm)
        Package Type
    """
    package_type = _clean_text(
        part_row.get("Package Type", "-")
    )

    if "axial" in package_type.lower():
        return None
    specs_result = {
        "Device Name": str(part_number).strip(),
        "Source File": "onsemi_esd_specs.csv",
        "Grade": (
            "Automotive"
            if str(part_number).strip().upper().startswith("SZ")
            else "Commercial"
        ),

        "Package": "-",
        "Direction": "-",

        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",

        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",

        "Capacitance": "-",
        "Channels": "1",

        "Power Dissipation (Pd)": "-",
    }

    # Package
    package = _clean_text(
        part_row.get("Package Type", "-")
    )

    if package != "-":
        specs_result["Package"] = package

    # Direction
    direction = _clean_text(
        part_row.get("Direction", "-")
    )

    if direction != "-":
        if "bi" in direction.lower():
            specs_result["Direction"] = "Bidirectional"
        elif "uni" in direction.lower():
            specs_result["Direction"] = "Unidirectional"

    # Number of protected lines -> Channels
    lines = _clean_numeric(
        part_row.get("Number of Lines")
    )

    if lines is not None and lines > 0:
        specs_result["Channels"] = str(
            int(lines)
        )

    # VRWM Max
    vrwm = _clean_numeric(
        part_row.get("VRWM Max (V)")
    )

    if vrwm is not None:
        specs_result[
            "Voltage - Reverse Standoff (Typ)"
        ] = f"{vrwm:g} V"

    # Capacitance
    capacitance = _clean_numeric(
        part_row.get("C Max (pF)")
    )

    if capacitance is not None:
        specs_result[
            "Capacitance"
        ] = f"{capacitance:g} pF"

    # Peak pulse power
    ppk = _clean_numeric(
        part_row.get("PPK Max (W)")
    )

    if ppk is not None:
        specs_result[
            "Power Dissipation (Pd)"
        ] = f"{ppk:g} W"

    # Optional fields retained for reference
    vbr = _clean_numeric(
        part_row.get("V(BR) Min (V)")
    )

    if vbr is not None:
        specs_result[
            "Breakdown Voltage Min"
        ] = f"{vbr:g} V"

    leakage = _clean_numeric(
        part_row.get("IR Max (µA)")
    )

    if leakage is not None:
        specs_result[
            "Leakage Current"
        ] = f"{leakage:g} uA"

    return specs_result


def _parse_onsemi_zener_row(part_row, part_number):
    """
    Parses onsemi_zener_specs.xlsx.

    Expected headers:
        Product Group
        Status
        Compliance
        VZ Typ (V)
        P Max (W)
        Package Size (mm)
        Package Type
        MSL Type
        MSL Temp
        Reference Price
    """
    package_type = _clean_text(
        part_row.get("Package Type", "-")
    )

    if "axial" in package_type.lower():
        return None
    specs_result = {
        "Device Name": str(part_number).strip(),
        "Source File": "onsemi_zener_specs.csv",
        "Grade": (
            "Automotive"
            if str(part_number).strip().upper().startswith("SZ")
            else "Commercial"
        ),

        "Package": "-",

        # applesauce uses this key as Zener Vz
        "Voltage - Reverse Standoff (Typ)": "-",

        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",

        "Direction": "-",
        "Channels": "1",
    }

    package = _clean_text(
        part_row.get("Package Type", "-")
    )

    if package != "-":
        specs_result["Package"] = package

    # VZ Typ
    vz = _clean_numeric(
        part_row.get("VZ Typ (V)")
    )

    if vz is not None:
        specs_result[
            "Voltage - Reverse Standoff (Typ)"
        ] = f"{vz:g} V"

    # P Max
    power = _clean_numeric(
        part_row.get("P Max (W)")
    )

    if power is not None:
        specs_result[
            "Power Dissipation (Pd)"
        ] = f"{power:g} W"

    # Reference price
    price = _clean_numeric(
        part_row.get("Reference Price")
    )

    if price is not None:
        specs_result[
            "Price ($/ku)"
        ] = price

    return specs_result


def fetch_onsemi_specs_from_excel(
    part_number,
    onsemi_df_map,
):
    """
    Searches only:
        onsemi_esd_specs.xlsx
        onsemi_zener_specs.xlsx
    """

    if (
        part_number is None
        or not onsemi_df_map
    ):
        return None

    norm_part = re.sub(
        r"[\W_]+",
        "",
        str(part_number).upper()
    )

    if not norm_part:
        return None

    # ESD first, then Zener.
    for category in ("ESD", "Zener"):
        df = onsemi_df_map.get(category)

        if df is None or df.empty:
            continue

        # Your new Onsemi files use Product Group as the part number.
        part_col = next(
            (
                col
                for col in df.columns
                if str(col).strip().lower()
                == "product group"
            ),
            None,
        )

        if part_col is None:
            continue

        normalized_column = (
            df[part_col]
            .astype(str)
            .str.upper()
            .str.replace(
                r"[\W_]+",
                "",
                regex=True,
            )
        )

        matches = df[
            normalized_column
            == norm_part
        ]

        if matches.empty:
            continue

        row = matches.iloc[0]

        exact_part = _clean_text(
            row.get(
                part_col,
                part_number,
            )
        )

        if category == "ESD":
            specs = _parse_onsemi_esd_row(
                row,
                exact_part,
            )
        else:
            specs = _parse_onsemi_zener_row(
                row,
                exact_part,
            )

        if specs is None:
            print(
                f"Skipping Onsemi part '{exact_part}' "
                f"because Package Type contains 'Axial'."
            )
            return None

        return specs

    print(
        f"Part '{part_number}' not found in "
        "onsemi_esd_specs.xlsx or "
        "onsemi_zener_specs.xlsx."
    )

    return None