import re
import pandas as pd


def _clean_text(value):
    if value is None or pd.isna(value):
        return "-"

    text = str(value).strip()

    if text.lower() in ("", "-", "nan", "none"):
        return "-"

    return text


def _num(value):
    if value is None or pd.isna(value):
        return None

    match = re.search(
        r"[-+]?\d*\.?\d+",
        str(value)
    )

    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


def _fmt(value):
    if value is None:
        return "-"

    return f"{value:g}"


def _find_part_column(df):
    if df is None or df.empty:
        return None

    for col in df.columns:
        if str(col).strip().lower() == "part number":
            return col

    return None


def _find_exact_row(df, part_number):
    part_col = _find_part_column(df)

    if part_col is None:
        return None

    wanted = re.sub(
        r"[\W_]+",
        "",
        str(part_number).upper()
    )

    normalized = (
        df[part_col]
        .astype(str)
        .str.upper()
        .str.replace(
            r"[\W_]+",
            "",
            regex=True
        )
    )

    matches = df[
        normalized == wanted
    ]

    if matches.empty:
        return None

    return matches.iloc[0]


# ============================================================
# COMMON PANJIT FIELDS
# ============================================================

def _grade_from_row(row):
    """
    Panjit sheets expose AEC-Q101 Qualified as Y / -.
    """
    aec = _clean_text(
        row.get(
            "AEC-Q101 Qualified",
            "-"
        )
    ).upper()

    if aec in (
        "Y",
        "YES",
        "TRUE",
        "1"
    ):
        return "Automotive"

    return "Commercial"


def _direction(value):
    text = _clean_text(
        value
    ).upper()

    if text == "-":
        return "-"

    if (
        text.startswith("BI")
        or "BIDIRECTION" in text
    ):
        return "Bidirectional"

    if (
        text.startswith("UNI")
        or "UNIDIRECTION" in text
    ):
        return "Unidirectional"

    return text


# ============================================================
# ESD
# ============================================================

def _parse_panjit_esd_row(
    row,
    part_number,
):
    """
    panjit_esd_specs.xls

    Relevant columns from Panjit:
        Part Number
        Package
        Product Status
        Replacement Part
        AEC-Q101 Qualified
        Configuration
        Number of protected lines
        VRWM Max.
        VBR Min.
        IR @ VRWM Max.
        VC @ IPP Max.
        IPP
        CJ Max.
        CJ Typ.
    """

    vrwm = _num(
        row.get("VRWM Max.")
    )

    vbr = _num(
        row.get("VBR Min.")
    )

    clamp = _num(
        row.get("VC @ IPP Max.")
    )

    ipp = _num(
        row.get("IPP")
    )

    cj_max = _num(
        row.get("CJ Max.")
    )

    cj_typ = _num(
        row.get("CJ Typ.")
    )

    lines = _num(
        row.get(
            "Number of protected lines"
        )
    )

    # Prefer typical capacitance when Panjit provides it.
    # Otherwise use maximum.
    capacitance = (
        cj_typ
        if cj_typ is not None
        else cj_max
    )

    result = {
        "Device Name":
            str(part_number).strip(),

        "Source File":
            "panjit_esd_specs.xls",

        "Grade":
            _grade_from_row(row),

        "Package":
            _clean_text(
                row.get("Package")
            ),

        "Direction":
            _direction(
                row.get("Configuration")
            ),

        "Channels":
            (
                str(int(lines))
                if lines is not None
                else "1"
            ),

        "Voltage - Reverse Standoff (Typ)":
            (
                f"{_fmt(vrwm)} V"
                if vrwm is not None
                else "-"
            ),

        "Voltage - Clamping (Max) @ Ipp":
            (
                f"{_fmt(clamp)} V"
                if clamp is not None
                else "-"
            ),

        "Capacitance":
            (
                f"{_fmt(capacitance)} pF"
                if capacitance is not None
                else "-"
            ),

        # No IEC values are visible in the supplied
        # Panjit ESD columns, so do not invent them.
        "IEC 61000-4-5":
            (
                f"{_fmt(ipp)} A"
                if ipp is not None
                else "-"
            ),

        "IEC 61000-4-2":
            "-",

        "Power Dissipation (Pd)":
            "-",
    }

    if vbr is not None:
        result[
            "Breakdown Voltage Min"
        ] = f"{_fmt(vbr)} V"

    if ipp is not None:
        result[
            "Peak Pulse Current"
        ] = f"{_fmt(ipp)} A"

    return result


# ============================================================
# TVS
# ============================================================

def _parse_panjit_tvs_row(
    row,
    part_number,
):
    """
    panjit_tvs_specs.xls

    Relevant columns:
        Part Number
        Package
        Product Status
        Replacement Part
        AEC-Q101 Qualified
        Configuration
        PPP
        VRWM
        VBR @ IT Min.
        VBR @ IT Max.
        IT
        IR @ VRWM UNI
        IR @ VRWM BI
        VC @ IPP Max.
        IPP
    """

    vrwm = _num(
        row.get("VRWM")
    )

    vbr_min = _num(
        row.get("VBR @ IT Min.")
    )

    vbr_max = _num(
        row.get("VBR @ IT Max.")
    )

    ppp = _num(
        row.get("PPP")
    )

    clamp = _num(
        row.get("VC @ IPP Max.")
    )

    ipp = _num(
        row.get("IPP")
    )

    result = {
        "Device Name":
            str(part_number).strip(),

        "Source File":
            "panjit_tvs_specs.xls",

        "Grade":
            _grade_from_row(row),

        "Package":
            _clean_text(
                row.get("Package")
            ),

        "Direction":
            _direction(
                row.get("Configuration")
            ),

        "Channels":
            "1",

        "Voltage - Reverse Standoff (Typ)":
            (
                f"{_fmt(vrwm)} V"
                if vrwm is not None
                else "-"
            ),

        "Voltage - Clamping (Max) @ Ipp":
            (
                f"{_fmt(clamp)} V"
                if clamp is not None
                else "-"
            ),

        "Capacitance":
            "-",

        "IEC 61000-4-5":
            (
                f"{_fmt(ipp)} A"
                if ipp is not None
                else "-"
            ),

        "IEC 61000-4-2":
            "-",

        "Power Dissipation (Pd)":
            (
                f"{_fmt(ppp)} W"
                if ppp is not None
                else "-"
            ),
    }

    if vbr_min is not None:
        result[
            "Breakdown Voltage Min"
        ] = (
            f"{_fmt(vbr_min)} V"
        )

    if vbr_max is not None:
        result[
            "Breakdown Voltage Max"
        ] = (
            f"{_fmt(vbr_max)} V"
        )

    if ipp is not None:
        result[
            "Peak Pulse Current"
        ] = (
            f"{_fmt(ipp)} A"
        )

    return result


# ============================================================
# ZENER
# ============================================================

def _parse_panjit_zener_row(
    row,
    part_number,
):
    """
    panjit_zener_specs.xls

    Relevant columns:
        Part Number
        Package
        Product Status
        Replacement Part
        AEC-Q101 Qualified
        PD
        VZ
        VZ @ IZT Nom.
        VZ @ IZT Min.
        VZ @ IZT Max.
        ZZT @ IZT Max.
        IZT
        ZZK @ IZK Max.
        IZK
        IR @ VR Max.   [uA]
        IR @ VR Max.   [V]
        Circuit Figure
    """

    vz_nom = _num(
        row.get("VZ @ IZT Nom.")
    )

    vz_min = _num(
        row.get("VZ @ IZT Min.")
    )

    vz_max = _num(
        row.get("VZ @ IZT Max.")
    )

    tolerance = _num(
        row.get("VZ")
    )

    # Panjit gives PD in mW.
    pd_mw = _num(
        row.get("PD")
    )

    pd_w = (
        pd_mw / 1000.0
        if pd_mw is not None
        else None
    )

    result = {
        "Device Name":
            str(part_number).strip(),

        "Source File":
            "panjit_zener_specs.xls",

        "Grade":
            _grade_from_row(row),

        "Package":
            _clean_text(
                row.get("Package")
            ),

        "Voltage - Reverse Standoff (Typ)":
            (
                f"{_fmt(vz_nom)} V"
                if vz_nom is not None
                else "-"
            ),

        "Tolerance":
            (
                f"{_fmt(tolerance)}%"
                if tolerance is not None
                else "-"
            ),

        "Power Dissipation (Pd)":
            (
                f"{_fmt(pd_w)} W"
                if pd_w is not None
                else "-"
            ),

        "Direction":
            "-",

        "Channels":
            "1",
    }

    if vz_min is not None:
        result[
            "Zener Voltage Min"
        ] = f"{_fmt(vz_min)} V"

    if vz_max is not None:
        result[
            "Zener Voltage Max"
        ] = f"{_fmt(vz_max)} V"

    return result


# ============================================================
# PUBLIC FETCH FUNCTION
# ============================================================

def fetch_panjit_specs_from_excel(
    part_number,
    panjit_df_map,
):
    """
    Search only:
        panjit_esd_specs.xls
        panjit_tvs_specs.xls
        panjit_zener_specs.xls
    """

    if (
        part_number is None
        or not panjit_df_map
    ):
        return None

    for category in (
        "ESD",
        "TVS",
        "Zener",
    ):
        df = panjit_df_map.get(
            category
        )

        if (
            df is None
            or df.empty
        ):
            continue

        row = _find_exact_row(
            df,
            part_number
        )

        if row is None:
            continue

        part_col = (
            _find_part_column(df)
        )

        exact_part = _clean_text(
            row.get(
                part_col,
                part_number
            )
        )

        if category == "ESD":
            return _parse_panjit_esd_row(
                row,
                exact_part
            )

        if category == "TVS":
            return _parse_panjit_tvs_row(
                row,
                exact_part
            )

        return _parse_panjit_zener_row(
            row,
            exact_part
        )

    print(
        f"Part '{part_number}' not found in "
        "panjit_esd_specs.xls, "
        "panjit_tvs_specs.xls, or "
        "panjit_zener_specs.xls."
    )

    return None