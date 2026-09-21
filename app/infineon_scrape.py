"""
Infineon ESD / TVS parametric table.

There is no scraping here - the parametric export is downloaded by hand from
Infineon's product finder and dropped in next to this file. This module only
reads that workbook and turns a row of it into the spec dictionary the cross
reference expects.
"""

import re

import pandas as pd

from parsing import safe_strip


SPECS_FILE = "infineon_specs.xlsx"


# ============================================================
# SPEC COLUMNS
#
# The export's header names, kept as constants so the parsers
# below read the columns by name instead of by position.
#
# Note that the part is keyed on OPN - the orderable part
# number - rather than on the "Part number" column, which
# carries the truncated marketing name.
# ============================================================

PART_NUMBER_COLUMN = "Part number"

PART_LINK_COLUMN = "Part link"

PACKAGE_LINK_COLUMN = "Package link"

OPN_COLUMN = "OPN"

PRODUCT_STATUS_COLUMN = "Product status"

PACKAGE_COLUMN = "Infineon package"

GREEN_COLUMN = "Green"

PROTECTED_LINES_COLUMN = "Protected Lines"

CONFIGURATION_COLUMN = "Configuration"

CAPACITANCE_COLUMN = "CL"

CAPACITANCE_MAX_COLUMN = "CL max"

STANDOFF_MIN_COLUMN = "VWM min"

STANDOFF_MAX_COLUMN = "VWM max"

TARGET_APPLICATION_COLUMN = "Target Application"

LEAKAGE_CURRENT_COLUMN = "IR"

LEAKAGE_CURRENT_MAX_COLUMN = "IR max"

EFT_CURRENT_COLUMN = "IEFT"

PEAK_PULSE_CURRENT_COLUMN = "IPP"

# Columns that sit off the right of the visible export. Each is
# looked up leniently and simply left out when absent, so
# adding the real header name here is all that is needed.

CLAMPING_COLUMNS = [
    "VCL",
    "VCL max",
    "VC",
    "Clamping voltage",
]

BREAKDOWN_COLUMNS = [
    "VBR",
    "VBR min",
    "Breakdown voltage",
]

ESD_CONTACT_COLUMNS = [
    "VESD contact",
    "ESD contact",
    "VESD (contact)",
    "IEC 61000-4-2 contact",
]

ESD_AIR_COLUMNS = [
    "VESD air",
    "ESD air",
    "VESD (air)",
    "IEC 61000-4-2 air",
]

POWER_COLUMNS = [
    "PPP",
    "PPPM",
    "Peak pulse power",
]


# ============================================================
# SPEC HELPERS
# ============================================================

def normalize_header(
    column_name
):

    # --------------------------------------------------------
    # Headers arrive with stray double spaces ("CL  max"), so
    # every comparison runs on a whitespace-collapsed,
    # lowercased form.
    # --------------------------------------------------------

    return re.sub(
        r"\s+",
        " ",
        str(column_name),
    ).strip().lower()


def get_cell(
    part_row,
    column_name,
):

    if column_name in part_row.index:
        return part_row[column_name]

    wanted = normalize_header(
        column_name
    )

    for column in part_row.index:

        if normalize_header(column) == wanted:
            return part_row[column]

    return None


def get_text(
    part_row,
    column_name,
):

    value = get_cell(
        part_row,
        column_name,
    )

    if value is None or (
        not isinstance(value, str)
        and pd.isna(value)
    ):
        return ""

    return safe_strip(value)


def get_number(
    part_row,
    column_name,
):

    # --------------------------------------------------------
    # Cells carry their unit ("0.13 pF", "-5.5 V", "20 nA") and
    # occasionally two values ("-5.5 V, -5.5 V"); the first
    # number is the one that counts.
    # --------------------------------------------------------

    value = get_cell(
        part_row,
        column_name,
    )

    if value is None or (
        not isinstance(value, str)
        and pd.isna(value)
    ):
        return None

    numeric_match = re.search(
        r"-?\d+(?:\.\d+)?",
        str(value),
    )

    if not numeric_match:
        return None

    try:
        return float(
            numeric_match.group(0)
        )

    except (ValueError, TypeError):
        return None


def get_first_number(
    part_row,
    column_names,
):

    for column_name in column_names:

        value = get_number(
            part_row,
            column_name,
        )

        if value is not None:
            return value

    return None


def get_grade(
    part_row
):

    # --------------------------------------------------------
    # The export states no grade column. Infineon flags its
    # qualified parts in the application text instead.
    # --------------------------------------------------------

    text = " ".join([
        get_text(part_row, TARGET_APPLICATION_COLUMN),
        get_text(part_row, PRODUCT_STATUS_COLUMN),
        get_text(part_row, OPN_COLUMN),
    ]).lower()

    if "aec" in text or "automotive" in text:
        return "Automotive"

    return "Non-Automotive"


def get_direction(
    part_row
):

    # --------------------------------------------------------
    # Held in Configuration, written out in full.
    # --------------------------------------------------------

    configuration = get_text(
        part_row,
        CONFIGURATION_COLUMN,
    ).upper().replace("-", " ")

    if "BI" in configuration:
        return "Bidirectional"

    if "UNI" in configuration:
        return "Unidirectional"

    return "-"


def get_channels(
    part_row
):

    # --------------------------------------------------------
    # Protected Lines is the channel count.
    # --------------------------------------------------------

    lines = get_number(
        part_row,
        PROTECTED_LINES_COLUMN,
    )

    if lines is None:
        return "1"

    return f"{int(lines)}"


# ============================================================
# PACKAGE NAMES
#
# Infineon names its bodies its own way, so each one is
# translated to the name the package matcher works in.
#
# Matching runs on the letters and digits alone, with the
# PG-/SG- prefix dropped, so any hyphenation of the same code
# lands on the same package: "PG-TSLP-2-20", "PGTSLP220" and
# "TSLP 2 20" are one and the same.
# ============================================================

INFINEON_PACKAGE_MAP = {
    "TSLP220": "DFN1006-2",     # 2 pin DFN1006
    "TSLP219": "DFN1006-2",
    "TSSLP23": "DFN0603-2",     # 2 pin DFN0603
    "TSSLP24": "DFN0603-2",
    "TSLP37": "SOT-883-3",      # SOT-883, which resolves on to DFN1006 3 pin
    "TSLP47": "4-XFDFN",
    "TSNP22": "DFN1608-2",      # 2 pin DFN1608
    "WLL3": "SOT-723",
    "TSFP31": "SOT-723",
    "SOT143410": "SOT-23-4",
}

# A family whose number ends at the 2 - nothing hyphenated
# after it - is the 2 pin DFN0603.
INFINEON_TWO_PIN = re.compile(r"^[A-Z]+2$")


def normalize_infineon_package(
    package_string
):

    raw = str(package_string or "").strip().upper()

    if not raw:
        return ""

    flat = re.sub(
        r"[\W_]+",
        "",
        raw,
    )

    flat = re.sub(
        r"^(?:PG|SG)(?=[A-Z])",
        "",
        flat,
    )

    if flat in INFINEON_PACKAGE_MAP:
        return INFINEON_PACKAGE_MAP[flat]

    if INFINEON_TWO_PIN.match(flat):
        return "DFN0603-2"

    # No rule for this body - pass Infineon's own name through.

    return str(package_string).strip()


# ============================================================
# PARSE ROW
# ============================================================

def parse_infineon_row(
    part_row,
    part_number,
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": "Infineon ESD",
        "Grade": get_grade(part_row),
        "Direction": get_direction(part_row),
        "Channels": get_channels(part_row),
        "Package": normalize_infineon_package(
            get_text(part_row, PACKAGE_COLUMN)
        ),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Breakdown (Min)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-",
        "IEC 61000-4-5": "-",
        "Capacitance": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    # VWM is stated as a window. The positive end is the
    # standoff; a bidirectional part's negative end mirrors it,
    # so its magnitude stands in when the maximum is blank.

    standoff = get_number(
        part_row,
        STANDOFF_MAX_COLUMN,
    )

    if standoff is None:

        standoff_min = get_number(
            part_row,
            STANDOFF_MIN_COLUMN,
        )

        if standoff_min is not None:
            standoff = abs(standoff_min)

    if standoff is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{standoff:g} V"

    breakdown = get_first_number(
        part_row,
        BREAKDOWN_COLUMNS,
    )

    if breakdown is not None:
        specs_result["Voltage - Breakdown (Min)"] = f"{breakdown:g} V"

    clamping = get_first_number(
        part_row,
        CLAMPING_COLUMNS,
    )

    if clamping is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamping:g} V"

    # Contact discharge is the 61000-4-2 figure; air discharge
    # is the fallback.

    esd_level = get_first_number(
        part_row,
        ESD_CONTACT_COLUMNS,
    )

    if esd_level is None:
        esd_level = get_first_number(
            part_row,
            ESD_AIR_COLUMNS,
        )

    if esd_level is not None:
        specs_result["IEC 61000-4-2"] = f"{esd_level:g} kV"

    # IPP is the 8/20 us surge rating - the 61000-4-5 figure.
    # IEFT is the 5/50 ns fast transient one, which nothing
    # downstream compares against, so it is left out.

    surge_current = get_number(
        part_row,
        PEAK_PULSE_CURRENT_COLUMN,
    )

    if surge_current is not None:
        specs_result["IEC 61000-4-5"] = f"{surge_current:g} A"

    power = get_first_number(
        part_row,
        POWER_COLUMNS,
    )

    if power is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power:g} W"

    # CL is the typical line capacitance; CL max stands in when
    # only the limit is stated.

    capacitance = get_number(
        part_row,
        CAPACITANCE_COLUMN,
    )

    if capacitance is None:
        capacitance = get_number(
            part_row,
            CAPACITANCE_MAX_COLUMN,
        )

    if capacitance is not None:
        specs_result["Capacitance"] = f"{capacitance:.2f} pF"

    return specs_result


# ============================================================
# FETCH SPECS FROM EXCEL
#
# infineon_df is the single parametric sheet.
# ============================================================

def fetch_infineon_specs_from_excel(
    part_number,
    infineon_df,
):

    if not part_number or infineon_df is None or (
        hasattr(infineon_df, "empty")
        and infineon_df.empty
    ):
        return None

    df = infineon_df

    df.columns = df.columns.str.strip()

    part_column = next(
        (
            column
            for column in df.columns
            if normalize_header(column) == OPN_COLUMN.lower()
        ),
        None,
    )

    if part_column is None:

        print(
            f"No '{OPN_COLUMN}' column in the "
            f"Infineon database."
        )

        return None

    target = str(part_number).strip().lower()

    target_loose = re.sub(
        r"[\W_]+",
        "",
        target,
    )

    column_values = df[part_column].astype(str).str.strip()

    part_rows = df[column_values.str.lower() == target]

    if part_rows.empty:

        loose_values = column_values.str.lower().str.replace(
            r"[\W_]+",
            "",
            regex=True,
        )

        part_rows = df[loose_values == target_loose]

    # The OPN carries a packing suffix the marketing name does
    # not, so a request for the base part matches the OPN it
    # starts with.

    if part_rows.empty:

        loose_values = column_values.str.lower().str.replace(
            r"[\W_]+",
            "",
            regex=True,
        )

        part_rows = df[loose_values.str.startswith(target_loose, na=False)]

    if part_rows.empty:

        print(
            f"Part '{part_number}' not found "
            f"in the Infineon database."
        )

        return None

    part_row = part_rows.iloc[0]

    exact_name = safe_strip(
        part_row[part_column]
    )

    print(
        f"Found specs for Part '{part_number}' "
        f"in the Infineon database."
    )

    return parse_infineon_row(
        part_row,
        exact_name,
    )