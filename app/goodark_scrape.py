import re
import traceback
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font

from parsing import safe_strip


TVS_URL = "https://goodarksemi.com/transient-voltage-suppressors/"
ESD_URL = "https://goodarksemi.com/esd-transient-voltage-suppressors/"
ZENER_URL = "https://goodarksemi.com/zener-diodes/"

TVS_OUTPUT = "goodark_tvs_specs.xlsx"
ESD_OUTPUT = "goodark_esd_specs.xlsx"
ZENER_OUTPUT = "goodark_zener_specs.xlsx"

TIMEOUT = 30

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Upgrade-Insecure-Requests": "1",
}


TVS_COLUMNS = [
    "Part Number",
    "Data Sheet",
    "Package",
    "Direction",
    "Pppm (W)",
    "VBR Min (V)",
    "VBR Max (V)",
    "VBR Nom (V)",
    "VWM (V)",
]


ESD_COLUMNS = [
    "Part Number",
    "Data Sheet",
    "Package",
    "Direction",
    "Configuration",
    "VRWM (V)",
    "IR Max (uA)",
    "VBR Min (V)",
    "CJ Typ (pF)",
    "PPP (W)",
    "IPP (A)",
    "VESD Contact (kV)",
    "VESD Air (kV)",
    "VC Typ (V)",
    "VC Max (V)",
    "@ IPP (A)",
]


ZENER_COLUMNS = [
    "Part Number",
    "Data Sheet",
    "Package",
    "PTOT (W)",
    "VZ Nom (V)",
    "Tolerance (%)",
    "VZ Min (V)",
    "VZ Max (V)",
    "IZT (mA)",
]


def clean_text(value):
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def extract_datasheet_url(td, page_url):
    if td is None:
        return ""

    for a in td.find_all(
        "a",
        href=True,
    ):
        href = clean_text(
            a.get("href")
        )

        if not href:
            continue

        if ".pdf" in href.lower():
            return urljoin(
                page_url,
                href,
            )

    return ""


def looks_like_part_number(text):
    text = clean_text(text)

    if not text:
        return False

    lowered = text.lower()

    if lowered in {
        "part number",
        "search",
        "clear all filters",
    }:
        return False

    return (
        bool(
            re.search(
                r"[A-Za-z]",
                text,
            )
        )
        and bool(
            re.search(
                r"\d",
                text,
            )
        )
    )


# ============================================================
# TVS PARSER
# ============================================================

def parse_tvs_row(tr):
    """
    Good-Ark TVS table:

      0 Part Number
      1 Data Sheet
      2 Package
      3 Direction
      4 Pppm
      5 VBR min
      6 VBR max
      7 VBR nom
      8 VWM
    """

    tds = tr.find_all(
        "td",
        recursive=False,
    )

    if len(tds) < 9:
        tds = tr.find_all("td")

    if len(tds) < 9:
        return None

    part_number = clean_text(
        tds[0].get_text(
            " ",
            strip=True,
        )
    )

    if not looks_like_part_number(
        part_number
    ):
        return None

    return {
        "Part Number":
            part_number,

        "Data Sheet":
            extract_datasheet_url(
                tds[1],
                TVS_URL,
            ),

        "Package":
            clean_text(
                tds[2].get_text(
                    " ",
                    strip=True,
                )
            ),

        "Direction":
            clean_text(
                tds[3].get_text(
                    " ",
                    strip=True,
                )
            ),

        "Pppm (W)":
            clean_text(
                tds[4].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VBR Min (V)":
            clean_text(
                tds[5].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VBR Max (V)":
            clean_text(
                tds[6].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VBR Nom (V)":
            clean_text(
                tds[7].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VWM (V)":
            clean_text(
                tds[8].get_text(
                    " ",
                    strip=True,
                )
            ),
    }


# ============================================================
# ESD PARSER
# ============================================================

def parse_esd_row(tr):
    """
    Good-Ark ESD TVS table:

      0  Part Number
      1  Data Sheet
      2  Package
      3  Direction
      4  Configuration
      5  VRWM
      6  IR max
      7  VBR min
      8  CJ typ
      9  PPP
      10 IPP
      11 VESD Contact
      12 VESD Air
      13 VC Typ
      14 VC Max
      15 @ IPP
    """

    tds = tr.find_all(
        "td",
        recursive=False,
    )

    if len(tds) < 16:
        tds = tr.find_all("td")

    if len(tds) < 16:
        return None

    part_number = clean_text(
        tds[0].get_text(
            " ",
            strip=True,
        )
    )

    if not looks_like_part_number(
        part_number
    ):
        return None

    return {
        "Part Number":
            part_number,

        "Data Sheet":
            extract_datasheet_url(
                tds[1],
                ESD_URL,
            ),

        "Package":
            clean_text(
                tds[2].get_text(
                    " ",
                    strip=True,
                )
            ),

        "Direction":
            clean_text(
                tds[3].get_text(
                    " ",
                    strip=True,
                )
            ),

        "Configuration":
            clean_text(
                tds[4].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VRWM (V)":
            clean_text(
                tds[5].get_text(
                    " ",
                    strip=True,
                )
            ),

        "IR Max (uA)":
            clean_text(
                tds[6].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VBR Min (V)":
            clean_text(
                tds[7].get_text(
                    " ",
                    strip=True,
                )
            ),

        "CJ Typ (pF)":
            clean_text(
                tds[8].get_text(
                    " ",
                    strip=True,
                )
            ),

        "PPP (W)":
            clean_text(
                tds[9].get_text(
                    " ",
                    strip=True,
                )
            ),

        "IPP (A)":
            clean_text(
                tds[10].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VESD Contact (kV)":
            clean_text(
                tds[11].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VESD Air (kV)":
            clean_text(
                tds[12].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VC Typ (V)":
            clean_text(
                tds[13].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VC Max (V)":
            clean_text(
                tds[14].get_text(
                    " ",
                    strip=True,
                )
            ),

        "@ IPP (A)":
            clean_text(
                tds[15].get_text(
                    " ",
                    strip=True,
                )
            ),
    }


# ============================================================
# ZENER PARSER
# ============================================================

def parse_zener_row(tr):
    """
    Good-Ark Zener table:

      0 Part Number
      1 Data Sheet
      2 Package
      3 PTOT
      4 VZ nom
      5 Tolerance
      6 VZ Min
      7 VZ Max
      8 IZT
    """

    tds = tr.find_all(
        "td",
        recursive=False,
    )

    if len(tds) < 9:
        tds = tr.find_all("td")

    if len(tds) < 9:
        return None

    part_number = clean_text(
        tds[0].get_text(
            " ",
            strip=True,
        )
    )

    if not looks_like_part_number(
        part_number
    ):
        return None

    return {
        "Part Number":
            part_number,

        "Data Sheet":
            extract_datasheet_url(
                tds[1],
                ZENER_URL,
            ),

        "Package":
            clean_text(
                tds[2].get_text(
                    " ",
                    strip=True,
                )
            ),

        "PTOT (W)":
            clean_text(
                tds[3].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VZ Nom (V)":
            clean_text(
                tds[4].get_text(
                    " ",
                    strip=True,
                )
            ),

        "Tolerance (%)":
            clean_text(
                tds[5].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VZ Min (V)":
            clean_text(
                tds[6].get_text(
                    " ",
                    strip=True,
                )
            ),

        "VZ Max (V)":
            clean_text(
                tds[7].get_text(
                    " ",
                    strip=True,
                )
            ),

        "IZT (mA)":
            clean_text(
                tds[8].get_text(
                    " ",
                    strip=True,
                )
            ),
    }


# ============================================================
# SHARED SCRAPER
# ============================================================

def scrape_catalog(
    session,
    catalog_name,
    url,
    output_file,
    output_columns,
    parser,
):
    print()
    print("=" * 70)
    print(
        f"GOOD-ARK {catalog_name} SCRAPER"
    )
    print("=" * 70)

    print(
        f"URL: {url}"
    )

    print(
        f"Output: {output_file}"
    )

    print()
    print(
        "Downloading page..."
    )

    headers = dict(
        HEADERS
    )

    headers["Referer"] = url

    response = session.get(
        url,
        headers=headers,
        timeout=TIMEOUT,
    )

    response.raise_for_status()

    if not response.encoding:
        response.encoding = (
            "utf-8"
        )

    soup = BeautifulSoup(
        response.text,
        "html.parser",
    )

    records = []

    for tr in soup.find_all(
        "tr"
    ):
        record = parser(
            tr
        )

        if record:
            records.append(
                record
            )

    if not records:
        raise RuntimeError(
            f"No Good-Ark "
            f"{catalog_name} rows "
            f"were parsed."
        )

    print(
        f"Found "
        f"{len(records)} "
        f"raw product rows."
    )

    print()
    print(
        "First parsed row:"
    )

    for key in output_columns:
        print(
            f"  {key}: "
            f"{records[0].get(key, '')}"
        )

    df = pd.DataFrame(
        records
    )

    for column in output_columns:
        if column not in df.columns:
            df[column] = ""

    df = df[
        output_columns
    ]

    df = (
        df
        .drop_duplicates(
            subset=[
                "Part Number"
            ],
            keep="first",
        )
        .reset_index(
            drop=True
        )
        .fillna("")
    )

    print()
    print(
        f"Unique products: "
        f"{len(df)}"
    )

    print(
        f"Saving to "
        f"{output_file}..."
    )

    df.to_excel(
        output_file,
        index=False,
    )

    format_excel(
        output_file,
        output_columns,
        f"Good-Ark {catalog_name}",
    )

    print()
    print(
        f"Finished "
        f"{catalog_name}."
    )

    print(
        f"Saved {len(df)} rows."
    )

    return df


# ============================================================
# EXCEL FORMATTING
# ============================================================

def format_excel(
    filename,
    output_columns,
    sheet_name,
):
    wb = load_workbook(
        filename
    )

    ws = wb.active

    ws.title = (
        sheet_name[:31]
    )

    ws.freeze_panes = (
        "A2"
    )

    ws.auto_filter.ref = (
        ws.dimensions
    )

    for cell in ws[1]:
        cell.font = Font(
            bold=True
        )

    for column_cells in ws.columns:
        letter = (
            column_cells[0]
            .column_letter
        )

        max_length = 0

        for cell in column_cells:
            value = (
                ""
                if cell.value is None
                else str(
                    cell.value
                )
            )

            max_length = max(
                max_length,
                len(value),
            )

        if letter == "B":
            width = min(
                max(
                    max_length + 2,
                    20,
                ),
                70,
            )

        else:
            width = min(
                max(
                    max_length + 2,
                    12,
                ),
                30,
            )

        ws.column_dimensions[
            letter
        ].width = width

    if "Data Sheet" in output_columns:
        datasheet_column = (
            output_columns.index(
                "Data Sheet"
            )
            + 1
        )

        for row in range(
            2,
            ws.max_row + 1,
        ):
            cell = ws.cell(
                row=row,
                column=datasheet_column,
            )

            if (
                isinstance(
                    cell.value,
                    str,
                )
                and cell.value.startswith(
                    "http"
                )
            ):
                cell.hyperlink = (
                    cell.value
                )

                cell.style = (
                    "Hyperlink"
                )

    wb.save(
        filename
    )


# ============================================================
# MAIN
# ============================================================

def main():
    print(
        "Good-Ark Semiconductor scraper"
    )

    print("=" * 70)

    print(
        "No Selenium/browser is used."
    )

    print(
        "The cookie banner does not "
        "affect direct HTML scraping."
    )

    session = (
        requests.Session()
    )

    # --------------------------------------------------------
    # TVS
    # --------------------------------------------------------

    tvs_df = scrape_catalog(
        session=session,

        catalog_name="TVS",

        url=TVS_URL,

        output_file=TVS_OUTPUT,

        output_columns=TVS_COLUMNS,

        parser=parse_tvs_row,
    )

    # --------------------------------------------------------
    # ESD
    # --------------------------------------------------------

    esd_df = scrape_catalog(
        session=session,

        catalog_name="ESD",

        url=ESD_URL,

        output_file=ESD_OUTPUT,

        output_columns=ESD_COLUMNS,

        parser=parse_esd_row,
    )

    # --------------------------------------------------------
    # ZENER
    # --------------------------------------------------------

    zener_df = scrape_catalog(
        session=session,

        catalog_name="Zener",

        url=ZENER_URL,

        output_file=ZENER_OUTPUT,

        output_columns=ZENER_COLUMNS,

        parser=parse_zener_row,
    )

    print()
    print("=" * 70)
    print(
        "ALL GOOD-ARK SCRAPES COMPLETE"
    )
    print("=" * 70)

    print(
        f"{TVS_OUTPUT}: "
        f"{len(tvs_df)} rows"
    )

    print(
        f"{ESD_OUTPUT}: "
        f"{len(esd_df)} rows"
    )

    print(
        f"{ZENER_OUTPUT}: "
        f"{len(zener_df)} rows"
    )


def main():
    print("Good-Ark Semiconductor scraper")
    print("=" * 70)
    print("No Selenium/browser is used.")
    print("The cookie banner does not affect direct HTML scraping.")

    session = requests.Session()

    # TVS
    tvs_df = scrape_catalog(
        session=session,
        catalog_name="TVS",
        url=TVS_URL,
        output_file=TVS_OUTPUT,
        output_columns=TVS_COLUMNS,
        parser=parse_tvs_row,
    )

    # ESD
    esd_df = scrape_catalog(
        session=session,
        catalog_name="ESD",
        url=ESD_URL,
        output_file=ESD_OUTPUT,
        output_columns=ESD_COLUMNS,
        parser=parse_esd_row,
    )

    # ZENER
    zener_df = scrape_catalog(
        session=session,
        catalog_name="Zener",
        url=ZENER_URL,
        output_file=ZENER_OUTPUT,
        output_columns=ZENER_COLUMNS,
        parser=parse_zener_row,
    )

    print()
    print("=" * 70)
    print("ALL GOOD-ARK SCRAPES COMPLETE")
    print("=" * 70)

    print(f"{TVS_OUTPUT}: {len(tvs_df)} rows")
    print(f"{ESD_OUTPUT}: {len(esd_df)} rows")
    print(f"{ZENER_OUTPUT}: {len(zener_df)} rows")

# ============================================================
# SPEC COLUMNS
#
# The saved sheets' header names, kept as constants so the
# parsers below read the columns by name instead of by
# position. They match TVS_COLUMNS / ESD_COLUMNS /
# ZENER_COLUMNS exactly.
# ============================================================

PART_COLUMN = "Part Number"

DATASHEET_COLUMN = "Data Sheet"

PACKAGE_COLUMN = "Package"

DIRECTION_COLUMN = "Direction"

# --- TVS ---

PEAK_PULSE_POWER_COLUMN = "Pppm (W)"

BREAKDOWN_MIN_COLUMN = "VBR Min (V)"

BREAKDOWN_MAX_COLUMN = "VBR Max (V)"

BREAKDOWN_NOM_COLUMN = "VBR Nom (V)"

WORKING_VOLTAGE_COLUMN = "VWM (V)"

# --- ESD ---

CONFIGURATION_COLUMN = "Configuration"

STANDOFF_COLUMN = "VRWM (V)"

LEAKAGE_CURRENT_COLUMN = "IR Max (uA)"

CAPACITANCE_COLUMN = "CJ Typ (pF)"

ESD_PEAK_PULSE_POWER_COLUMN = "PPP (W)"

PEAK_PULSE_CURRENT_COLUMN = "IPP (A)"

ESD_CONTACT_COLUMN = "VESD Contact (kV)"

ESD_AIR_COLUMN = "VESD Air (kV)"

CLAMPING_TYP_COLUMN = "VC Typ (V)"

CLAMPING_MAX_COLUMN = "VC Max (V)"

CLAMPING_CURRENT_COLUMN = "@ IPP (A)"

# --- ZENER ---

POWER_DISSIPATION_COLUMN = "PTOT (W)"

ZENER_NOM_COLUMN = "VZ Nom (V)"

TOLERANCE_COLUMN = "Tolerance (%)"

ZENER_MIN_COLUMN = "VZ Min (V)"

ZENER_MAX_COLUMN = "VZ Max (V)"

ZENER_TEST_CURRENT_COLUMN = "IZT (mA)"


# ============================================================
# SPEC HELPERS
# ============================================================

def get_cell(
    part_row,
    column_name,
):

    if column_name in part_row.index:
        return part_row[column_name]

    wanted = str(column_name).strip().lower()

    for column in part_row.index:

        if str(column).strip().lower() == wanted:
            return part_row[column]

    return None


def get_number(
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
        return None

    numeric_match = re.search(
        r"[\d.]+",
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


def get_grade(
    part_row
):

    # --------------------------------------------------------
    # Good-Ark states no grade column, so the only marker is
    # an AEC/automotive note in the datasheet path.
    # --------------------------------------------------------

    datasheet = safe_strip(
        get_cell(part_row, DATASHEET_COLUMN)
    ).lower()

    if "aec" in datasheet or "automotive" in datasheet:
        return "Automotive"

    return "Non-Automotive"


def get_direction(
    part_row
):

    # --------------------------------------------------------
    # Written "Uni-Dir" / "Bi-Dir" in both the TVS and the ESD
    # table.
    # --------------------------------------------------------

    direction = safe_strip(
        get_cell(part_row, DIRECTION_COLUMN)
    ).upper().replace("-", " ")

    if direction.startswith("BI"):
        return "Bidirectional"

    if direction.startswith("UNI"):
        return "Unidirectional"

    return "-"


def get_channels(
    part_row
):

    # --------------------------------------------------------
    # Configuration reads "Single", "Array", "Dual" and the
    # like. Only a stated count or an unmistakable single
    # channel is used - "Array" gives no number, so it is left
    # unknown rather than guessed at.
    # --------------------------------------------------------

    configuration = safe_strip(
        get_cell(part_row, CONFIGURATION_COLUMN)
    )

    if not configuration or configuration == "-":
        return "1"

    channel_match = re.search(
        r"(\d+)\s*-?\s*(?:CHANNEL|CH\b|LINE)",
        configuration,
        flags=re.IGNORECASE,
    )

    if channel_match:
        return channel_match.group(1)

    text = configuration.upper()

    if "SINGLE" in text or "DISCRETE" in text or "UNI" in text:
        return "1"

    if "DUAL" in text:
        return "2"

    if "ARRAY" in text:
        return "-"

    return "1"


# ============================================================
# PARSE TVS / ESD ROW
#
# One parser for both tables - they describe the same kind of
# device, and every column below is looked up by name, so each
# sheet simply fills in the ones it carries.
# ============================================================

def parse_goodark_tvs_row(
    part_row,
    source_name,
    part_number,
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"Good-Ark {source_name}",
        "Grade": get_grade(part_row),
        "Direction": get_direction(part_row),
        "Channels": get_channels(part_row),
        "Package": safe_strip(get_cell(part_row, PACKAGE_COLUMN)),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Breakdown (Min)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-",
        "IEC 61000-4-5": "-",
        "Capacitance": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    # The ESD table calls the standoff VRWM, the TVS table
    # calls it VWM.

    standoff = get_number(
        part_row,
        STANDOFF_COLUMN,
    )

    if standoff is None:
        standoff = get_number(
            part_row,
            WORKING_VOLTAGE_COLUMN,
        )

    if standoff is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{standoff:g} V"

    breakdown_min = get_number(
        part_row,
        BREAKDOWN_MIN_COLUMN,
    )

    if breakdown_min is not None:
        specs_result["Voltage - Breakdown (Min)"] = f"{breakdown_min:g} V"

    # VC Max is the clamping voltage at IPP; VC Typ stands in
    # when the maximum is not stated.

    clamping = get_number(
        part_row,
        CLAMPING_MAX_COLUMN,
    )

    if clamping is None:
        clamping = get_number(
            part_row,
            CLAMPING_TYP_COLUMN,
        )

    if clamping is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamping:g} V"

    # Contact discharge is the 61000-4-2 figure; air discharge
    # is the fallback.

    esd_contact = get_number(
        part_row,
        ESD_CONTACT_COLUMN,
    )

    if esd_contact is None:
        esd_contact = get_number(
            part_row,
            ESD_AIR_COLUMN,
        )

    if esd_contact is not None:
        specs_result["IEC 61000-4-2"] = f"{esd_contact:g} kV"

    surge_current = get_number(
        part_row,
        PEAK_PULSE_CURRENT_COLUMN,
    )

    if surge_current is not None:
        specs_result["IEC 61000-4-5"] = f"{surge_current:g} A"

    # PPP on the ESD sheet, Pppm on the TVS sheet.

    power = get_number(
        part_row,
        ESD_PEAK_PULSE_POWER_COLUMN,
    )

    if power is None:
        power = get_number(
            part_row,
            PEAK_PULSE_POWER_COLUMN,
        )

    if power is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power:g} W"

    capacitance = get_number(
        part_row,
        CAPACITANCE_COLUMN,
    )

    if capacitance is not None:
        specs_result["Capacitance"] = f"{capacitance:.2f} pF"

    return specs_result


# ============================================================
# PARSE ZENER ROW
# ============================================================

def parse_goodark_zener_row(
    part_row,
    source_name,
    part_number,
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"Good-Ark {source_name}",
        "Grade": get_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Package": safe_strip(get_cell(part_row, PACKAGE_COLUMN)),
        "Price ($/ku)": "-",
    }

    zener_nom = get_number(
        part_row,
        ZENER_NOM_COLUMN,
    )

    zener_min = get_number(
        part_row,
        ZENER_MIN_COLUMN,
    )

    zener_max = get_number(
        part_row,
        ZENER_MAX_COLUMN,
    )

    zener_value = None

    if zener_nom is not None:
        zener_value = zener_nom

    elif zener_min is not None and zener_max is not None:
        zener_value = (zener_min + zener_max) / 2

    elif zener_min is not None:
        zener_value = zener_min

    elif zener_max is not None:
        zener_value = zener_max

    if zener_value is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{zener_value:g} V"

    # Good-Ark states the tolerance outright. Where it does
    # not, VZ Min against the nominal gives the same figure:
    #     (VZ Nom - VZ Min) / VZ Nom

    tolerance = get_number(
        part_row,
        TOLERANCE_COLUMN,
    )

    if tolerance is not None:
        specs_result["Tolerance"] = f"\u00b1{tolerance:g}%"

    elif zener_value is not None and zener_min is not None and zener_value > 0:
        specs_result["Tolerance"] = f"\u00b1{(zener_value - zener_min) / zener_value * 100:.3g}%"

    power = get_number(
        part_row,
        POWER_DISSIPATION_COLUMN,
    )

    if power is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power:g} W"

    return specs_result


# ============================================================
# FETCH SPECS FROM EXCEL
#
# goodark_dfs is {"ESD": df, "TVS": df, "Zener": df}.
# ============================================================

def fetch_goodark_specs_from_excel(
    part_number,
    goodark_dfs,
):

    if not part_number or not goodark_dfs:
        return None

    target = str(part_number).strip().lower()

    target_loose = re.sub(
        r"[\W_]+",
        "",
        target,
    )

    search_priority = [
        "ESD",
        "TVS",
        "Zener",
    ]

    ordered_keys = [
        key
        for key in search_priority
        if key in goodark_dfs
    ] + [
        key
        for key in goodark_dfs
        if key not in search_priority
    ]

    for df_name in ordered_keys:

        df = goodark_dfs.get(df_name)

        if df is None or df.empty:
            continue

        df.columns = df.columns.str.strip()

        part_column = next(
            (
                column
                for column in df.columns
                if str(column).strip().lower() == PART_COLUMN.lower()
            ),
            None,
        )

        if not part_column:
            continue

        column_values = df[part_column].astype(str).str.strip()

        part_rows = df[column_values.str.lower() == target]

        if part_rows.empty:

            loose_values = column_values.str.lower().str.replace(
                r"[\W_]+",
                "",
                regex=True,
            )

            part_rows = df[loose_values == target_loose]

        if part_rows.empty:
            continue

        part_row = part_rows.iloc[0]

        exact_name = safe_strip(
            part_row[part_column]
        )

        print(
            f"Found specs for Part '{part_number}' "
            f"in Good-Ark database '{df_name}'."
        )

        if "zener" in df_name.lower():

            return parse_goodark_zener_row(
                part_row,
                df_name,
                exact_name,
            )

        return parse_goodark_tvs_row(
            part_row,
            df_name,
            exact_name,
        )

    print(
        f"Part '{part_number}' not found "
        f"in any Good-Ark database."
    )

    return None


if __name__ == "__main__":
    try:
        main()

    except Exception:
        print()
        print("=" * 70)
        print("SCRAPER ERROR")
        print("=" * 70)

        traceback.print_exc()

    finally:
        print()
        print("Window will stay open so you can read the output.")
        input("Press ENTER to close...")