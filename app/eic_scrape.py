import re
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font

from parsing import safe_strip


ZENER_URL = "https://www.eicsemi.com/productZener.aspx"
TVS_URL = "https://www.eicsemi.com/ProductLCTVS.aspx"

ZENER_OUTPUT = "eic_zener_specs.xlsx"
TVS_OUTPUT = "eic_tvs_specs.xlsx"

REQUEST_DELAY = 0.20
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


ZENER_COLUMNS = [
    "Part Number",
    "Data Sheet",
    "VZ Min (V)",
    "VZ Nom (V)",
    "VZ Max (V)",
    "IZT (mA)",
    "ZZT @ IZT (ohm)",
    "ZZK @ IZK (ohm)",
    "IZK (mA)",
    "IR (uA)",
    "VR (V)",
    "IZM (mA)",
    "PD (mW)",
    "Package",
    "Series",
]


TVS_COLUMNS = [
    "Part Number",
    "Data Sheet",
    "PPK (W)",
    "VBR Min (V)",
    "VBR Max (V)",
    "IT (mA)",
    "VRWM (V)",
    "IR @ VRWM (mA)",
    "IRSM (A)",
    "VRSM (V)",
    "Junction Capacitance @ 0V (pF)",
    "Package",
    "Series",
]


def clean_text(value):
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def cell_values(cell):
    values = []

    for s in cell.stripped_strings:
        s = clean_text(s)

        if not s:
            continue

        if s.lower() in {
            "pdf",
            "rfq",
            "added",
        }:
            continue

        values.append(s)

    return values


def first_value(cell):
    values = cell_values(cell)

    if values:
        return values[0]

    return ""


def pad(values, size):
    values = list(values[:size])

    while len(values) < size:
        values.append("")

    return values


def looks_like_part_number(text):
    text = clean_text(text)

    if not text:
        return False

    lowered = text.lower()

    if lowered in {
        "part number",
        "no record found",
        "product category",
    }:
        return False

    return (
        bool(re.search(r"[A-Za-z]", text))
        and bool(re.search(r"\d", text))
        and len(text) <= 100
    )


def extract_pdf_url(td, page_url):
    if td is None:
        return ""

    for link in td.find_all("a", href=True):
        href = clean_text(link.get("href"))

        if ".pdf" in href.lower():
            return urljoin(page_url, href)

    html = str(td)

    match = re.search(
        r"(?:href\s*=\s*[\"']?)?([^\"'<>]+?\.pdf)",
        html,
        flags=re.IGNORECASE,
    )

    if match:
        return urljoin(
            page_url,
            match.group(1).strip(),
        )

    return ""


# ============================================================
# ZENER PARSER
# ============================================================

def parse_zener_row(tr, page_url):
    """
    Expected EIC Zener row layout:

      0 Part Number
      1 Data Sheet
      2 RFQ
      3 VZ Min / Nom / Max / IZT
      4 ZZT / ZZK / IZK
      5 IR / VR
      6 IZM
      7 PD
      8 Package
      9 Series
    """

    tds = tr.find_all(
        "td",
        recursive=False,
    )

    if len(tds) < 10:
        tds = tr.find_all("td")

    if len(tds) < 10:
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

    vz = pad(
        cell_values(tds[3]),
        4,
    )

    impedance = pad(
        cell_values(tds[4]),
        3,
    )

    leakage = pad(
        cell_values(tds[5]),
        2,
    )

    return {
        "Part Number": part_number,
        "Data Sheet": extract_pdf_url(
            tds[1],
            page_url,
        ),

        "VZ Min (V)": vz[0],
        "VZ Nom (V)": vz[1],
        "VZ Max (V)": vz[2],
        "IZT (mA)": vz[3],

        "ZZT @ IZT (ohm)": impedance[0],
        "ZZK @ IZK (ohm)": impedance[1],
        "IZK (mA)": impedance[2],

        "IR (uA)": leakage[0],
        "VR (V)": leakage[1],

        "IZM (mA)": first_value(
            tds[6]
        ),

        "PD (mW)": first_value(
            tds[7]
        ),

        "Package": first_value(
            tds[8]
        ),

        "Series": " ".join(
            cell_values(tds[9])
        ),
    }


# ============================================================
# TVS PARSER
# ============================================================

def parse_tvs_row(tr, page_url):
    """
    Expected EIC Low Cap TVS layout:

      0 Part Number
      1 Data Sheet
      2 RFQ
      3 PPK
      4 VBR Min / VBR Max / IT
      5 VRWM
      6 IR @ VRWM
      7 IRSM
      8 VRSM
      9 Junction Capacitance
     10 Package
     11 Series
    """

    tds = tr.find_all(
        "td",
        recursive=False,
    )

    if len(tds) < 12:
        tds = tr.find_all("td")

    if len(tds) < 12:
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

    vbr = pad(
        cell_values(tds[4]),
        3,
    )

    return {
        "Part Number": part_number,

        "Data Sheet": extract_pdf_url(
            tds[1],
            page_url,
        ),

        "PPK (W)": first_value(
            tds[3]
        ),

        "VBR Min (V)": vbr[0],
        "VBR Max (V)": vbr[1],
        "IT (mA)": vbr[2],

        "VRWM (V)": first_value(
            tds[5]
        ),

        "IR @ VRWM (mA)": first_value(
            tds[6]
        ),

        "IRSM (A)": first_value(
            tds[7]
        ),

        "VRSM (V)": first_value(
            tds[8]
        ),

        "Junction Capacitance @ 0V (pF)": first_value(
            tds[9]
        ),

        "Package": first_value(
            tds[10]
        ),

        "Series": " ".join(
            cell_values(tds[11])
        ),
    }


# ============================================================
# PAGE HANDLING
# ============================================================

def get_page(
    session,
    base_url,
    page_number,
):

    if page_number == 1:
        url = base_url

    else:
        url = (
            f"{base_url}"
            f"?vPage={page_number}"
        )

    headers = dict(HEADERS)

    headers["Referer"] = base_url

    response = session.get(
        url,
        headers=headers,
        timeout=TIMEOUT,
    )

    response.raise_for_status()

    if not response.encoding:
        response.encoding = "utf-8"

    return (
        url,
        response.text,
    )


def detect_last_page(html):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    page_numbers = [1]

    for a in soup.find_all(
        "a",
        href=True,
    ):

        href = a.get(
            "href",
            "",
        )

        match = re.search(
            r"[?&]vPage=(\d+)",
            href,
            flags=re.IGNORECASE,
        )

        if match:

            page_numbers.append(
                int(
                    match.group(1)
                )
            )

    return max(page_numbers)


def scrape_page(
    html,
    page_url,
    parser,
):

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    records = []

    for tr in soup.find_all("tr"):

        record = parser(
            tr,
            page_url,
        )

        if record:
            records.append(record)

    return records


# ============================================================
# EXCEL SAVING
# ============================================================

def save_excel(
    records,
    output_file,
    columns,
    sheet_name,
):

    df = pd.DataFrame(
        records
    )

    for col in columns:

        if col not in df.columns:
            df[col] = ""

    df = df[columns]

    df = (
        df
        .drop_duplicates(
            subset=["Part Number"],
            keep="first",
        )
        .reset_index(
            drop=True
        )
        .fillna("")
    )

    df.to_excel(
        output_file,
        index=False,
    )

    wb = load_workbook(
        output_file
    )

    ws = wb.active

    ws.title = sheet_name

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = (
        ws.dimensions
    )

    for cell in ws[1]:

        cell.font = Font(
            bold=True
        )

    # Auto-size Excel columns
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
                else str(cell.value)
            )

            max_length = max(
                max_length,
                len(value),
            )

        if letter == "B":

            width = min(
                max(
                    max_length + 2,
                    25,
                ),
                70,
            )

        else:

            width = min(
                max(
                    max_length + 2,
                    12,
                ),
                35,
            )

        ws.column_dimensions[
            letter
        ].width = width

    # Make datasheet URLs clickable
    if "Data Sheet" in columns:

        datasheet_col = (
            columns.index(
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
                column=datasheet_col,
            )

            if (
                isinstance(
                    cell.value,
                    str,
                )
                and
                cell.value.startswith(
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
        output_file
    )

    return df


# ============================================================
# SCRAPE ONE CATALOG
# ============================================================

def scrape_catalog(
    session,
    catalog_name,
    base_url,
    output_file,
    columns,
    parser,
):

    print()
    print("=" * 70)
    print(
        f"SCRAPING: "
        f"{catalog_name}"
    )
    print("=" * 70)

    print(
        f"URL: "
        f"{base_url}"
    )

    print(
        f"Output: "
        f"{output_file}"
    )

    print()

    print(
        "Downloading page 1..."
    )

    (
        page1_url,
        page1_html,
    ) = get_page(
        session,
        base_url,
        1,
    )

    last_page = detect_last_page(
        page1_html
    )

    print(
        f"Detected "
        f"{last_page} "
        f"page(s)."
    )

    page1_records = scrape_page(
        page1_html,
        page1_url,
        parser,
    )

    print(
        f"Page 1/{last_page}: "
        f"found "
        f"{len(page1_records)} "
        f"product rows."
    )

    if not page1_records:

        raise RuntimeError(
            f"No {catalog_name} "
            f"product rows could "
            f"be parsed."
        )

    print()
    print(
        "First parsed row:"
    )

    for key in columns:

        print(
            f"  {key}: "
            f"{page1_records[0].get(key, '')}"
        )

    all_records = list(
        page1_records
    )

    seen = {
        record["Part Number"]
        for record
        in page1_records
    }

    for page_number in range(
        2,
        last_page + 1,
    ):

        (
            page_url,
            html,
        ) = get_page(
            session,
            base_url,
            page_number,
        )

        records = scrape_page(
            html,
            page_url,
            parser,
        )

        new_records = []

        for record in records:

            part_number = (
                record[
                    "Part Number"
                ]
            )

            if (
                part_number
                not in seen
            ):

                seen.add(
                    part_number
                )

                new_records.append(
                    record
                )

        all_records.extend(
            new_records
        )

        print(
            f"Page "
            f"{page_number}/"
            f"{last_page}: "
            f"{len(records)} rows, "
            f"{len(new_records)} new, "
            f"{len(all_records)} total"
        )

        if not records:

            print(
                f"WARNING: "
                f"no rows parsed "
                f"from {page_url}"
            )

        time.sleep(
            REQUEST_DELAY
        )

    print()

    print(
        f"Saving "
        f"{len(all_records)} "
        f"unique products to "
        f"{output_file}..."
    )

    df = save_excel(
        all_records,
        output_file,
        columns,
        catalog_name,
    )

    print(
        f"Finished "
        f"{catalog_name}: "
        f"{len(df)} rows saved."
    )

    return df


# ============================================================
# SPEC COLUMNS
#
# The saved sheets' header names, kept as constants so the
# parsers below read the columns by name instead of by
# position. They match ZENER_COLUMNS / TVS_COLUMNS exactly.
# ============================================================

PART_COLUMN = "Part Number"

DATASHEET_COLUMN = "Data Sheet"

PACKAGE_COLUMN = "Package"

SERIES_COLUMN = "Series"

# --- TVS ---

PEAK_PULSE_POWER_COLUMN = "PPK (W)"

BREAKDOWN_MIN_COLUMN = "VBR Min (V)"

BREAKDOWN_MAX_COLUMN = "VBR Max (V)"

BREAKDOWN_TEST_CURRENT_COLUMN = "IT (mA)"

STANDOFF_COLUMN = "VRWM (V)"

STANDOFF_LEAKAGE_COLUMN = "IR @ VRWM (mA)"

SURGE_CURRENT_COLUMN = "IRSM (A)"

CLAMPING_COLUMN = "VRSM (V)"

CAPACITANCE_COLUMN = "Junction Capacitance @ 0V (pF)"

# --- ZENER ---

ZENER_MIN_COLUMN = "VZ Min (V)"

ZENER_NOM_COLUMN = "VZ Nom (V)"

ZENER_MAX_COLUMN = "VZ Max (V)"

ZENER_TEST_CURRENT_COLUMN = "IZT (mA)"

ZENER_IMPEDANCE_IZT_COLUMN = "ZZT @ IZT (ohm)"

ZENER_IMPEDANCE_IZK_COLUMN = "ZZK @ IZK (ohm)"

ZENER_KNEE_CURRENT_COLUMN = "IZK (mA)"

ZENER_LEAKAGE_CURRENT_COLUMN = "IR (uA)"

ZENER_LEAKAGE_VOLTAGE_COLUMN = "VR (V)"

ZENER_MAX_CURRENT_COLUMN = "IZM (mA)"

POWER_DISSIPATION_COLUMN = "PD (mW)"


# ============================================================
# SPEC HELPERS
# ============================================================

def get_cell(
    part_row,
    column_name,
):

    # --------------------------------------------------------
    # Exact column name first, then a stripped/case match in
    # case a spreadsheet round trip altered the header.
    # --------------------------------------------------------

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
    # EIC states no grade column. The datasheet path carries
    # it instead: "Normal grade/..." for the commercial
    # catalog, an AEC/automotive folder for the qualified one.
    # --------------------------------------------------------

    datasheet = safe_strip(
        get_cell(part_row, DATASHEET_COLUMN)
    ).lower()

    if "aec" in datasheet or "automotive" in datasheet:
        return "Automotive"

    return "Non-Automotive"


def get_direction(
    part_number
):

    # --------------------------------------------------------
    # No direction column either. EIC follows the usual TVS
    # naming, where a trailing C (or CA) marks the
    # bidirectional version of a part and everything else is
    # unidirectional.
    # --------------------------------------------------------

    name = str(part_number).strip().upper()

    if re.search(r"C[AB]?$", name):
        return "Bidirectional"

    return "Unidirectional"


# ============================================================
# PARSE TVS ROW
# ============================================================

def parse_eic_tvs_row(
    part_row,
    source_name,
    part_number,
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"EIC {source_name}",
        "Grade": get_grade(part_row),
        "Direction": get_direction(part_number),
        "Channels": "1",
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

    standoff = get_number(
        part_row,
        STANDOFF_COLUMN,
    )

    if standoff is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{standoff:g} V"

    breakdown_min = get_number(
        part_row,
        BREAKDOWN_MIN_COLUMN,
    )

    if breakdown_min is not None:
        specs_result["Voltage - Breakdown (Min)"] = f"{breakdown_min:g} V"

    # VRSM is the maximum clamping voltage, rated at IRSM.

    clamping = get_number(
        part_row,
        CLAMPING_COLUMN,
    )

    if clamping is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamping:g} V"

    # IRSM, the peak pulse current that clamping voltage is
    # rated at, is the surge rating.

    surge_current = get_number(
        part_row,
        SURGE_CURRENT_COLUMN,
    )

    if surge_current is not None:
        specs_result["IEC 61000-4-5"] = f"{surge_current:g} A"

    peak_pulse_power = get_number(
        part_row,
        PEAK_PULSE_POWER_COLUMN,
    )

    if peak_pulse_power is not None:
        specs_result["Power Dissipation (Pd)"] = f"{peak_pulse_power:g} W"

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

def parse_eic_zener_row(
    part_row,
    source_name,
    part_number,
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"EIC {source_name}",
        "Grade": get_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Package": safe_strip(get_cell(part_row, PACKAGE_COLUMN)),
        "Price ($/ku)": "-",
    }

    # --------------------------------------------------------
    # Vz comes from VZ Nom, falling back to the midpoint of
    # the VZ Min / VZ Max window. EIC leaves the unused two
    # of the three columns as "-".
    # --------------------------------------------------------

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

    # Tolerance is how far VZ Min sits below the nominal:
    #     (VZ Nom - VZ Min) / VZ Nom
    # where the nominal is the midpoint found above whenever
    # EIC states the window instead of a VZ Nom.

    if zener_value is not None and zener_min is not None and zener_value > 0:
        specs_result["Tolerance"] = f"±{(zener_value - zener_min) / zener_value * 100:.3g}%"

    # PD is stated in mW, everything downstream is in W.

    power_dissipation = get_number(
        part_row,
        POWER_DISSIPATION_COLUMN,
    )

    if power_dissipation is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power_dissipation / 1000:g} W"

    return specs_result


# ============================================================
# FETCH SPECS FROM EXCEL
#
# eic_dfs is {"TVS": df, "Zener": df}.
# ============================================================

def fetch_eic_specs_from_excel(
    part_number,
    eic_dfs,
):

    if not part_number or not eic_dfs:
        return None

    target = str(part_number).strip().lower()

    target_loose = re.sub(
        r"[\W_]+",
        "",
        target,
    )

    search_priority = [
        "TVS",
        "Zener",
    ]

    ordered_keys = [
        key
        for key in search_priority
        if key in eic_dfs
    ] + [
        key
        for key in eic_dfs
        if key not in search_priority
    ]

    for df_name in ordered_keys:

        df = eic_dfs.get(df_name)

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
            f"in EIC database '{df_name}'."
        )

        if "zener" in df_name.lower():

            return parse_eic_zener_row(
                part_row,
                df_name,
                exact_name,
            )

        return parse_eic_tvs_row(
            part_row,
            df_name,
            exact_name,
        )

    print(
        f"Part '{part_number}' not found "
        f"in any EIC database."
    )

    return None


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "EIC Semiconductor scraper"
    )

    print("=" * 70)

    print(
        "No Selenium/browser is used."
    )

    print(
        "Pages are loaded directly "
        "using ?vPage=N."
    )

    session = (
        requests.Session()
    )

    # --------------------------------------------------------
    # ZENER
    # --------------------------------------------------------

    zener_df = scrape_catalog(

        session=session,

        catalog_name=(
            "EIC Zener"
        ),

        base_url=(
            ZENER_URL
        ),

        output_file=(
            ZENER_OUTPUT
        ),

        columns=(
            ZENER_COLUMNS
        ),

        parser=(
            parse_zener_row
        ),
    )

    # --------------------------------------------------------
    # TVS
    # --------------------------------------------------------

    tvs_df = scrape_catalog(

        session=session,

        catalog_name=(
            "EIC Low Cap TVS"
        ),

        base_url=(
            TVS_URL
        ),

        output_file=(
            TVS_OUTPUT
        ),

        columns=(
            TVS_COLUMNS
        ),

        parser=(
            parse_tvs_row
        ),
    )

    print()
    print("=" * 70)
    print(
        "ALL DONE"
    )
    print("=" * 70)

    print(
        f"{ZENER_OUTPUT}: "
        f"{len(zener_df)} rows"
    )

    print(
        f"{TVS_OUTPUT}: "
        f"{len(tvs_df)} rows"
    )


if __name__ == "__main__":

    import traceback

    try:

        main()

    except Exception:

        print()
        print("=" * 70)
        print(
            "SCRAPER ERROR"
        )
        print("=" * 70)

        traceback.print_exc()

    finally:

        print()

        print(
            "Window will stay open "
            "so you can read the output."
        )

        input(
            "Press ENTER to close..."
        )