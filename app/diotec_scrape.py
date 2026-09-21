from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

import pandas as pd
from pathlib import Path
import time
import re
from parsing import safe_strip


# ==========================================================
# SETTINGS
# ==========================================================

BASE_FOLDER = Path(__file__).resolve().parent

ESD_URL = "https://diotec.com/en/productlist/ESD.html"
TVS_URL = "https://diotec.com/en/productlist/TVS.html"
ZENER_URL = "https://diotec.com/en/productlist/Z.html"

ESD_OUTPUT = BASE_FOLDER / "diotec_esd_specs.xlsx"
TVS_OUTPUT = BASE_FOLDER / "diotec_tvs_specs.xlsx"
ZENER_OUTPUT = BASE_FOLDER / "diotec_zener_specs.xlsx"

# The scrape run builds these in MAIN. They stay None when applesauce imports
# this module for fetch_diotec_specs_from_excel, so importing it never opens a
# browser or re-runs the scrape.

driver = None
wait = None


# ==========================================================
# COOKIE POPUP
# ==========================================================

def close_cookie_popup():

    try:
        buttons = driver.find_elements(
            By.XPATH,
            "//*[self::button or self::a]"
            "[contains(normalize-space(.),'Accept all')]"
        )

        for button in buttons:

            if button.is_displayed():

                print("Closing cookie popup...")

                driver.execute_script(
                    "arguments[0].click();",
                    button
                )

                time.sleep(0.8)
                return

    except Exception:
        pass


# ==========================================================
# FIND SHOW ENTRIES DROPDOWN
# ==========================================================

def find_length_dropdown():

    selects = driver.find_elements(
        By.TAG_NAME,
        "select"
    )

    for element in selects:

        try:

            if not element.is_displayed():
                continue

            dropdown = Select(element)

            options = [
                option.text.strip()
                for option in dropdown.options
            ]

            if (
                "10" in options
                and "25" in options
                and "50" in options
                and "All" in options
            ):
                return element

        except Exception:
            continue

    return None


# ==========================================================
# SELECT ALL
# ==========================================================

def select_all_rows():

    print("Looking for Show entries dropdown...")

    element = wait.until(
        lambda d: find_length_dropdown()
    )

    dropdown = Select(element)

    print("Dropdown options:")

    for option in dropdown.options:

        print(
            " ",
            repr(option.text),
            "value=",
            repr(option.get_attribute("value"))
        )

    print("Selecting All...")

    dropdown.select_by_visible_text("All")


# ==========================================================
# WAIT UNTIL ALL RESULTS ARE DISPLAYED
# ==========================================================

def wait_for_all_rows():

    print("Waiting for all entries to display...")

    def all_is_displayed(d):

        try:

            text = d.find_element(
                By.TAG_NAME,
                "body"
            ).text

            matches = re.findall(
                r"Showing\s+1\s+to\s+(\d+)\s+of\s+(\d+)\s+entries",
                text
            )

            for shown, total in matches:

                if int(shown) == int(total):
                    return True

        except Exception:
            pass

        return False

    wait.until(all_is_displayed)

    time.sleep(2)

    print("All entries loaded.")


# ==========================================================
# EXTRACT PRODUCT DATATABLE
# ==========================================================

def extract_datatable():

    print("Reading data directly from DataTables...")

    result = driver.execute_script("""
        let results = [];

        if (
            typeof jQuery === 'undefined' ||
            !jQuery.fn ||
            !jQuery.fn.dataTable
        ) {
            return {
                error: "jQuery DataTables was not found"
            };
        }

        let tables = document.querySelectorAll("table");

        for (let table of tables) {

            try {

                if (
                    !jQuery.fn.dataTable.isDataTable(table)
                ) {
                    continue;
                }

                let dt = jQuery(table).DataTable();

                let data = dt.rows({
                    search: "applied"
                }).data().toArray();

                let settings = dt.settings()[0];

                results.push({
                    id: table.id || "",
                    rowCount: data.length,
                    columnCount: settings.aoColumns.length,
                    data: data
                });

            } catch (e) {
                // Ignore unrelated DataTables
            }
        }

        return {
            tables: results
        };
    """)

    if not result:

        raise Exception(
            "DataTables returned no result."
        )

    if "error" in result:

        raise Exception(
            result["error"]
        )

    tables = result.get(
        "tables",
        []
    )

    print(
        f"DataTables instances found: {len(tables)}"
    )

    for i, table in enumerate(tables):

        print(
            f"  Table {i + 1}: "
            f"{table.get('rowCount', 0)} rows, "
            f"{table.get('columnCount', 0)} columns"
        )

    usable_tables = [
        table
        for table in tables
        if table.get("rowCount", 0) > 0
    ]

    if not usable_tables:

        raise Exception(
            "Could not identify Diotec product DataTable."
        )

    # Main product table is the DataTable with the most rows
    product_table = max(
        usable_tables,
        key=lambda x:
        x.get(
            "rowCount",
            0
        )
    )

    return product_table.get(
        "data",
        []
    )


# ==========================================================
# CLEAN INDIVIDUAL CELL
# ==========================================================

def clean_cell(value):

    if value is None:
        return ""

    # ------------------------------------------------------
    # LIST
    # ------------------------------------------------------

    if isinstance(value, list):

        items = [
            str(clean_cell(item))
            for item in value
            if item is not None
        ]

        return " | ".join(
            item
            for item in items
            if item.strip()
        )

    # ------------------------------------------------------
    # DICTIONARY
    # ------------------------------------------------------

    if isinstance(value, dict):

        items = []

        for key, item in value.items():

            cleaned = clean_cell(item)

            if str(cleaned).strip():

                items.append(
                    f"{key}: {cleaned}"
                )

        return " | ".join(items)

    # ------------------------------------------------------
    # NUMBER / BOOL
    # ------------------------------------------------------

    if not isinstance(value, str):

        return str(value)

    # ------------------------------------------------------
    # HTML -> VISIBLE TEXT
    # ------------------------------------------------------

    try:

        text = driver.execute_script("""
            const div =
                document.createElement("div");

            div.innerHTML =
                arguments[0];

            return (
                div.innerText ||
                div.textContent ||
                ""
            ).trim();
        """, value)

    except Exception:

        text = value

    text = str(text)

    text = (
        text
        .replace("»", "")
        .replace("\\n", " ")
        .strip()
    )

    return " ".join(
        text.split()
    )


# ==========================================================
# CLEAN ARTICLE NUMBER
# ==========================================================

def clean_article_number(value):

    value = str(value).strip()

    value = value.replace(
        "»",
        ""
    )

    # Example:
    #
    # DI5315-02FDatasheetProduct info & Stock
    #
    # becomes:
    #
    # DI5315-02F

    value = re.sub(
        r"\s*Datasheet\s*Product\s*info\s*&\s*Stock.*$",
        "",
        value,
        flags=re.IGNORECASE
    )

    # Handles:
    # DatasheetProduct info & Stock

    value = re.sub(
        r"DatasheetProduct\s*info\s*&\s*Stock.*$",
        "",
        value,
        flags=re.IGNORECASE
    )

    # Final fallback:
    # remove anything beginning with "Datasheet"

    value = re.sub(
        r"Datasheet.*$",
        "",
        value,
        flags=re.IGNORECASE
    )

    return value.strip()


# ==========================================================
# GENERIC DIOTEC PAGE SCRAPER
# ==========================================================

def scrape_diotec_page(
    url,
    page_name
):

    print()
    print(
        "########################################"
    )

    print(
        f"SCRAPING DIOTEC {page_name}"
    )

    print(
        "########################################"
    )

    driver.get(url)

    # ------------------------------------------------------
    # Wait for DataTables page
    # ------------------------------------------------------

    wait.until(
        lambda d:
        "Showing"
        in d.find_element(
            By.TAG_NAME,
            "body"
        ).text
    )

    time.sleep(2)

    # ------------------------------------------------------
    # Close cookies
    # ------------------------------------------------------

    close_cookie_popup()

    # ------------------------------------------------------
    # Show All
    # ------------------------------------------------------

    select_all_rows()

    # ------------------------------------------------------
    # Wait until All is loaded
    # ------------------------------------------------------

    wait_for_all_rows()

    # ------------------------------------------------------
    # Get DataTable
    # ------------------------------------------------------

    raw_data = extract_datatable()

    print(
        f"Raw DataTables rows returned: "
        f"{len(raw_data)}"
    )

    cleaned_data = []

    # ------------------------------------------------------
    # Clean all cells
    # ------------------------------------------------------

    for row in raw_data:

        if isinstance(row, dict):

            row_values = list(
                row.values()
            )

        elif isinstance(row, list):

            row_values = row

        else:

            row_values = [
                row
            ]

        cleaned_row = [
            clean_cell(value)
            for value in row_values
        ]

        cleaned_data.append(
            cleaned_row
        )

    print(
        f"{page_name} rows extracted: "
        f"{len(cleaned_data)}"
    )

    return cleaned_data


# ==========================================================
# ESD RAW COLUMN MAPPING
# ==========================================================

ESD_ALL_HEADERS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "VBRmin [V] Breakdown voltage",
    "VBRnom [V] Breakdown voltage",
    "VBRmax [V] Breakdown voltage",
    "@ IT [mA] Breakdown voltage",

    "VWM [V] Stand-off voltage",
    "@ ID [A] Stand-off voltage",
    "@ IR [mA] Stand-off voltage",

    "PPPM [W] Peak pulse power dissipation",

    "Channels No of channels",

    "Vpp [kV] ESD contact",
    "Vpp [kV] ESD HBM",

    "Tjmin [°C] Junction temperature",
    "Tjmax [°C] Junction temperature",

    "ID @ VWM [µA] Maximum reverse current",

    "VC [V] Maximum clamping voltage",

    "@ IPPM [A] Maximum clamping voltage",

    "Cj [pF] Junction capacitance",

    "@ VR [V] Junction capacitance"
]


# ==========================================================
# ESD FINAL COLUMNS
# ==========================================================

ESD_KEEP_COLUMNS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "VBRmin [V] Breakdown voltage",

    "VBRmax [V] Breakdown voltage",

    "VWM [V] Stand-off voltage",

    "PPPM [W] Peak pulse power dissipation",

    "Channels No of channels",

    "Vpp [kV] ESD contact",

    "Vpp [kV] ESD HBM",

    "ID @ VWM [µA] Maximum reverse current",

    "VC [V] Maximum clamping voltage",

    "@ IPPM [A] Maximum clamping voltage",

    "Cj [pF] Junction capacitance"
]


# ==========================================================
# TVS RAW COLUMN MAPPING
# ==========================================================

TVS_HEADERS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "VWM [V] Stand-off voltage",
    "@ ID [A] Stand-off voltage",
    "@ IR [mA] Stand-off voltage",

    "ID @ VWM [µA] Maximum reverse current",

    "VBRmin [V] Breakdown voltage",
    "VBRnom [V] Breakdown voltage",
    "VBRmax [V] Breakdown voltage",
    "@ IT [mA] Breakdown voltage",

    "Voltage Tolerance",

    "VC [V] Maximum clamping voltage",
    "@ IPPM [A] Maximum clamping voltage",

    "PPPM [W] Peak pulse power dissipation",

    "Ptot [W] Power dissipation",
    "@ T Loc [°C] Power dissipation",
    "Location Power dissipation",

    "Tjmin [°C] Junction temperature",
    "Tjmax [°C] Junction temperature",

    "Cj [pF] Junction capacitance",
    "@ VR [V] Junction capacitance",

    "m net [g]",
    "Marking",

    "UL 94-V0",

    "RTH [K/W]",
    "RTH Location",

    "UL-Recognition",

    "RTH (b) [K/W]",
    "RTH Location (b)",

    "Tsmin [°C]",
    "Tsmax [°C]",

    "Solder & Assembly"
]


# ==========================================================
# TVS FINAL COLUMNS
#
# Same desired columns as ESD where they exist.
# ==========================================================

TVS_KEEP_COLUMNS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "VBRmin [V] Breakdown voltage",

    "VBRmax [V] Breakdown voltage",

    "VWM [V] Stand-off voltage",

    "PPPM [W] Peak pulse power dissipation",

    "ID @ VWM [µA] Maximum reverse current",

    "VC [V] Maximum clamping voltage",

    "@ IPPM [A] Maximum clamping voltage",

    "Cj [pF] Junction capacitance"
]


# ==========================================================
# ZENER RAW COLUMN MAPPING
# ==========================================================

ZENER_HEADERS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "Vzmin [V] Zener voltage",
    "Vznom [V] Zener voltage",
    "Vzmax [V] Zener voltage",
    "@ Iztest [mA] Zener voltage",

    "Ptot [W] Power dissipation",
    "@ T Loc [°C] Power dissipation",

    # This field exists in the raw table but is not needed
    "Location Power dissipation",

    "Voltage Tolerance",

    "Tjmin [°C] Junction temperature",
    "Tjmax [°C] Junction temperature",

    "IR [µA] Leakage current 25°C",
    "@ VR [V] Leakage current 25°C",

    "rzj [Ω] Dynamic resistance",
    "@ Iz [mA] Dynamic resistance",

    "Cj [pF] Junction capacitance",
    "@ VR [V] Junction capacitance",

    "rzj [Ω] Dynamic resistance (b)",
    "@ Iz [mA] Dynamic resistance (b)",

    "aVz [1E-4/°C] Temperature coefficient",

    "IZmax [mA] Zener current",
    "@ T Loc [°C] Zener current",
    "Location Zener current",

    "m net [g]",
    "Marking",

    "UL 94-V0",

    "RTH [K/W]",
    "RTH Location",

    "UL-Recognition",

    "RTH (b) [K/W]",
    "RTH Location (b)",

    "Tsmin [°C]",
    "Tsmax [°C]",

    "Solder & Assembly"
]


# ==========================================================
# ZENER FINAL COLUMNS
#
# ONLY these columns will be written to Excel.
# ==========================================================

ZENER_KEEP_COLUMNS = [

    "Article number",
    "Type",
    "Package",
    "Qualification",
    "Config",
    "Life Cycle",

    "Vzmin [V] Zener voltage",

    "Vznom [V] Zener voltage",

    "Vzmax [V] Zener voltage",

    "@ Iztest [mA] Zener voltage",

    "Ptot [W] Power dissipation",

    "@ T Loc [°C] Power dissipation",

    "Voltage Tolerance",

    "@ VR [V] Leakage current 25°C"
]


# ==========================================================
# BUILD DATAFRAME
# ==========================================================

def build_dataframe(
    data,
    headers
):

    expected_columns = len(
        headers
    )

    processed = []

    for row in data:

        # --------------------------------------------------
        # Only map the real product fields.
        #
        # Diotec's DataTables can return additional internal
        # metadata fields after the actual specifications.
        # --------------------------------------------------

        selected = row[
            :expected_columns
        ]

        # --------------------------------------------------
        # Pad unexpectedly short rows
        # --------------------------------------------------

        while (
            len(selected)
            < expected_columns
        ):

            selected.append("")

        processed.append(
            selected
        )

    # ------------------------------------------------------
    # Build DataFrame
    # ------------------------------------------------------

    df = pd.DataFrame(
        processed,
        columns=headers
    )

    # ------------------------------------------------------
    # Clean article numbers
    # ------------------------------------------------------

    df["Article number"] = (
        df["Article number"]
        .apply(
            clean_article_number
        )
    )

    # ------------------------------------------------------
    # Remove blank rows
    # ------------------------------------------------------

    df = df[
        df["Article number"] != ""
    ].copy()

    # ------------------------------------------------------
    # Remove duplicates
    # ------------------------------------------------------

    df = (
        df
        .drop_duplicates()
        .reset_index(
            drop=True
        )
    )

    return df


# ==========================================================
# KEEP REQUESTED COLUMNS
# ==========================================================

def keep_columns(
    df,
    requested_columns
):

    # ------------------------------------------------------
    # Only keep requested columns that actually exist.
    #
    # This means if a spec exists for ESD but not TVS,
    # TVS simply skips it instead of creating a blank field.
    # ------------------------------------------------------

    available_columns = [
        column
        for column in requested_columns
        if column in df.columns
    ]

    return df[
        available_columns
    ].copy()


# ==========================================================
# SAVE EXCEL
# ==========================================================

def save_excel(
    df,
    output_file
):

    try:

        df.to_excel(
            output_file,
            index=False
        )

    except PermissionError:

        print()
        print(
            f"ERROR: {output_file.name} "
            f"is currently open in Excel."
        )

        print(
            "Close the workbook and "
            "run the scraper again."
        )

        raise

    print()
    print(
        f"Saved {len(df)} rows "
        f"and {len(df.columns)} columns."
    )

    print(
        f"File: {output_file}"
    )


# ==========================================================
# SPEC COLUMNS
#
# The columns the parsers below read, named exactly as they
# are written to the spreadsheets by KEEP_COLUMNS above.
# ==========================================================

PART_COLUMN = "Article number"

PACKAGE_COLUMN = "Package"

QUALIFICATION_COLUMN = "Qualification"

CONFIG_COLUMN = "Config"

STANDOFF_COLUMN = "VWM [V] Stand-off voltage"

CLAMPING_COLUMN = "VC [V] Maximum clamping voltage"

SURGE_CURRENT_COLUMN = "@ IPPM [A] Maximum clamping voltage"

PEAK_PULSE_POWER_COLUMN = "PPPM [W] Peak pulse power dissipation"

CHANNELS_COLUMN = "Channels No of channels"

ESD_CONTACT_COLUMN = "Vpp [kV] ESD contact"

CAPACITANCE_COLUMN = "Cj [pF] Junction capacitance"

ZENER_MIN_COLUMN = "Vzmin [V] Zener voltage"

ZENER_NOM_COLUMN = "Vznom [V] Zener voltage"

ZENER_MAX_COLUMN = "Vzmax [V] Zener voltage"

TOLERANCE_COLUMN = "Voltage Tolerance"

POWER_DISSIPATION_COLUMN = "Ptot [W] Power dissipation"


# ==========================================================
# SPEC HELPERS
# ==========================================================

def get_cell(
    part_row,
    column_name
):

    # ------------------------------------------------------
    # Exact column name first, then a stripped/case match in
    # case a spreadsheet round trip altered the header.
    # ------------------------------------------------------

    if column_name in part_row.index:
        return part_row[column_name]

    wanted = str(column_name).strip().lower()

    for column in part_row.index:

        if str(column).strip().lower() == wanted:
            return part_row[column]

    return None


def get_number(
    part_row,
    column_name
):

    value = get_cell(
        part_row,
        column_name
    )

    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None

    numeric_match = re.search(
        r"[\d.]+",
        str(value)
    )

    if not numeric_match:
        return None

    try:
        return float(numeric_match.group(0))

    except (ValueError, TypeError):
        return None


def get_grade(
    part_row
):

    # ------------------------------------------------------
    # Qualification reads "AEC-Q compliant", "AEC-Q101",
    # "Commercial Grade" or "Industrial Grade". The AEC-Q
    # ones are automotive, everything else is not.
    # ------------------------------------------------------

    qualification = safe_strip(
        get_cell(part_row, QUALIFICATION_COLUMN)
    ).lower()

    if "aec" in qualification:
        return "Automotive"

    return "Non-Automotive"


def get_direction(
    part_row
):

    # ------------------------------------------------------
    # Config states the direction when Diotec states one at
    # all. "Single" and "Array" describe the channel
    # arrangement and say nothing about direction.
    # ------------------------------------------------------

    config_text = safe_strip(
        get_cell(part_row, CONFIG_COLUMN)
    ).lower()

    config_text = re.sub(
        r"[^a-z]",
        "",
        config_text
    )

    if "unidire" in config_text or "single" in config_text:
        return "Unidirectional"

    if "bidire" in config_text:
        return "Bidirectional"

    return "-"


# ==========================================================
# PARSE ESD / TVS ROW
#
# Both sheets carry the same columns, except that only ESD
# states a channel count and ESD discharge voltage.
# ==========================================================

def parse_esd_tvs_row(
    part_row,
    source_name,
    part_number
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"Diotec {source_name}",
        "Grade": get_grade(part_row),
        "Direction": get_direction(part_row),
        "Channels": "1",
        "Package": safe_strip(get_cell(part_row, PACKAGE_COLUMN)),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-2": "-",
        "IEC 61000-4-5": "-",
        "Capacitance": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-"
    }

    # A blank channel count means a single channel part.

    channels = get_number(part_row, CHANNELS_COLUMN)

    if channels:
        specs_result["Channels"] = f"{int(channels)}"

    standoff = get_number(part_row, STANDOFF_COLUMN)

    if standoff is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{standoff:g} V"

    clamping = get_number(part_row, CLAMPING_COLUMN)

    if clamping is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamping:g} V"

    peak_pulse_power = get_number(part_row, PEAK_PULSE_POWER_COLUMN)

    if peak_pulse_power is not None:
        specs_result["Power Dissipation (Pd)"] = f"{peak_pulse_power:g} W"

    # IPPM, the peak pulse current the clamping voltage is
    # rated at, is the IEC 61000-4-5 surge rating.

    surge_current = get_number(part_row, SURGE_CURRENT_COLUMN)

    if surge_current is not None:
        specs_result["IEC 61000-4-5"] = f"{surge_current:g} A (8/20µs)"

    esd_contact = get_number(part_row, ESD_CONTACT_COLUMN)

    if esd_contact is not None:
        specs_result["IEC 61000-4-2"] = f"±{esd_contact:g} kV"

    capacitance = get_number(part_row, CAPACITANCE_COLUMN)

    if capacitance is not None:
        specs_result["Capacitance"] = f"{capacitance:.2f} pF"

    return specs_result


# ==========================================================
# PARSE ZENER ROW
# ==========================================================

def parse_zener_row(
    part_row,
    source_name,
    part_number
):

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": f"Diotec {source_name}",
        "Grade": get_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Package": safe_strip(get_cell(part_row, PACKAGE_COLUMN))
    }

    # ------------------------------------------------------
    # Vz comes from Vznom, falling back to the midpoint of
    # the Vzmin / Vzmax window.
    # ------------------------------------------------------

    zener_nom = get_number(part_row, ZENER_NOM_COLUMN)
    zener_min = get_number(part_row, ZENER_MIN_COLUMN)
    zener_max = get_number(part_row, ZENER_MAX_COLUMN)

    if zener_nom is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{zener_nom:g} V"

    elif zener_min is not None and zener_max is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{(zener_min + zener_max) / 2:g} V"

    elif zener_min is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{zener_min:g} V"

    elif zener_max is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{zener_max:g} V"

    # ------------------------------------------------------
    # A stated tolerance wins, otherwise the Vz window gives
    # it.
    # ------------------------------------------------------

    tolerance = get_number(part_row, TOLERANCE_COLUMN)

    if tolerance is not None:
        specs_result["Tolerance"] = f"±{tolerance:g}%"

    elif zener_min is not None and zener_max is not None and (zener_min + zener_max) > 0:
        specs_result["Tolerance"] = f"±{(zener_max - zener_min) / (zener_max + zener_min) * 100:.3g}%"

    power_dissipation = get_number(part_row, POWER_DISSIPATION_COLUMN)

    if power_dissipation is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power_dissipation:g} W"

    return specs_result


# ==========================================================
# FETCH SPECS FROM EXCEL
#
# diotec_dfs is {"ESD": df, "TVS": df, "Zener": df}.
# ==========================================================

def fetch_diotec_specs_from_excel(
    part_number,
    diotec_dfs
):

    if not part_number or not diotec_dfs:
        return None

    target = str(part_number).strip().lower()

    target_loose = re.sub(
        r"[\W_]+",
        "",
        target
    )

    search_priority = [
        "ESD",
        "TVS",
        "Zener"
    ]

    ordered_keys = [
        key
        for key in search_priority
        if key in diotec_dfs
    ] + [
        key
        for key in diotec_dfs
        if key not in search_priority
    ]

    for df_name in ordered_keys:

        df = diotec_dfs.get(df_name)

        if df is None or df.empty:
            continue

        df.columns = df.columns.str.strip()

        part_column = next(
            (
                column
                for column in df.columns
                if str(column).strip().lower() == PART_COLUMN.lower()
            ),
            None
        )

        if not part_column:
            continue

        column_values = df[part_column].astype(str).str.strip()

        part_rows = df[column_values.str.lower() == target]

        if part_rows.empty:

            loose_values = column_values.str.lower().str.replace(
                r"[\W_]+",
                "",
                regex=True
            )

            part_rows = df[loose_values == target_loose]

        if part_rows.empty:
            continue

        part_row = part_rows.iloc[0]

        exact_name = safe_strip(part_row[part_column])

        print(
            f"Found specs for Part '{part_number}' in Diotec database '{df_name}'."
        )

        if "zener" in df_name.lower():

            return parse_zener_row(
                part_row,
                df_name,
                exact_name
            )

        return parse_esd_tvs_row(
            part_row,
            df_name,
            exact_name
        )

    print(
        f"Part '{part_number}' not found in any Diotec database."
    )

    return None


# ==========================================================
# MAIN
# ==========================================================

if __name__ == "__main__":

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 30)

    try:

        # ======================================================
        # ESD
        # ======================================================

        esd_data = scrape_diotec_page(
            ESD_URL,
            "ESD"
        )

        esd_df = build_dataframe(
            esd_data,
            ESD_ALL_HEADERS
        )

        esd_df = keep_columns(
            esd_df,
            ESD_KEEP_COLUMNS
        )

        save_excel(
            esd_df,
            ESD_OUTPUT
        )


        # ======================================================
        # TVS
        # ======================================================

        tvs_data = scrape_diotec_page(
            TVS_URL,
            "TVS"
        )

        tvs_df = build_dataframe(
            tvs_data,
            TVS_HEADERS
        )

        # Same desired specs as ESD.
        # Anything not applicable to TVS is skipped.

        tvs_df = keep_columns(
            tvs_df,
            TVS_KEEP_COLUMNS
        )

        save_excel(
            tvs_df,
            TVS_OUTPUT
        )


        # ======================================================
        # ZENER
        # ======================================================

        zener_data = scrape_diotec_page(
            ZENER_URL,
            "ZENER"
        )

        zener_df = build_dataframe(
            zener_data,
            ZENER_HEADERS
        )

        # Only save the requested Zener columns

        zener_df = keep_columns(
            zener_df,
            ZENER_KEEP_COLUMNS
        )

        save_excel(
            zener_df,
            ZENER_OUTPUT
        )


    finally:

        driver.quit()


    # ==========================================================
    # FINISHED
    # ==========================================================

    print()
    print(
        "=========================================="
    )

    print(
        "ALL DIOTEC SCRAPES COMPLETE"
    )

    print(
        "=========================================="
    )

    print(
        f"ESD:   {ESD_OUTPUT}"
    )

    print(
        f"TVS:   {TVS_OUTPUT}"
    )

    print(
        f"Zener: {ZENER_OUTPUT}"
    )