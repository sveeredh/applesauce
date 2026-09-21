import re
import time
from urllib.parse import urljoin

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# CONFIG
# ============================================================

PRODUCTS = [
    {
        "name": "ESD",
        "url": (
            "https://www.yenyo.com.tw/Pages/format.aspx?"
            "INTID=0e3d3ae1-c626-49e8-b28a-1ef9430dd30e"
            "&PKID=4954c987-95e5-49e2-b728-18c1a211b811"
        ),
        "output": "yenyo_esd_specs.xlsx",
        "sheet": "Yenyo ESD",
        "headers": [
            "Part Number",
            "Data Sheet",
            "Family",
            "Status",
            "Package",
            "Configuration",
            "AEC-Q101",
            "VRWM (V)",
            "IR Max. (μA)",
            "VBR Min. (V)",
            "CJ Typ. (pF)",
            "CJ Max. (pF)",
            "IPP @ 8/20μs (A)",
            "PPP (W)",
        ],
    },

    {
        "name": "TVS",
        "url": (
            "https://www.yenyo.com.tw/Pages/format.aspx?"
            "INTID=1ef1e173-c461-4ad5-9166-9b937eea0e1c"
            "&PKID=e9f00006-ad87-4017-bfc5-861b806f2f84"
        ),
        "output": "yenyo_tvs_specs.xlsx",
        "sheet": "Yenyo TVS",
        "headers": [
            "Part Number",
            "Data Sheet",
            "Family",
            "Status",
            "Configuration",
            "Package",
            "AEC-Q101 Qualified",
            "VRWM (V)",
            "VBR Min. (V)",
            "VBR Max. (V)",
            "IR (μA)",
            "PPK (W)",
            "IPP (A)",
            "VC clamp (V)",
            "TJ Max. (°C)",
        ],
    },

    {
        "name": "Zener",
        "url": (
            "https://www.yenyo.com.tw/Pages/format.aspx?"
            "INTID=ca03c0cc-efd2-4bf3-a7ae-e6418712fe43"
            "&PKID=6c7a9302-882f-406c-8df7-44cd221117da"
            "&reset"
        ),
        "output": "yenyo_zener_specs.xlsx",
        "sheet": "Yenyo Zener",
        "headers": [
            "Part Number",
            "Data Sheet",
            "Family",
            "Status",
            "Package",
            "AEC-Q101 Qualified",
            "VZ Min. (V)",
            "VZ Nom. (V)",
            "VZ Max. (V)",
            "IZT (mA)",
            "ZZT @ IZT (Ω)",
            "ZZK @ IZK (Ω)",
            "IZK (mA)",
            "IR (μA)",
            "VR (V)",
            "PD (mW)",
            "TJ Max. (°C)",
            "Tolerance ± (%)",
        ],
    },
]


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def get_text_content(element):
    """
    Use textContent instead of Selenium .text.

    Yenyo uses an internally scrollable table. Selenium .text can
    return blank for rows outside the currently visible area, while
    textContent reads the value directly from the DOM.
    """

    try:
        return clean_text(
            element.get_attribute("textContent")
        )
    except Exception:
        return ""

def yenyo_channel_count(part_number, package):
    """
    Infer Yenyo ESD protected-channel count.
    Returns None when it cannot be determined safely.
    """

    pn = str(part_number or "").upper().strip()
    pkg = str(package or "").upper().strip()

    # 1-channel discrete devices
    if re.search(r"-2$", pkg):
        return 1

    if pkg.startswith((
        "SOD-123",
        "SOD-323",
        "SOD-523",
    )):
        return 1

    # Known multi-channel Yenyo families
    if pn.startswith("YEUST26"):
        return 4

    if pn.startswith("YEU38C9"):
        return 8

    if pn.startswith("YEU33CA"):
        return 6

    if pn.startswith("YEU41FA"):
        return 6

    if pn.startswith("YEU55KF"):
        return 14

    if pn.startswith("YEDSOP8"):
        return 2

    return None

def make_driver():
    options = Options()

    # Leave Chrome visible so the internal scrolling can be seen.
    # Uncomment for headless mode:
    # options.add_argument("--headless=new")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-gpu")

    service = Service(
        ChromeDriverManager().install()
    )

    return webdriver.Chrome(
        service=service,
        options=options
    )


# ============================================================
# FIND PRODUCT TABLE
# ============================================================

def find_product_table(driver):
    """
    Yenyo has nested/layout tables.

    The actual product table is one of the tables with the largest
    number of <tr> elements, so choose the table containing the
    greatest number of rows.
    """

    tables = driver.find_elements(
        By.TAG_NAME,
        "table"
    )

    print(
        f"Tables found on page: {len(tables)}"
    )

    best_table = None
    best_count = 0

    for number, table in enumerate(
        tables,
        start=1
    ):

        rows = table.find_elements(
            By.CSS_SELECTOR,
            "tr"
        )

        count = len(rows)

        print(
            f"  Table {number}: "
            f"{count} rows"
        )

        if count > best_count:
            best_count = count
            best_table = table

    if best_table is None:
        raise RuntimeError(
            "Could not find product table."
        )

    print(
        f"Selected table with "
        f"{best_count} rows."
    )

    return best_table


# ============================================================
# FIND INTERNAL SCROLLBAR
# ============================================================

def find_scroll_container(driver, table):

    scroll_container = driver.execute_script(
        """
        let el = arguments[0];

        while (el) {

            const style =
                window.getComputedStyle(el);

            const canScroll =
                el.scrollHeight >
                el.clientHeight + 5;

            if (
                canScroll &&
                (
                    style.overflowY === 'auto' ||
                    style.overflowY === 'scroll' ||
                    style.overflowY === 'overlay'
                )
            ) {
                return el;
            }

            el = el.parentElement;
        }

        return null;
        """,
        table
    )

    return scroll_container


# ============================================================
# SCROLL ENTIRE YENYO TABLE
# ============================================================

def scroll_entire_table(driver, table):

    print(
        "Looking for internal table scrollbar..."
    )

    container = find_scroll_container(
        driver,
        table
    )

    if container is None:
        print(
            "No internal scrollbar found. "
            "Will still read all DOM rows."
        )
        return

    info = driver.execute_script(
        """
        return {
            height: arguments[0].scrollHeight,
            client: arguments[0].clientHeight
        };
        """,
        container
    )

    print(
        f"Scrollbar height: "
        f"{info['client']} visible / "
        f"{info['height']} total"
    )

    # Start at absolute top.
    driver.execute_script(
        "arguments[0].scrollTop = 0;",
        container
    )

    time.sleep(0.4)

    last_position = -1
    iteration = 0

    while True:

        iteration += 1

        info = driver.execute_script(
            """
            return {
                top: arguments[0].scrollTop,
                height: arguments[0].scrollHeight,
                client: arguments[0].clientHeight
            };
            """,
            container
        )

        current = int(info["top"])
        height = int(info["height"])
        client = int(info["client"])

        maximum = max(
            height - client,
            0
        )

        percentage = (
            100
            if maximum == 0
            else (
                current /
                maximum *
                100
            )
        )

        print(
            f"  Scroll {iteration}: "
            f"{current}/{maximum} "
            f"({percentage:.1f}%)"
        )

        if current >= maximum - 2:
            break

        if current == last_position:
            break

        last_position = current

        next_position = min(
            current + int(client * 0.8),
            maximum
        )

        driver.execute_script(
            """
            arguments[0].scrollTop =
                arguments[1];

            arguments[0].dispatchEvent(
                new Event(
                    'scroll',
                    {bubbles: true}
                )
            );
            """,
            container,
            next_position
        )

        time.sleep(0.1)

    # Force absolute bottom.
    driver.execute_script(
        """
        arguments[0].scrollTop =
            arguments[0].scrollHeight;

        arguments[0].dispatchEvent(
            new Event(
                'scroll',
                {bubbles: true}
            )
        );
        """,
        container
    )

    time.sleep(0.8)

    print(
        "Reached bottom of product table."
    )


# ============================================================
# CHECK WHETHER A ROW IS AN ACTUAL PRODUCT
# ============================================================

def is_product_row(cells, expected_columns):
    """
    Determine whether this is a real part-number row.

    We intentionally DO NOT check for a Y prefix because:
      ESD   -> YED...
      TVS   -> 1.5KE..., SMA..., SMC...
      Zener -> 1N4728A, BZT52..., etc.
    """

    if not cells:
        return False

    # Product rows should have approximately the number
    # of columns expected for this table.
    if len(cells) < expected_columns - 1:
        return False

    first = get_text_content(
        cells[0]
    )

    if not first:
        return False

    first_lower = first.lower()

    # Header / filter row.
    if first_lower in {
        "part number",
        "part number ↕",
        "part number↕",
        "part number refresh",
    }:
        return False

    # Filter/header cells may contain SELECT widgets.
    selects = cells[0].find_elements(
        By.TAG_NAME,
        "select"
    )

    if selects:
        return False

    # Avoid any weird layout text.
    if "recommended part number" in first_lower:
        return False

    return True


# ============================================================
# PARSE TABLE
# ============================================================

def parse_rows(
    table,
    url,
    headers
):

    rows = table.find_elements(
        By.CSS_SELECTOR,
        "tr"
    )

    print(
        f"Rows in table DOM: "
        f"{len(rows)}"
    )

    records = []

    for row_number, row in enumerate(
        rows,
        start=1
    ):

        # Direct children only so nested table cells
        # don't get mixed together.
        cells = row.find_elements(
            By.XPATH,
            "./td"
        )

        if not is_product_row(
            cells,
            len(headers)
        ):
            continue

        values = []

        for column_index, cell in enumerate(
            cells
        ):

            # ------------------------------------------------
            # COLUMN 2 = DATASHEET LINK
            # ------------------------------------------------

            if column_index == 1:

                datasheet = ""

                links = cell.find_elements(
                    By.TAG_NAME,
                    "a"
                )

                for link in links:

                    href = link.get_attribute(
                        "href"
                    )

                    if href:
                        datasheet = urljoin(
                            url,
                            href
                        )
                        break

                values.append(
                    datasheet
                )

            # ------------------------------------------------
            # NORMAL SPEC COLUMN
            # ------------------------------------------------

            else:

                values.append(
                    get_text_content(cell)
                )

        # Normalize width.
        if len(values) < len(headers):

            values.extend(
                [""] * (
                    len(headers) -
                    len(values)
                )
            )

        elif len(values) > len(headers):

            part = values[0]

            print(
                f"WARNING: {part} has "
                f"{len(values)} cells; "
                f"expected {len(headers)}."
            )

            values = values[
                :len(headers)
            ]

        records.append(values)

    return records


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(
    df,
    output_file,
    sheet_name
):

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

        ws = writer.book[
            sheet_name
        ]

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = (
            ws.dimensions
        )

        # Reasonable column widths.
        for column_cells in ws.columns:

            column_letter = (
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
                    len(value)
                )

            ws.column_dimensions[
                column_letter
            ].width = min(
                max(
                    max_length + 2,
                    12
                ),
                60
            )


# ============================================================
# SCRAPE ONE PRODUCT CATEGORY
# ============================================================

def scrape_product(
    driver,
    product
):

    name = product["name"]
    url = product["url"]
    headers = product["headers"]
    output = product["output"]
    sheet = product["sheet"]

    print()
    print(
        "=" * 70
    )
    print(
        f"SCRAPING YENYO {name}"
    )
    print(
        "=" * 70
    )

    print(
        f"Opening Yenyo {name} page..."
    )

    driver.get(url)

    WebDriverWait(
        driver,
        30
    ).until(
        EC.presence_of_all_elements_located(
            (By.TAG_NAME, "table")
        )
    )

    time.sleep(2)

    # --------------------------------------------------------
    # FIND TABLE
    # --------------------------------------------------------

    table = find_product_table(
        driver
    )

    # --------------------------------------------------------
    # SCROLL INTERNAL BAR
    # --------------------------------------------------------

    scroll_entire_table(
        driver,
        table
    )

    # Site may modify DOM while scrolling,
    # so find the table again.
    table = find_product_table(
        driver
    )

    # --------------------------------------------------------
    # PARSE
    # --------------------------------------------------------

    records = parse_rows(
        table,
        url,
        headers
    )

    print(
        f"Raw product rows collected: "
        f"{len(records)}"
    )

    if not records:
        raise RuntimeError(
            f"No {name} products extracted."
        )

    df = pd.DataFrame(
        records,
        columns=headers
    )

    # --------------------------------------------------------
    # ADD YENYO ESD CHANNEL COUNT
    # --------------------------------------------------------

    if name == "ESD":
        df["Channels"] = df.apply(
            lambda row: yenyo_channel_count(
                row["Part Number"],
                row["Package"]
            ),
            axis=1
        )
    # --------------------------------------------------------
    # CLEAN
    # --------------------------------------------------------

    df["Part Number"] = (
        df["Part Number"]
        .astype(str)
        .str.strip()
    )

    df = df[
        df["Part Number"].ne("")
    ].copy()

    before = len(df)

    df = (
        df
        .drop_duplicates(
            subset=["Part Number"],
            keep="first"
        )
        .reset_index(drop=True)
    )

    duplicates = (
        before - len(df)
    )

    if duplicates:
        print(
            f"Removed "
            f"{duplicates} duplicate rows."
        )

    # --------------------------------------------------------
    # RESULTS
    # --------------------------------------------------------

    print()
    print(
        f"TOTAL UNIQUE {name} PARTS: "
        f"{len(df)}"
    )

    print(
        "\nFirst 5:"
    )

    for part in (
        df["Part Number"]
        .head(5)
    ):
        print(
            f"  {part}"
        )

    print(
        "\nLast 5:"
    )

    for part in (
        df["Part Number"]
        .tail(5)
    ):
        print(
            f"  {part}"
        )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    print(
        f"\nSaving to "
        f"'{output}'..."
    )

    save_excel(
        df,
        output,
        sheet
    )

    print(
        f"Saved {len(df)} rows "
        f"to {output}"
    )

    return df


# ============================================================
# MAIN
# ============================================================

def main():

    driver = make_driver()

    try:

        results = {}

        for product in PRODUCTS:

            try:

                df = scrape_product(
                    driver,
                    product
                )

                results[
                    product["name"]
                ] = len(df)

            except Exception as exc:

                print()
                print(
                    f"ERROR scraping "
                    f"{product['name']}: "
                    f"{exc}"
                )

                results[
                    product["name"]
                ] = "FAILED"

        print()
        print(
            "=" * 70
        )
        print(
            "YENYO SCRAPE COMPLETE"
        )
        print(
            "=" * 70
        )

        for name, count in results.items():

            print(
                f"{name}: {count}"
            )

        print()
        print(
            "Output files:"
        )
        print(
            "  yenyo_esd_specs.xlsx"
        )
        print(
            "  yenyo_tvs_specs.xlsx"
        )
        print(
            "  yenyo_zener_specs.xlsx"
        )

    finally:

        driver.quit()


# ============================================================
# SPEC LOOKUP (used by applesauce, not by the scraper)
# ============================================================

# A VZ window wider than this around the nominal is a series span rather than
# the part's own tolerance, so it is not read as a tolerance.
YENYO_MAX_PLAUSIBLE_TOLERANCE_PCT = 20.0


def _norm_header(col_header):
    """Header text with spaces, underscores and case removed, for matching."""
    return re.sub(r'[\s_]+', '', str(col_header)).upper()


def _get_header_by_prefix(row, *prefixes):
    """
    First column header whose normalized text starts with one of the prefixes.

    Yenyo nests parameter names inside other headers - "IPP @ 8/20us (A)" and
    "PPP (W)" both contain PP, "VR (V)" sits alongside "VRWM (V)" - so a plain
    substring search returns whichever column happens to sit first. Anchoring
    at the start of the header keeps each parameter on its own column.
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
    Tells apart headers that share a name ("VZ Min. (V)" vs "VZ Nom. (V)",
    "CJ Typ. (pF)" vs "CJ Max. (pF)").
    """
    for tokens in token_sets:
        for col_header in row.index:
            header = _norm_header(col_header)
            if all(_norm_header(token) in header for token in tokens):
                return col_header
    return None


def _yenyo_text(value):
    """Cell text, with pandas blanks reduced to an empty string."""
    if value is None:
        return ""
    if not isinstance(value, str) and pd.isna(value):
        return ""
    return clean_text(value)


def _numeric_or_none(value):
    """First number in a cell, or None if the cell is blank or non-numeric."""
    text = _yenyo_text(value)
    if text in ("", "/", "-", "--", "N/A", "NA"):
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


def _yenyo_grade(part_row):
    """
    All three sheets carry an AEC-Q101 column. It holds Y / Yes on a qualified
    part and a blank otherwise, so the grade is read per row.
    """
    aec = _yenyo_text(_value_by_prefix(part_row, "AEC-Q101", "AECQ101", "AEC")).upper()
    return "Automotive" if aec.startswith("Y") else "Commercial"


def _yenyo_package(part_row):
    return _yenyo_text(_value_by_prefix(part_row, "Package")) or "-"


def _yenyo_direction(part_row):
    """Configuration states the directionality outright where it is given."""
    config = _yenyo_text(_value_by_prefix(part_row, "Configuration")).lower()
    config = re.sub(r'[^a-z]', '', config)
    if config.startswith("uni"):
        return "Unidirectional"
    if config.startswith("bi"):
        return "Bidirectional"
    return "-"


def _parse_yenyo_esd_row(part_row, found_df_name, part_number):
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yenyo_grade(part_row),
        "Direction": _yenyo_direction(part_row),
        "Channels": "-",
        "Package": _yenyo_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    # Read channel count from scraped Excel first.
    channel_val = _numeric_or_none(
        _value_by_prefix(part_row, "Channels")
    )

    # Fallback to Yenyo naming/package inference if needed.
    if channel_val is None:
        channel_val = yenyo_channel_count(
            part_number,
            _yenyo_package(part_row)
        )

    if channel_val is not None:
        specs_result["Channels"] = str(int(channel_val))

    vrwm_val = _numeric_or_none(
        _value_by_prefix(part_row, "VRWM")
    )
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"


    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "IPP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20us)"

    # CJ Typ. is the figure a designer selects on; CJ Max. is the fallback.
    cap_val = _numeric_or_none(_value_by_tokens(part_row, ("CJ", "TYP"), ("CJ", "MAX")))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR "))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    ppp_val = _numeric_or_none(_value_by_prefix(part_row, "PPP"))
    if ppp_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp_val:g} W"

    return specs_result


def _parse_yenyo_tvs_row(part_row, found_df_name, part_number):
    """
    Parses a row from yenyo_tvs_specs.

    Columns: Part Number, Data Sheet, Family, Status, Configuration, Package,
    AEC-Q101 Qualified, VRWM (V), VBR Min. (V), VBR Max. (V), IR (uA),
    PPK (W), IPP (A), VC clamp (V), TJ Max. (C).

    The sheet states no capacitance, no ESD contact rating and no channel
    count.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yenyo_grade(part_row),
        "Direction": _yenyo_direction(part_row),
        "Channels": "1",
        "Package": _yenyo_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    vrwm_val = _numeric_or_none(_value_by_prefix(part_row, "VRWM"))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "VC"))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "IPP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20us)"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR "))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    ppk_val = _numeric_or_none(_value_by_prefix(part_row, "PPK"))
    if ppk_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppk_val:g} W"

    return specs_result


def _parse_yenyo_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from yenyo_zener_specs.

    Columns: Part Number, Data Sheet, Family, Status, Package,
    AEC-Q101 Qualified, VZ Min. (V), VZ Nom. (V), VZ Max. (V), IZT (mA),
    ZZT @ IZT (Ohm), ZZK @ IZK (Ohm), IZK (mA), IR (uA), VR (V), PD (mW),
    TJ Max. (C), Tolerance +- (%).

    Tolerance is a real column here, so the VZ window is only used as a
    fallback when that column is blank.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _yenyo_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",  # not present in the Zener file headers
        "Package": _yenyo_package(part_row),
    }

    # The three VZ columns share a name, so each is picked by its token.
    vz_nom = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "NOM")))
    vz_min = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "MIN")))
    vz_max = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "MAX")))

    if vz_nom is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"
    elif vz_min is not None and vz_max is not None:
        vz_nom = (vz_min + vz_max) / 2
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"

    # The stated tolerance wins; the VZ window is the fallback, guarded so a
    # series-wide span is not mistaken for one part's tolerance.
    tol_val = _numeric_or_none(_value_by_prefix(part_row, "Tolerance"))
    if tol_val is not None:
        specs_result["Tolerance"] = f"±{tol_val:g}%"
    elif (vz_nom is not None and vz_nom > 0
            and vz_min is not None and vz_max is not None
            and vz_min <= vz_nom <= vz_max):
        spread_pct = max(vz_nom - vz_min, vz_max - vz_nom) / vz_nom * 100
        if 0 < spread_pct <= YENYO_MAX_PLAUSIBLE_TOLERANCE_PCT:
            specs_result["Tolerance"] = f"±{spread_pct:.3g}%"

    # PD is stated in mW.
    pd_val = _numeric_or_none(_value_by_prefix(part_row, "PD "))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val / 1000.0:g} W"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR "))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def fetch_yenyo_specs_from_excel(part_number, yenyo_dfs):
    """
    Searches across Yenyo DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS vs. ESD).
    """
    if part_number is None or not yenyo_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["Zener", "ESD", "TVS"]

    all_df_keys = list(yenyo_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["part number", "mfr part", "product"]

    for df_name in ordered_search_keys:
        df = yenyo_dfs[df_name]
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
        print(f"Part '{part_number}' not found in any Yenyo database.")
        return None

    print(f"Found specs for Part '{part_number}' in Yenyo database '{found_df_name}'.")

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_yenyo_zener_row(part_row, found_df_name, part_number)
    elif "tvs" in name_lower:
        return _parse_yenyo_tvs_row(part_row, found_df_name, part_number)
    else:
        return _parse_yenyo_esd_row(part_row, found_df_name, part_number)
    
if __name__ == "__main__":
    main()