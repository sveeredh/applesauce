import time
import re
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

LESHAN_MAX_PLAUSIBLE_TOLERANCE_PCT = 20.0

PAGES = [
    {
        "name": "ESD",
        "url": "https://en.lrc.cn/product/esd.html",
        "output": "leshan_esd_specs.xlsx",
    },
    {
        "name": "TVS",
        "url": "https://en.lrc.cn/product/tvs.html",
        "output": "leshan_tvs_specs.xlsx",
    },
    {
        "name": "Zener",
        "url": "https://en.lrc.cn/product/zenerdiode.html",
        "output": "leshan_zener_specs.xlsx",
    },
]

PAGE_WAIT = 1.2
TIMEOUT = 20


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def make_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    driver.set_page_load_timeout(60)
    return driver


def close_popups(driver):
    candidates = [
        (
            By.XPATH,
            "//button[contains("
            "translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
        ),
        (
            By.XPATH,
            "//button[contains("
            "translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'agree')]",
        ),
        (
            By.XPATH,
            "//a[contains("
            "translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
        ),
        (
            By.XPATH,
            "//a[contains("
            "translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'agree')]",
        ),
        (
            By.XPATH,
            "//*[self::button or self::a][normalize-space()='关闭']",
        ),
        (
            By.XPATH,
            "//*[self::button or self::a][normalize-space()='Close']",
        ),
    ]

    for by, locator in candidates:
        try:
            elements = driver.find_elements(by, locator)

            for el in elements:
                if el.is_displayed() and el.is_enabled():
                    driver.execute_script(
                        "arguments[0].click();",
                        el,
                    )
                    time.sleep(0.3)

        except Exception:
            pass


def find_product_table(driver):
    """
    Find the LRC product table.

    Generic enough to work for:
      - ESD
      - TVS
      - Zener
    """
    tables = driver.find_elements(By.TAG_NAME, "table")

    best_table = None
    best_score = -1

    for table in tables:
        try:
            text = clean_text(table.text).upper()

            score = 0

            for token in (
                "DEVICE",
                "DATA SHEET",
                "PACKAGE",
                "VRWM",
                "VBR",
                "VWM",
                "VZ",
                "PD",
                "IPP",
            ):
                if token in text:
                    score += 1

            row_count = len(
                table.find_elements(By.CSS_SELECTOR, "tbody tr")
            )

            if row_count:
                score += 2

            if score > best_score:
                best_score = score
                best_table = table

        except StaleElementReferenceException:
            continue

    if best_table is None or best_score < 3:
        raise RuntimeError(
            "Could not find the LRC product table."
        )

    return best_table


def get_headers(table):
    """
    Get the actual column header row while ignoring
    the site's filter-control row.
    """
    header_rows = table.find_elements(
        By.CSS_SELECTOR,
        "thead tr",
    )

    if not header_rows:
        header_rows = [
            row
            for row in table.find_elements(
                By.CSS_SELECTOR,
                "tr",
            )
            if row.find_elements(By.TAG_NAME, "th")
        ]

    best_headers = []
    best_score = -1

    for row in header_rows:
        cells = row.find_elements(
            By.CSS_SELECTOR,
            "th, td",
        )

        headers = [
            clean_text(cell.text)
            for cell in cells
        ]

        joined = " ".join(headers).upper()

        score = 0

        for token in (
            "DEVICE",
            "MARKING",
            "AECQ",
            "DATA SHEET",
            "PACKAGE",
            "VRWM",
            "VBR",
            "VZ",
            "PD",
        ):
            if token in joined:
                score += 1

        if score > best_score:
            best_score = score
            best_headers = headers

    if not best_headers:
        raise RuntimeError(
            "Could not determine product table headers."
        )

    # Make blank or duplicate headers unique.
    seen = {}
    unique_headers = []

    for i, header in enumerate(
        best_headers,
        start=1,
    ):
        name = header if header else f"Column {i}"

        if name in seen:
            seen[name] += 1
            name = f"{name} {seen[name]}"
        else:
            seen[name] = 1

        unique_headers.append(name)

    return unique_headers


def get_data_rows(table):
    tbody_rows = table.find_elements(
        By.CSS_SELECTOR,
        "tbody tr",
    )

    if tbody_rows:
        return tbody_rows

    return [
        row
        for row in table.find_elements(
            By.CSS_SELECTOR,
            "tr",
        )
        if row.find_elements(By.TAG_NAME, "td")
    ]


def get_datasheet_url(cell):
    """
    Get the actual datasheet URL.

    IMPORTANT:
    This is used ONLY for the DATA SHEET column.

    It will NOT replace the Device / part-number text.
    """
    links = cell.find_elements(
        By.TAG_NAME,
        "a",
    )

    for link in links:
        href = clean_text(
            link.get_attribute("href")
        )

        if not href:
            continue

        if href.lower().startswith("javascript:"):
            continue

        if href.lower().startswith(
            ("http://", "https://")
        ):
            return href

    return ""


def extract_current_page(driver):
    table = find_product_table(driver)

    headers = get_headers(table)

    rows = get_data_rows(table)

    records = []

    for row in rows:
        try:
            cells = row.find_elements(
                By.TAG_NAME,
                "td",
            )

            if not cells:
                continue

            values = []

            for col_index, cell in enumerate(cells):
                visible_text = clean_text(
                    cell.text
                )

                if col_index < len(headers):
                    header = headers[col_index]
                else:
                    header = ""

                header_upper = header.upper()

                # ------------------------------------------
                # DEVICE / PART NUMBER
                # ------------------------------------------
                # Always use the visible text.
                #
                # Even if the part number itself is wrapped
                # in a PDF hyperlink, DO NOT replace it
                # with the URL.
                # ------------------------------------------
                if (
                    header_upper == "DEVICE"
                    or (
                        col_index == 0
                        and headers
                        and "DEVICE"
                        in headers[0].upper()
                    )
                ):
                    value = visible_text

                # ------------------------------------------
                # DATA SHEET
                # ------------------------------------------
                # Only this column gets the actual URL.
                # ------------------------------------------
                elif (
                    "DATA SHEET" in header_upper
                    or "DATASHEET" in header_upper
                ):
                    value = get_datasheet_url(cell)

                # ------------------------------------------
                # ALL OTHER SPECS
                # ------------------------------------------
                else:
                    value = visible_text

                values.append(value)

            if not any(values):
                continue

            # Pad short rows.
            if len(values) < len(headers):
                values += [""] * (
                    len(headers) - len(values)
                )

            # Preserve unexpected extra columns.
            elif len(values) > len(headers):
                old_length = len(headers)

                for i in range(
                    old_length,
                    len(values),
                ):
                    headers.append(
                        f"Extra Column {i + 1}"
                    )

                for existing in records:
                    for header in headers:
                        existing.setdefault(
                            header,
                            "",
                        )

            record = {
                headers[i]:
                    values[i]
                    if i < len(values)
                    else ""
                for i in range(len(headers))
            }

            first_value = (
                clean_text(values[0])
                if values
                else ""
            )

            # Ignore control/filter rows.
            if (
                first_value
                and first_value.lower()
                not in {
                    "device",
                    "reset filters",
                    "add/remove parameters",
                }
            ):
                records.append(record)

        except StaleElementReferenceException:
            continue

    return headers, records


def first_row_signature(driver):
    """
    Used to detect whether clicking Next Page
    actually changed the product list.
    """
    try:
        table = find_product_table(driver)

        rows = get_data_rows(table)

        for row in rows:
            cells = row.find_elements(
                By.TAG_NAME,
                "td",
            )

            if cells:
                values = [
                    clean_text(cell.text)
                    for cell in cells[:3]
                ]

                if any(values):
                    return " | ".join(values)

    except Exception:
        pass

    return ""


def find_next_button(driver):
    """
    Find the site's Next Page button.
    """
    xpaths = [
        (
            "//*[self::a or self::button]"
            "[normalize-space()='Next Page']"
        ),
        (
            "//*[self::a or self::button]"
            "[contains(normalize-space(.), "
            "'Next Page')]"
        ),
        (
            "//*[contains(@class,'next') "
            "and "
            "(self::a or self::button or self::li)]"
        ),
        (
            "//*[contains(@aria-label,'Next')]"
        ),
        (
            "//*[contains(@title,'Next')]"
        ),
    ]

    for xpath in xpaths:
        try:
            matches = driver.find_elements(
                By.XPATH,
                xpath,
            )

            for el in matches:
                if el.is_displayed():
                    return el

        except Exception:
            continue

    return None


def next_is_disabled(element):
    try:
        element_class = (
            element.get_attribute("class")
            or ""
        ).lower()

        aria_disabled = (
            element.get_attribute(
                "aria-disabled"
            )
            or ""
        ).lower()

        disabled_attribute = (
            element.get_attribute(
                "disabled"
            )
        )

        parent_class = ""

        try:
            parent = element.find_element(
                By.XPATH,
                "..",
            )

            parent_class = (
                parent.get_attribute("class")
                or ""
            ).lower()

        except Exception:
            pass

        return (
            "disabled" in element_class
            or "disabled" in parent_class
            or aria_disabled == "true"
            or disabled_attribute is not None
        )

    except StaleElementReferenceException:
        return False


def go_to_next_page(driver):
    old_signature = first_row_signature(
        driver
    )

    next_button = find_next_button(
        driver
    )

    if next_button is None:
        print(
            "No 'Next Page' button found. "
            "Reached final page."
        )
        return False

    if next_is_disabled(next_button):
        print(
            "'Next Page' is disabled. "
            "Reached final page."
        )
        return False

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        next_button,
    )

    time.sleep(0.4)

    try:
        next_button.click()

    except (
        ElementClickInterceptedException,
        StaleElementReferenceException,
    ):
        try:
            next_button = find_next_button(
                driver
            )

            driver.execute_script(
                "arguments[0].click();",
                next_button,
            )

        except Exception:
            return False

    # Wait until the product rows change.
    end_time = time.time() + TIMEOUT

    while time.time() < end_time:
        time.sleep(0.4)

        new_signature = first_row_signature(
            driver
        )

        if (
            new_signature
            and new_signature
            != old_signature
        ):
            time.sleep(PAGE_WAIT)
            return True

    # Final-page fallback:
    # page may change and immediately disable Next.
    button = find_next_button(driver)

    if (
        button is not None
        and next_is_disabled(button)
    ):
        time.sleep(PAGE_WAIT)
        return True

    print(
        "Page did not appear to change "
        "after clicking Next Page."
    )

    return False


def scrape_category(
    driver,
    name,
    url,
    output_file,
):
    print("\n" + "=" * 70)

    print(
        f"Scraping Leshan/LRC {name}"
    )

    print(
        f"URL: {url}"
    )

    print("=" * 70)

    driver.get(url)

    WebDriverWait(
        driver,
        TIMEOUT,
    ).until(
        lambda d:
            find_product_table(d)
            is not None
    )

    close_popups(driver)

    all_records = []
    all_headers = []

    seen_signatures = set()

    page_num = 1

    while True:
        print(
            f"\nReading {name} "
            f"page {page_num}..."
        )

        WebDriverWait(
            driver,
            TIMEOUT,
        ).until(
            lambda d:
                find_product_table(d)
                is not None
        )

        signature = first_row_signature(
            driver
        )

        # Prevent infinite pagination loops.
        if (
            signature
            and signature
            in seen_signatures
        ):
            print(
                "Detected a repeated page. "
                "Stopping to prevent duplicates."
            )
            break

        if signature:
            seen_signatures.add(
                signature
            )

        headers, records = (
            extract_current_page(driver)
        )

        for header in headers:
            if header not in all_headers:
                all_headers.append(
                    header
                )

        print(
            f"Rows found: {len(records)}"
        )

        if records:
            device_column = all_headers[0]

            print(
                "First part: "
                f"{records[0].get(device_column, '')}"
            )

            print(
                "Last part:  "
                f"{records[-1].get(device_column, '')}"
            )

        all_records.extend(records)

        # Keep clicking Next Page until the end.
        if not go_to_next_page(driver):
            break

        page_num += 1

    if not all_records:
        raise RuntimeError(
            f"No {name} product rows "
            "were extracted."
        )

    df = pd.DataFrame(
        all_records
    )

    # Keep columns in the same order
    # they appeared on the website.
    ordered_columns = [
        column
        for column in all_headers
        if column in df.columns
    ]

    ordered_columns += [
        column
        for column in df.columns
        if column not in ordered_columns
    ]

    df = df[ordered_columns]

    # Remove completely blank rows.
    df = df.replace(
        r"^\s*$",
        pd.NA,
        regex=True,
    )

    df = df.dropna(
        how="all"
    )

    # Remove completely blank columns.
    df = df.dropna(
        axis=1,
        how="all",
    )

    # Remove exact duplicate product rows.
    df = df.drop_duplicates(
    ).reset_index(
        drop=True
    )

    df = df.fillna("")

    output_path = Path(
        output_file
    ).resolve()

    df.to_excel(
        output_path,
        index=False,
    )

    print(
        f"\n{name} complete."
    )

    print(
        f"Pages scraped: {page_num}"
    )

    print(
        f"Unique rows saved: {len(df)}"
    )

    print(
        f"Columns saved: {len(df.columns)}"
    )

    print(
        f"Output: {output_path}"
    )


def main():
    driver = make_driver()

    try:
        # --------------------------------
        # ESD
        # --------------------------------
        # https://en.lrc.cn/product/esd.html
        # saves:
        # leshan_esd_specs.xlsx
        #
        # --------------------------------
        # TVS
        # --------------------------------
        # https://en.lrc.cn/product/tvs.html
        # saves:
        # leshan_tvs_specs.xlsx
        #
        # --------------------------------
        # ZENER
        # --------------------------------
        # https://en.lrc.cn/product/zenerdiode.html
        # saves:
        # leshan_zener_specs.xlsx
        # --------------------------------

        for page in PAGES:
            scrape_category(
                driver=driver,
                name=page["name"],
                url=page["url"],
                output_file=page["output"],
            )

    finally:
        driver.quit()

    print(
        "\nAll Leshan/LRC "
        "scrapes complete."
    )

# ============================================================
# SPEC LOOKUP (used by applesauce, not by the scraper)
# ============================================================

def _norm_header(col_header):
    """Header text with spaces, underscores and case removed, for matching."""
    return re.sub(r'[\s_]+', '', str(col_header)).upper()


def _get_header_by_prefix(row, *prefixes):
    """
    First column header whose normalized text starts with one of the prefixes.

    Leshan nests parameter names inside other headers - "IR @VRWM" contains
    "VRWM", "VC(V)" contains "V" - so a plain substring search returns
    whichever column happens to sit first. Anchoring at the start of the
    header keeps each parameter on its own column.
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
    Tells apart headers that share a name ("VZ (V) Min" vs "VZ (V) Nom").
    """
    for tokens in token_sets:
        for col_header in row.index:
            header = _norm_header(col_header)
            if all(_norm_header(token) in header for token in tokens):
                return col_header
    return None


def _leshan_text(value):
    """Cell text, with pandas blanks reduced to an empty string."""
    if value is None:
        return ""
    if not isinstance(value, str) and pd.isna(value):
        return ""
    return clean_text(value)


def _numeric_or_none(value):
    """First number in a cell, or None if the cell is blank or non-numeric."""
    text = _leshan_text(value)
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
    """Cell value for the first header matching one of the prefixes."""
    header = _get_header_by_prefix(row, *prefixes)
    return row.get(header) if header is not None else None


def _value_by_tokens(row, *token_sets):
    header = _get_header_by_tokens(row, *token_sets)
    return row.get(header) if header is not None else None


def _leshan_grade(part_row):
    """
    All three sheets carry a "Qualified to AECQ101" column holding Y or a
    blank, so the grade is read per row rather than per file.
    """
    aec = _leshan_text(_value_by_prefix(part_row, "Qualified to AECQ101", "Qualified")).upper()
    return "Automotive" if aec.startswith("Y") else "Commercial"


def _leshan_package(part_row):
    return _leshan_text(_value_by_prefix(part_row, "Package")) or "-"

def _leshan_esd_direction(part_number):
    """
    Leshan/LRC LESD naming:

    C immediately after the voltage = Bidirectional

    Examples:
        LESD8D5.0T5G       -> Unidirectional
        LESD8D5.0N3T5G     -> Unidirectional
        LESD8D5.0CT5G      -> Bidirectional
        LESD8D3.3CN3T5G    -> Bidirectional
    """
    part = str(part_number or "").upper().strip()

    # C directly after the numeric voltage means bidirectional.
    if re.search(r"\d+(?:\.\d+)?C", part):
        return "Bidirectional"

    return "Unidirectional"


def _leshan_esd_channels(part_number):
    """
    Leshan/LRC LESD naming:

    N3 = 3-pin / 2-line protection topology.

    Current verified LESD parts without N3 are single-channel.
    """
    part = str(part_number or "").upper().strip()

    if "N3" in part:
        return "2"

    return "1"

def _parse_leshan_esd_row(part_row, found_df_name, part_number):
    """
    Parses a row from leshan_esd_specs.

    Columns: Device, Device (marking), Qualified to AECQ101, DATA SHEET,
    VRWM, IR @VRWM, VBR (V), CJ (pF), VC @IPP, IPP tp=8/20us (A),
    Power tp=8/20us (WATT), Package.

    The sheet states no directionality and no ESD contact rating.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _leshan_grade(part_row),
        "Direction": _leshan_esd_direction(part_number),
        "Channels": _leshan_esd_channels(part_number),
        "Package": _leshan_package(part_row),
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

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "VC@", "VC ("))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "IPP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20us)"

    cap_val = _numeric_or_none(_value_by_prefix(part_row, "CJ"))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR@", "IR ("))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    power_val = _numeric_or_none(_value_by_prefix(part_row, "Power"))
    if power_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{power_val:g} W"

    return specs_result


def _parse_leshan_tvs_row(part_row, found_df_name, part_number):
    """
    Parses a row from leshan_tvs_specs.

    Columns: Device, Bi-Directional, Qualified to AECQ101, DATA SHEET,
    VWM(V), VBR(V) Min @ It, VBR(V) Max @ It, It(mA), VC(V), Ippm(A),
    IR(uA), Package.

    The sheet states no capacitance, no ESD contact rating and no peak pulse
    power.
    """
    # The Bi-Directional column is a flag, not a two-value field: a Y marks a
    # bidirectional part and a blank cell means unidirectional.
    bidi = _leshan_text(_value_by_prefix(part_row, "Bi-Directional", "BiDirectional")).upper()
    direction = "Bidirectional" if bidi.startswith("Y") else "Unidirectional"

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _leshan_grade(part_row),
        "Direction": direction,
        "Channels": "1",
        "Package": _leshan_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    vwm_val = _numeric_or_none(_value_by_prefix(part_row, "VWM"))
    if vwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vwm_val:g} V"

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "VC("))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ippm_val = _numeric_or_none(_value_by_prefix(part_row, "Ippm"))
    if ippm_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ippm_val:g} A (8/20us)"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "IR("))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def _parse_leshan_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from leshan_zener_specs.

    Columns: Device, Device Marking, Qualified to AECQ101, DATA SHEET,
    VZ (V) Min, VZ (V) Nom, VZ (V) Max, Vz@IZT mA, PD(mW), Package,
    More Documents.

    There is no tolerance column, so it comes out of the VZ min/nom/max
    window.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _leshan_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",
        "Package": _leshan_package(part_row),
    }

    # The three VZ columns share a name, so each is picked by its Min/Nom/Max
    # token rather than by position.
    vz_nom = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "NOM")))
    vz_min = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "MIN")))
    vz_max = _numeric_or_none(_value_by_tokens(part_row, ("VZ", "MAX")))

    if vz_nom is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"
    elif vz_min is not None and vz_max is not None:
        vz_nom = (vz_min + vz_max) / 2
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_nom:g} V"

    # Tolerance from the window around the nominal, guarded so a series-wide
    # span is not mistaken for one part's tolerance.
    if (vz_nom is not None and vz_nom > 0
            and vz_min is not None and vz_max is not None
            and vz_min <= vz_nom <= vz_max):
        spread_pct = max(vz_nom - vz_min, vz_max - vz_nom) / vz_nom * 100
        if 0 < spread_pct <= LESHAN_MAX_PLAUSIBLE_TOLERANCE_PCT:
            specs_result["Tolerance"] = f"±{spread_pct:.3g}%"

    # PD is stated in mW.
    pd_val = _numeric_or_none(_value_by_prefix(part_row, "PD("))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val / 1000.0:g} W"

    return specs_result


def fetch_leshan_specs_from_excel(part_number, leshan_dfs):
    """
    Searches across Leshan DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS vs. ESD).
    """
    if part_number is None or not leshan_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["Zener", "ESD", "TVS"]

    all_df_keys = list(leshan_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["device", "part number", "product"]

    for df_name in ordered_search_keys:
        df = leshan_dfs[df_name]
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
        print(f"Part '{part_number}' not found in any Leshan database.")
        return None

    print(f"Found specs for Part '{part_number}' in Leshan database '{found_df_name}'.")

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_leshan_zener_row(part_row, found_df_name, part_number)
    elif "tvs" in name_lower:
        return _parse_leshan_tvs_row(part_row, found_df_name, part_number)
    else:
        return _parse_leshan_esd_row(part_row, found_df_name, part_number)
    
if __name__ == "__main__":
    main()