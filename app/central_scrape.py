"""
Central Semiconductor: scraper and spec parser.

Running this file as a script scrapes my.centralsemi.com into
central_zener_specs.xlsx. Importing it gives you fetch_central_specs(),
which reads that spreadsheet - the import does not launch a browser, and
does not require selenium to be installed.
"""

import re
import pandas as pd
from pathlib import Path


URL = (
    "https://my.centralsemi.com/paraSearch/"
    "parametricSearch_v1_1.php?"
    "task=bldPartTable2&PGroup=10&subGroup=60"
)

TOTAL_PAGES = 71

OUTPUT_FILE = Path(__file__).resolve().parent / "central_zener_specs.xlsx"


# ============================================================
# Spec parsing - used by applesauce.py
# ============================================================

def _flat(text):
    """Header comparison form: lowercase, alphanumerics only."""
    return re.sub(r'[^a-z0-9]+', '', str(text).lower())


def _find_col(columns, prefixes=(), contains=()):
    """
    Column lookup by name - prefix matches first, in the order given, then
    substring matches.

    Anchoring matters here: 'VZ NOM', 'VZ MIN', 'VZ MAX' and 'VZ @ IZT' all
    contain 'vz', so a plain substring search would return whichever comes
    first rather than the one asked for.
    """
    flat_cols = [(_flat(c), c) for c in columns]
    for prefix in prefixes:
        for flat, col in flat_cols:
            if flat.startswith(prefix):
                return col
    for keyword in contains:
        for flat, col in flat_cols:
            if keyword in flat:
                return col
    return None


def _cell(row, col):
    """Raw value at a column, or None when missing/blank."""
    if col is None or col not in row.index:
        return None
    val = row[col]
    if pd.isna(val):
        return None
    text = str(val).strip()
    return None if text.lower() in ("", "nan", "none", "-") else text


def _num(val):
    """Leading number in a cell as a float, or None. '14.25V' -> 14.25"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    m = re.search(r'[-+]?\d*\.?\d+', str(val))
    return float(m.group(0)) if m else None


def _watts(val):
    """
    Power as watts. Central writes '50W', but also '500mW' on the small
    packages, so the unit decides the scale rather than being assumed.
    """
    n = _num(val)
    if n is None:
        return None
    if re.search(r'm\s*w', str(val), re.IGNORECASE):
        return n / 1000.0
    return n


def _tolerance(nom, vmin, vmax):
    """
    Central states VZ MIN/MAX rather than a tolerance percentage, so derive it:
    15V with 14.25/15.75 is +/-5%. The wider side wins when they differ.
    """
    if nom in (None, 0):
        return None
    spans = [abs(v - nom) / nom * 100 for v in (vmin, vmax) if v is not None]
    return max(spans) if spans else None


# Mounting styles TI has an equivalent for. Anything else - through-hole,
# axial, bare die and so on - has no TI cross and is refused outright rather
# than crossed to a surface-mount part in a package it cannot be assembled in.
CROSSABLE_MOUNTING = ("surface mount", "smd ruggedized devices")


def _mounting_is_crossable(mounting):
    """
    True when the Mounting cell names a surface-mount style.

    Compared with separators stripped, so 'Surface Mount', 'surface-mount' and
    the leading newline the scraper leaves on some cells all read alike.
    """
    flat = re.sub(r'[^a-z0-9]+', '', str(mounting or '').lower())
    return any(re.sub(r'[^a-z0-9]+', '', ok) in flat for ok in CROSSABLE_MOUNTING)


def _match_central_row(part_input, central_zener_df, verbose=True):
    """
    Finds the sheet row for a part. Returns (row, part_col) or (None, None).

    Split out from fetch_central_specs so a caller can ask whether a part is
    Central's without also asking for its specs - see central_has_part.
    """
    if central_zener_df is None or central_zener_df.empty:
        return None, None

    cols = list(central_zener_df.columns)
    part_col = _find_col(cols, prefixes=("partnumber", "partno"),
                         contains=("partnumber", "partno", "part"))
    if not part_col:
        if verbose:
            print(f"Central: no part number column found. Headers are: {cols}")
        return None, None

    norm_input = re.sub(r'[\W_]+', '', str(part_input).upper())
    if not norm_input:
        return None, None

    norm_series = central_zener_df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)

    # Exact, then a same-base-part match (suffixes such as -TR reel codes).
    # Nothing looser: a wrong zener is a wrong voltage.
    match_mask = (norm_series == norm_input)
    if not match_mask.any():
        match_mask = norm_series.str.startswith(norm_input, na=False)
    if not match_mask.any():
        match_mask = pd.Series([norm_input.startswith(v) and len(v) >= 6
                                for v in norm_series], index=norm_series.index)
    if not match_mask.any():
        return None, None

    return central_zener_df[match_mask].iloc[0], part_col


def central_has_part(part_input, central_zener_df):
    """
    True when the part is in Central's sheet at all, crossable or not.

    The dispatcher uses this to stop a part that failed the mounting check from
    falling through to the DigiKey CSVs and being crossed from there anyway.
    """
    row, _ = _match_central_row(part_input, central_zener_df, verbose=False)
    return row is not None


def fetch_central_specs(part_input, central_zener_df):
    """
    Looks up a Central Semiconductor part. Returns a specs dict, or None when
    the part is absent or is not a surface-mount style TI can cross to.

    'Source File' says Zener so the caller routes this to the TI zener
    cross-reference rather than the TVS/ESD one.
    """
    row, part_col = _match_central_row(part_input, central_zener_df)
    if row is None:
        return None

    cols = list(central_zener_df.columns)
    exact_part_name = str(row[part_col]).strip()
    print(f"Found '{exact_part_name}' in Central Semiconductor zener database.")

    c_case    = _find_col(cols, prefixes=("case",), contains=("case", "package"))
    c_mount   = _find_col(cols, prefixes=("mounting",), contains=("mounting",))
    c_config  = _find_col(cols, prefixes=("configuration",), contains=("configuration",))
    c_pd      = _find_col(cols, prefixes=("pdmax", "pd"), contains=("powerdissipation",))
    c_vz_nom  = _find_col(cols, prefixes=("vznom",), contains=("vznom",))
    c_vz_min  = _find_col(cols, prefixes=("vzmin",), contains=("vzmin",))
    c_vz_max  = _find_col(cols, prefixes=("vzmax",), contains=("vzmax",))
    c_vr      = _find_col(cols, prefixes=("vr",), contains=("vr",))

    vz_nom = _num(_cell(row, c_vz_nom))
    vz_min = _num(_cell(row, c_vz_min))
    vz_max = _num(_cell(row, c_vz_max))

    # VZ NOM is blank on some rows; the midpoint of MIN/MAX stands in for it.
    if vz_nom is None and vz_min is not None and vz_max is not None:
        vz_nom = (vz_min + vz_max) / 2

    tol = _tolerance(vz_nom, vz_min, vz_max)
    pd_watts = _watts(_cell(row, c_pd))

    package = _cell(row, c_case) or "-"
    config = _cell(row, c_config) or ""
    mounting = _cell(row, c_mount) or ""

    # Refuse non-surface-mount parts before any cross is attempted.
    if not _mounting_is_crossable(mounting):
        print(f"Central part '{exact_part_name}' is {mounting.strip() or 'of unstated mounting'} - "
              f"not surface mount, so no TI alternatives apply.")
        return None

    # 'Single: Power' style configuration - anything not single is an array.
    channels = "1"
    ch_match = re.search(r'(dual|triple|quad|single)', config, re.IGNORECASE)
    if ch_match:
        channels = {"single": "1", "dual": "2", "triple": "3", "quad": "4"}[ch_match.group(1).lower()]

    return {
        "Device Name": exact_part_name,
        "Source File": "Central Zener",
        "Package": package,
        "Mounting": mounting,
        "Configuration": config,
        "Grade": "Commercial",
        "Voltage - Reverse Standoff (Typ)": f"{vz_nom:g} V" if vz_nom is not None else "-",
        # Rounded - a derived tolerance like 5.128% is false precision, and the
        # scorer only reads the leading number anyway.
        "Tolerance": f"±{round(tol, 1):g}%" if tol is not None else "-",
        "Power Dissipation (Pd)": f"{pd_watts:g} W" if pd_watts is not None else "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Capacitance": "-",
        "Channels": channels,
        "Direction": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Price ($/ku)": "-",
    }


# ============================================================
# Scraping - only runs when this file is executed directly
# ============================================================

def scrape_current_dom(driver):
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(driver.page_source, "html.parser")

    found_rows = []

    # Look through ALL <tr> elements.
    # This is safer than relying on altcol1 / altcol2.
    for row in soup.find_all("tr"):

        first_cell = row.select_one(".col01")

        # Not a product row
        if first_cell is None:
            continue

        row_data = []

        # Central Semiconductor uses col01 ... col14
        for i in range(1, 15):

            cell = row.select_one(f".col{i:02d}")

            if cell is None:
                value = ""
            else:
                value = cell.get_text(" ", strip=True)

            row_data.append(value)

        # Must have a part number
        if row_data[0]:
            found_rows.append(tuple(row_data))

    return found_rows


if __name__ == "__main__":
    import time
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 30)

    # A set automatically removes identical duplicate rows
    all_rows = set()

    try:

        print("Opening Central Semiconductor...")
        driver.get(URL)

        wait.until(
            EC.presence_of_element_located((By.ID, "currPage"))
        )

        time.sleep(2)

        print("Starting page scrape...")
        print()


        # ========================================================
        # Go through all 71 pages
        # ========================================================

        for page_num in range(1, TOTAL_PAGES + 1):

            print(f"Scraping page {page_num}/{TOTAL_PAGES}...")

            # ----------------------------------------------------
            # For pages after page 1, change pagination dropdown
            # ----------------------------------------------------

            if page_num > 1:

                success = False

                for attempt in range(1, 6):

                    try:

                        # IMPORTANT:
                        # Find the dropdown FRESH every attempt.
                        # Never reuse an old Selenium element.
                        dropdown_element = wait.until(
                            EC.presence_of_element_located(
                                (By.ID, "currPage")
                            )
                        )

                        Select(dropdown_element).select_by_value(
                            str(page_num)
                        )

                        # ------------------------------------------------
                        # Don't use Selenium's Select object to check it.
                        #
                        # JavaScript gets the NEW currPage element directly
                        # from the DOM, so stale-element errors are avoided.
                        # ------------------------------------------------

                        wait.until(
                            lambda d:
                            d.execute_script(
                                "return document.getElementById('currPage')"
                                ".value;"
                            ) == str(page_num)
                        )

                        # Allow AJAX/table rendering to finish
                        time.sleep(1.5)

                        success = True
                        break

                    except Exception as e:

                        print(
                            f"  Page change attempt "
                            f"{attempt}/5 failed."
                        )

                        time.sleep(2)

                if not success:
                    print(
                        f"  WARNING: Could not reliably switch "
                        f"to page {page_num}."
                    )
                    print("  Retrying by refreshing page...")

                    driver.refresh()

                    wait.until(
                        EC.presence_of_element_located(
                            (By.ID, "currPage")
                        )
                    )

                    time.sleep(2)

                    # Try this page one more time after refresh
                    dropdown_element = wait.until(
                        EC.presence_of_element_located(
                            (By.ID, "currPage")
                        )
                    )

                    Select(dropdown_element).select_by_value(
                        str(page_num)
                    )

                    time.sleep(2)


            # ----------------------------------------------------
            # Scrape everything currently stored in page DOM
            # ----------------------------------------------------

            rows_now = scrape_current_dom(driver)

            before = len(all_rows)

            all_rows.update(rows_now)

            added = len(all_rows) - before

            print(
                f"  Rows visible in DOM: {len(rows_now)}"
            )

            print(
                f"  New unique rows added: {added}"
            )

            print(
                f"  Total unique rows collected: {len(all_rows)}"
            )


            # ----------------------------------------------------
            # SAVE PROGRESS every 5 pages
            # ----------------------------------------------------

            if page_num % 5 == 0:

                temp_df = pd.DataFrame(
                    list(all_rows),
                    columns=[
                        "Part Number",
                        "Mounting",
                        "Case",
                        "Configuration",
                        "Column 5",
                        "PD MAX",
                        "IZM MAX",
                        "VZ NOM",
                        "VZ MIN",
                        "VZ MAX",
                        "VZ @ IZT",
                        "ZZT MAX",
                        "ZZT @ IZT",
                        "IR MAX"
                    ]
                )

                try:

                    temp_df.to_excel(
                        OUTPUT_FILE,
                        index=False
                    )

                    print(
                        f"  Progress saved "
                        f"({len(temp_df)} rows)"
                    )

                except PermissionError:

                    print()
                    print(
                        "  WARNING: central_specs.xlsx is open "
                        "in Excel."
                    )

                    print(
                        "  Close it so progress can be saved."
                    )

            print()


    finally:

        driver.quit()


    # ============================================================
    # Final Excel output
    # ============================================================

    print("Browser closed.")
    print()

    columns = [
        "Part Number",
        "Mounting",
        "Case",
        "Configuration",
        "Column 5",
        "PD MAX",
        "IZM MAX",
        "VZ NOM",
        "VZ MIN",
        "VZ MAX",
        "VZ @ IZT",
        "ZZT MAX",
        "ZZT @ IZT",
        "IR MAX"
    ]


    df = pd.DataFrame(
        list(all_rows),
        columns=columns
    )


    # Sort by part number so spreadsheet isn't randomly ordered
    df = df.sort_values(
        by="Part Number",
        kind="stable"
    ).reset_index(drop=True)


    print(f"Final unique rows collected: {len(df)}")


    try:

        df.to_excel(
            OUTPUT_FILE,
            index=False
        )

    except PermissionError:

        print()
        print("ERROR:")
        print("central_specs.xlsx is currently open in Excel.")
        print("Close the workbook and rerun the script.")
        raise


    print()
    print("========================================")
    print("DONE")
    print("========================================")
    print(f"Rows saved: {len(df)}")
    print(f"File: {OUTPUT_FILE}")