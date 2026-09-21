"""
Comchip: scraper and spec parser.

Run this file as a script to scrape comchiptech.com into
comchip_esd_specs.xlsx, comchip_tvs_specs.xlsx and
comchip_zener_specs.xlsx. Importing it gives you fetch_comchip_specs(),
which reads those sheets - the import launches no browser and does not
require selenium to be installed.
"""

import re

import pandas as pd
from pathlib import Path
import time


# ==========================================================
# SETTINGS
# ==========================================================

ESD_ARRAY_URL = (
    "https://www.comchiptech.com/en/product/ESD"
    "?subcategory=ESD_Array"
)

TVS_URL = "https://www.comchiptech.com/en/product/TVS"

ZENER_URL = (
    "https://www.comchiptech.com/en/product/Zener"
    "?subcategory=Zener_Diodes"
)

BASE_FOLDER = Path(__file__).resolve().parent

ESD_OUTPUT = BASE_FOLDER / "comchip_esd_specs.xlsx"
TVS_OUTPUT = BASE_FOLDER / "comchip_tvs_specs.xlsx"
ZENER_OUTPUT = BASE_FOLDER / "comchip_zener_specs.xlsx"




# ==========================================================
# SPEC PARSING - used by applesauce.py
# ==========================================================

def _flat(text):
    """Header comparison form: lowercase, alphanumerics only."""
    return re.sub(r'[^a-z0-9]+', '', str(text).lower())


def _find_col(columns, prefixes=(), contains=()):
    """
    Column lookup by name - prefix matches first, in the order given, then
    substring matches.

    Anchoring matters on these sheets: 'Vrwm (V)', 'Vbr (min) (V)', 'Vbr (max)
    (V)' and 'Vc (V)' all contain 'v', and 'IPP (A)' sits beside 'Ir (uA)', so a
    plain substring search returns whichever column happens to come first.
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
    """Leading number in a cell as a float, or None."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    m = re.search(r'[-+]?\d*\.?\d+', str(val))
    return float(m.group(0)) if m else None


def _fmt(val, unit):
    n = _num(val)
    return f"{n:g} {unit}" if n is not None else "-"


def _direction(val):
    """'Uni' / 'Bi' column to the tool's wording."""
    text = str(val or "").strip().lower()
    if text.startswith("bi"):
        return "Bidirectional"
    if text.startswith("uni"):
        return "Unidirectional"
    return "-"


def _grade(val):
    """AEC-Q column: 'Yes' means automotive, '-' means not stated."""
    return "Automotive" if str(val or "").strip().lower().startswith("y") else "Commercial"


def _reheader(df):
    """
    Returns the frame with its real header row in place.

    If the loader read the sheet at the wrong offset the column names come back
    as data ('Unnamed: 0', or the contents of some other row), so the real
    header is found by looking for the row holding 'Part Number' and the frame
    is rebuilt from there. A frame that already has proper headers is returned
    untouched.
    """
    if df is None or df.empty:
        return df

    if _find_col(list(df.columns), prefixes=("partnumber", "partno"),
                 contains=("partnumber", "partno")) is not None:
        return df

    for i in range(min(15, len(df))):
        values = [str(v) for v in df.iloc[i].tolist()]
        if any(_flat(v) in ("partnumber", "partno") for v in values):
            fixed = df.iloc[i + 1:].copy()
            fixed.columns = [str(v).strip() for v in df.iloc[i].tolist()]
            return fixed.reset_index(drop=True)

    return df


def _match_comchip_row(part_input, df):
    """
    Finds a part's row in one Comchip sheet. Returns (row, part_col, frame) or
    (None, None, None) - the frame is returned too because it may have been
    re-headered, and the caller reads its columns.

    Comchip suffixes nearly every part '-HF' (halogen free), so a bare part
    number has to match the suffixed row and vice versa.
    """
    if df is None or df.empty:
        return None, None, None

    df = _reheader(df)

    part_col = _find_col(list(df.columns), prefixes=("partnumber", "partno"),
                         contains=("partnumber", "partno", "part"))
    if not part_col:
        print(f"Comchip: no part number column in this sheet. Headers are: {list(df.columns)}")
        return None, None, None

    norm_input = re.sub(r'[\W_]+', '', str(part_input).upper())
    if not norm_input:
        return None, None, None

    norm_series = df[part_col].astype(str).str.upper().str.replace(r'[\W_]+', '', regex=True)

    mask = (norm_series == norm_input)
    if not mask.any():                      # user typed the part without '-HF'
        mask = norm_series.str.startswith(norm_input, na=False)
    if not mask.any():                      # user typed a longer suffixed form
        mask = pd.Series([len(v) >= 6 and norm_input.startswith(v) for v in norm_series],
                         index=norm_series.index)
    if not mask.any():
        return None, None, None

    return df[mask].iloc[0], part_col, df


# Column order of each sheet, as the site's tables are laid out. The scraper
# writes real names only for the first two columns and calls the rest
# "Column 2", "Column 3" and so on, so when a name lookup fails these positions
# are what identify a field. Used only when the column count matches exactly,
# which keeps a differently shaped sheet from being read against the wrong
# positions.
ESD_LAYOUT = ["category", "part", "status", "aec", "dir", "ppp",
              "vrwm", "vc", "ipp", "cj", "esd", "package"]
TVS_LAYOUT = ["category", "part", "status", "aec", "dir", "ppp",
              "vrwm", "vbrmin", "vbrmax", "vc", "ipp", "ir", "package"]
ZENER_LAYOUT = ["category", "part", "status", "aec", "type",
                "pd", "vz", "iz", "tolerance", "package"]

# How each field is recognised by name when the headers have been filled in.
NAME_RULES = {
    "category":  (("category",),      ("category",)),
    "part":      (("partnumber", "partno"), ("partnumber", "partno")),
    "status":    (("status",),        ("status",)),
    "aec":       (("aec",),           ("aec",)),
    "dir":       (("unibi",),         ("unibi", "direction")),
    "type":      (("type",),          ("type",)),
    "ppp":       (("ppp",),           ("ppp", "peakpulse")),
    "vrwm":      (("vrwm",),          ("vrwm", "standoff")),
    "vbrmin":    (("vbrmin",),        ("vbrmin",)),
    "vbrmax":    (("vbrmax",),        ("vbrmax",)),
    "vc":        (("vc",),            ("clamp",)),
    "ipp":       (("ipp",),           ("ipp",)),
    "ir":        (("ir",),            ("ir",)),
    "cj":        (("cj",),            ("capacit",)),
    "esd":       (("esd",),           ("esd",)),
    "pd":        (("pd",),            ("power",)),
    "vz":        (("vz",),            ("vz",)),
    "iz":        (("iz",),            ("iz",)),
    "tolerance": (("tolerance", "tol"), ("tolerance",)),
    "package":   (("package",),       ("package",)),
}


def _resolve_columns(cols, layout):
    """
    Maps each field of a sheet to a real column: by header name where the name
    is meaningful, otherwise by its position in the layout.
    """
    positional_ok = (len(cols) == len(layout))
    resolved = {}
    for i, field in enumerate(layout):
        prefixes, contains = NAME_RULES.get(field, ((), ()))
        col = _find_col(cols, prefixes, contains)
        if col is None and positional_ok:
            col = cols[i]
        resolved[field] = col
    return resolved


def _parse_comchip_zener(row, cols, part_name):
    c = _resolve_columns(cols, ZENER_LAYOUT)
    c_pd, c_vz, c_tol = c["pd"], c["vz"], c["tolerance"]
    c_pkg, c_aec = c["package"], c["aec"]

    pd_mw = _num(_cell(row, c_pd))
    tol = _num(_cell(row, c_tol))
    return {
        "Device Name": part_name,
        "Source File": "Comchip Zener",
        "Package": _cell(row, c_pkg) or "-",
        "Grade": _grade(_cell(row, c_aec)),
        "Voltage - Reverse Standoff (Typ)": _fmt(_cell(row, c_vz), "V"),
        "Tolerance": f"±{tol:g}%" if tol is not None else "-",
        # Sheet states Pd in mW; the rest of the tool works in watts.
        "Power Dissipation (Pd)": f"{pd_mw / 1000:g} W" if pd_mw is not None else "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "Capacitance": "-",
        "Channels": "1",
        "Direction": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Price ($/ku)": "-",
    }


def _parse_comchip_tvs_esd(row, cols, part_name, source_label):
    # The TVS sheet carries Vbr min/max where the ESD sheet carries Cj and ESD,
    # so the layout is picked by which one the column count fits.
    layout = TVS_LAYOUT if len(cols) == len(TVS_LAYOUT) else ESD_LAYOUT
    c = _resolve_columns(cols, layout)
    c_pkg, c_aec, c_dir = c["package"], c["aec"], c["dir"]
    c_ppp, c_vrwm, c_vc = c["ppp"], c["vrwm"], c["vc"]
    c_ipp, c_cat = c["ipp"], c["category"]
    c_cj, c_esd = c.get("cj"), c.get("esd")

    esd_kv = _num(_cell(row, c_esd))
    ipp = _num(_cell(row, c_ipp))

    # The sheets carry no channel-count column. A part Comchip files under an
    # 'Array' category is multi-channel by definition but the count is not
    # stated, so it is left unknown rather than guessed at; everything else is
    # a single-line device.
    category = str(_cell(row, c_cat) or "")
    channels = "-" if "array" in category.lower() else "1"

    return {
        "Device Name": part_name,
        "Source File": source_label,
        "Package": _cell(row, c_pkg) or "-",
        "Grade": _grade(_cell(row, c_aec)),
        "Voltage - Reverse Standoff (Typ)": _fmt(_cell(row, c_vrwm), "V"),
        "Voltage - Clamping (Max) @ Ipp": _fmt(_cell(row, c_vc), "V"),
        "Power Dissipation (Pd)": _fmt(_cell(row, c_ppp), "W"),
        "Capacitance": f"{_num(_cell(row, c_cj)):.2f} pF" if _num(_cell(row, c_cj)) is not None else "-",
        "Channels": channels,
        "Direction": _direction(_cell(row, c_dir)),
        "IEC 61000-4-5": f"{ipp:g} A (8/20us)" if ipp is not None else "-",
        "IEC 61000-4-2": f"±{esd_kv:g} kV" if esd_kv is not None else "-",
        "Price ($/ku)": "-",
    }


def comchip_has_part(part_input, comchip_esd_df, comchip_tvs_df, comchip_zener_df):
    """True when the part appears in any Comchip sheet."""
    for df in (comchip_esd_df, comchip_tvs_df, comchip_zener_df):
        row, _, _ = _match_comchip_row(part_input, df)
        if row is not None:
            return True
    return False


def fetch_comchip_specs(part_input, comchip_esd_df, comchip_tvs_df, comchip_zener_df):
    """
    Looks a Comchip part up across the ESD, TVS and Zener sheets. Returns a
    specs dict or None.

    Zener rows report 'Comchip Zener' as their source so the caller routes them
    to the TI zener cross-reference rather than the TVS/ESD one.
    """
    sources = [
        (comchip_esd_df,   "ESD",   "Comchip ESD"),
        (comchip_tvs_df,   "TVS",   "Comchip TVS"),
        (comchip_zener_df, "Zener", "Comchip Zener"),
    ]

    for df, label, source_label in sources:
        row, part_col, df = _match_comchip_row(part_input, df)
        if row is None:
            continue

        cols = list(df.columns)
        part_name = str(row[part_col]).strip()
        print(f"Found '{part_name}' in Comchip {label} database.")

        # Axial-lead TVS are through-hole parts; TI has no equivalent, so they
        # are refused outright rather than crossed to a surface-mount part.
        category = str(_cell(row, _find_col(cols, prefixes=("category",), contains=("category",))) or "")
        if "axial" in category.lower():
            print(f"  '{part_name}' is {category.strip()} - not crossed, no TI equivalent.")
            return None

        status = str(_cell(row, _find_col(cols, prefixes=("status",), contains=("status",))) or "")
        if status and status.strip().lower() not in ("active", "-"):
            print(f"  Note: Comchip lists this part as '{status.strip()}'.")

        if label == "Zener":
            return _parse_comchip_zener(row, cols, part_name)
        return _parse_comchip_tvs_esd(row, cols, part_name, source_label)

    return None


# ==========================================================
# SCRAPING
# ==========================================================



# ==========================================================
# REMOVE POPUPS / OVERLAYS
# ==========================================================

def close_overlays():

    # Cookie popup
    try:
        yes_buttons = driver.find_elements(
            By.XPATH,
            "//*[normalize-space(text())='YES']"
        )

        for button in yes_buttons:
            if button.is_displayed():
                driver.execute_script(
                    "arguments[0].click();",
                    button
                )
                time.sleep(0.4)
                break

    except Exception:
        pass


    # Force-remove the "Drag the table..." popup
    try:
        driver.execute_script("""
            const elements = Array.from(
                document.querySelectorAll('*')
            );

            for (const el of elements) {

                const text = (el.innerText || '').trim();

                if (
                    text.includes(
                        'Drag the table to view the full information'
                    )
                    ||
                    text.includes(
                        'Click anywhere to close this message'
                    )
                ) {

                    let parent = el;

                    for (let i = 0; i < 7 && parent; i++) {

                        const style =
                            window.getComputedStyle(parent);

                        if (
                            style.position === 'fixed'
                            ||
                            style.position === 'absolute'
                        ) {
                            parent.remove();
                            break;
                        }

                        parent = parent.parentElement;
                    }
                }
            }

            document.querySelectorAll(
                '.modal-backdrop, ' +
                '.popup-mask, ' +
                '.mask, ' +
                '.overlay, ' +
                '[class*="overlay"], ' +
                '[class*="modal"]'
            ).forEach(el => {

                const text = el.innerText || '';

                if (
                    text.includes('Drag the table')
                    ||
                    text.includes(
                        'Click anywhere to close'
                    )
                ) {
                    el.remove();
                }
            });

            document.body.style.overflow = 'auto';
            document.documentElement.style.overflow = 'auto';
        """)

    except Exception:
        pass


    try:
        driver.find_element(
            By.TAG_NAME,
            "body"
        ).send_keys(Keys.ESCAPE)

    except Exception:
        pass


# ==========================================================
# READ CURRENT TABLE
# ==========================================================

def scrape_visible_rows():

    soup = BeautifulSoup(
        driver.page_source,
        "html.parser"
    )

    rows = soup.select(".bkProductItem")

    page_data = []

    for row in rows:

        cells = row.select(".td")

        values = [
            cell.get_text(" ", strip=True)
            for cell in cells
        ]

        if values and values[0]:
            page_data.append(values)

    return page_data


# ==========================================================
# CLICK NEXT PAGE
# ==========================================================

def click_next_page():

    close_overlays()

    driver.execute_script(
        "window.scrollTo(0, document.body.scrollHeight);"
    )

    time.sleep(0.8)

    close_overlays()

    buttons = driver.find_elements(
        By.CSS_SELECTOR,
        ".page-next"
    )

    for button in buttons:

        try:

            if not button.is_displayed():
                continue

            classes = (
                button.get_attribute("class")
                or ""
            ).lower()

            aria_disabled = (
                button.get_attribute("aria-disabled")
                or ""
            ).lower()

            if (
                "disabled" in classes
                or aria_disabled == "true"
            ):
                return False

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                button
            )

            time.sleep(0.2)

            close_overlays()

            refreshed = driver.find_elements(
                By.CSS_SELECTOR,
                ".page-next"
            )

            for new_button in refreshed:

                if new_button.is_displayed():

                    driver.execute_script(
                        "arguments[0].click();",
                        new_button
                    )

                    print("  Clicking Next...")

                    return True

        except Exception:
            continue

    return False


# ==========================================================
# SCRAPE ALL PAGES OF CURRENT TAB
# ==========================================================

def scrape_category(category_name):

    wait.until(
        lambda d:
        len(
            d.find_elements(
                By.CSS_SELECTOR,
                ".bkProductItem"
            )
        ) > 0
    )

    time.sleep(1.5)

    close_overlays()

    rows_collected = []
    seen_parts = set()
    page_num = 1

    expected_columns = None


    while True:

        print()
        print(
            f"{category_name} - Scraping page {page_num}..."
        )

        page_rows = scrape_visible_rows()

        current_parts = [
            row[0]
            for row in page_rows
            if row
        ]

        if page_rows and expected_columns is None:
            expected_columns = len(page_rows[0])

            print(
                f"  Columns detected: {expected_columns}"
            )

        added = 0

        for row in page_rows:

            if not row:
                continue

            part_number = row[0]

            if part_number not in seen_parts:

                seen_parts.add(part_number)
                rows_collected.append(row)

                added += 1


        print(
            f"  Rows visible: {len(page_rows)}"
        )
        print(
            f"  New rows added: {added}"
        )
        print(
            f"  Total unique parts: {len(seen_parts)}"
        )


        clicked = click_next_page()

        if not clicked:
            break


        page_changed = False

        for attempt in range(30):

            time.sleep(0.5)

            new_rows = scrape_visible_rows()

            new_parts = [
                row[0]
                for row in new_rows
                if row
            ]

            if (
                new_parts
                and new_parts != current_parts
            ):
                page_changed = True
                break


        # Retry once if page did not move
        if not page_changed:

            print(
                "  First click did not change page."
            )

            close_overlays()

            time.sleep(0.8)

            clicked_again = click_next_page()

            if clicked_again:

                for attempt in range(30):

                    time.sleep(0.5)

                    new_rows = scrape_visible_rows()

                    new_parts = [
                        row[0]
                        for row in new_rows
                        if row
                    ]

                    if (
                        new_parts
                        and new_parts != current_parts
                    ):
                        page_changed = True
                        break


        if not page_changed:
            print(
                "  Page did not change. "
                "Assuming last page."
            )
            break


        page_num += 1


        if page_num > 100:
            print("Safety stop reached.")
            break


    print()
    print(
        f"{category_name} complete: "
        f"{len(rows_collected)} parts."
    )

    return rows_collected


# ==========================================================
# CLICK A TAB BY ITS VISIBLE TEXT
# ==========================================================

def click_tab(tab_text):

    print()
    print(
        f"Switching to {tab_text}..."
    )

    close_overlays()

    driver.execute_script(
        "window.scrollTo(0, 0);"
    )

    time.sleep(0.8)

    close_overlays()

    old_rows = scrape_visible_rows()

    old_parts = [
        row[0]
        for row in old_rows
        if row
    ]


    candidates = driver.find_elements(
        By.XPATH,
        f"//*[normalize-space(text())='{tab_text}']"
    )

    clicked = False


    for tab in candidates:

        try:

            if not tab.is_displayed():
                continue

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                tab
            )

            time.sleep(0.3)

            close_overlays()


            refreshed = driver.find_elements(
                By.XPATH,
                f"//*[normalize-space(text())='{tab_text}']"
            )

            for new_tab in refreshed:

                if new_tab.is_displayed():

                    driver.execute_script(
                        "arguments[0].click();",
                        new_tab
                    )

                    clicked = True
                    break

            if clicked:
                break

        except Exception:
            continue


    if not clicked:
        raise Exception(
            f"Could not click tab: {tab_text}"
        )


    # Wait for product list to change
    changed = False

    for attempt in range(40):

        time.sleep(0.5)

        new_rows = scrape_visible_rows()

        new_parts = [
            row[0]
            for row in new_rows
            if row
        ]

        if (
            new_parts
            and new_parts != old_parts
        ):
            changed = True
            break


    if not changed:
        raise Exception(
            f"Clicked {tab_text}, but table did not change."
        )


    close_overlays()

    print(
        f"{tab_text} loaded successfully."
    )


# ==========================================================
# GET TABLE HEADER NAMES FROM PAGE
# ==========================================================

def get_table_headers():

    soup = BeautifulSoup(
        driver.page_source,
        "html.parser"
    )

    # Try likely header structures
    header_candidates = soup.select(
        ".product-table .th, "
        ".table .th, "
        ".bkProductTable .th, "
        ".th"
    )

    headers = []

    for cell in header_candidates:

        text = cell.get_text(
            " ",
            strip=True
        )

        if text:
            headers.append(text)


    # Remove duplicate adjacent header labels
    cleaned = []

    for value in headers:

        if not cleaned or cleaned[-1] != value:
            cleaned.append(value)


    return cleaned


# ==========================================================
# SAVE CATEGORY SET
# ==========================================================

def save_combined_file(
    category_results,
    output_file
):

    """
    category_results:
    [
        ("ESD Array", rows),
        ("ESD Diodes", rows)
    ]
    """

    combined = []

    max_columns = 0

    for category, rows in category_results:

        for row in rows:
            max_columns = max(
                max_columns,
                len(row)
            )


    # Generic headers if the exact table schema differs
    headers = [
        "Category",
        "Part Number"
    ]

    for i in range(
        2,
        max_columns + 1
    ):
        headers.append(
            f"Column {i}"
        )


    for category, rows in category_results:

        for row in rows:

            padded = row + (
                [""] *
                (max_columns - len(row))
            )

            combined.append(
                [category] + padded
            )


    df = pd.DataFrame(
        combined,
        columns=headers
    )

    df = df.drop_duplicates().reset_index(
        drop=True
    )

    try:

        df.to_excel(
            output_file,
            index=False
        )

    except PermissionError:

        print()
        print(
            f"ERROR: {output_file.name} "
            f"is open in Excel."
        )
        print(
            "Close it and run the script again."
        )

        raise


    print()
    print(
        f"Saved {len(df)} rows to:"
    )
    print(output_file)


# ==========================================================
# MAIN
# ==========================================================


if __name__ == "__main__":
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.common.keys import Keys
    from bs4 import BeautifulSoup

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 30)

    try:

        # ======================================================
        # ESD
        # ======================================================

        print()
        print("######################################")
        print("COMCHIP ESD")
        print("######################################")

        driver.get(
            ESD_ARRAY_URL
        )

        esd_array = scrape_category(
            "ESD Array"
        )

        click_tab(
            "ESD Diodes"
        )

        esd_diodes = scrape_category(
            "ESD Diodes"
        )

        save_combined_file(
            [
                ("ESD Array", esd_array),
                ("ESD Diodes", esd_diodes)
            ],
            ESD_OUTPUT
        )


        # ======================================================
        # TVS
        # ======================================================

        print()
        print("######################################")
        print("COMCHIP TVS")
        print("######################################")

        driver.get(
            TVS_URL
        )

        wait.until(
            lambda d:
            len(
                d.find_elements(
                    By.CSS_SELECTOR,
                    ".bkProductItem"
                )
            ) > 0
        )

        time.sleep(1.5)

        close_overlays()


        # We explicitly click TVS SMD first,
        # so we know which tab we're scraping.
        try:
            click_tab(
                "TVS SMD"
            )
        except Exception:
            # If TVS SMD is already active, the table
            # won't change, so just continue.
            print(
                "TVS SMD appears to already be active."
            )


        tvs_smd = scrape_category(
            "TVS SMD"
        )


        click_tab(
            "TVS Axial Lead"
        )

        tvs_axial = scrape_category(
            "TVS Axial Lead"
        )


        save_combined_file(
            [
                ("TVS SMD", tvs_smd),
                (
                    "TVS Axial Lead",
                    tvs_axial
                )
            ],
            TVS_OUTPUT
        )


        # ======================================================
        # ZENER
        # ======================================================

        print()
        print("######################################")
        print("COMCHIP ZENER")
        print("######################################")

        driver.get(
            ZENER_URL
        )

        zener_rows = scrape_category(
            "Zener Diodes"
        )


        save_combined_file(
            [
                (
                    "Zener Diodes",
                    zener_rows
                )
            ],
            ZENER_OUTPUT
        )


    finally:

        driver.quit()


    print()
    print("======================================")
    print("ALL COMCHIP SCRAPES COMPLETE")
    print("======================================")

    print(
        f"ESD file:   {ESD_OUTPUT}"
    )

    print(
        f"TVS file:   {TVS_OUTPUT}"
    )

    print(
        f"Zener file: {ZENER_OUTPUT}"
    )