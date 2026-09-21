import re
import time
from urllib.parse import urljoin

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException,
    NoAlertPresentException,
    UnexpectedAlertPresentException,
)

from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# CONFIG
# ============================================================

BASE_URL = "https://productselect.wendell.com.tw/?site=nichtek"

ROWS_PER_PAGE = 100


PRODUCTS = [
    {
        "name": "ESD",
        "aliases": [
            "ESD Diode",
            "ESD Diodes",
            "ESD",
        ],
        "output": "nichtek_esd_specs.xlsx",
        "sheet": "Nichtek ESD",
    },

    {
        "name": "TVS",
        "aliases": [
            "TVS Diode",
            "TVS Diodes",
            "TVS",
        ],
        "output": "nichtek_tvs_specs.xlsx",
        "sheet": "Nichtek TVS",
    },

    {
        "name": "Zener",
        "aliases": [
            "Zener Diode",
            "Zener Diodes",
            "Zener",
        ],
        "output": "nichtek_zener_specs.xlsx",
        "sheet": "Nichtek Zener",
    },
]


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def text_content(element):
    try:
        return clean_text(
            element.get_attribute("textContent")
        )

    except Exception:
        return ""


def make_driver():
    options = Options()

    # Keep browser visible while debugging.
    # Uncomment later for headless:
    # options.add_argument("--headless=new")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-gpu")

    service = Service(
        ChromeDriverManager().install()
    )

    driver = webdriver.Chrome(
        service=service,
        options=options
    )

    driver.set_page_load_timeout(60)

    return driver


# ============================================================
# ALERT
# ============================================================

def dismiss_alert_if_present(driver):

    try:
        alert = driver.switch_to.alert

        print(
            f"Closing alert: {alert.text}"
        )

        alert.accept()

        time.sleep(0.25)

        return True

    except NoAlertPresentException:
        return False

    except Exception:
        return False


# ============================================================
# WAIT FOR INITIAL PAGE
# ============================================================

def wait_for_page(driver):

    print(
        "Waiting for NichTek product selector..."
    )

    WebDriverWait(
        driver,
        40
    ).until(
        EC.presence_of_element_located(
            (By.TAG_NAME, "body")
        )
    )

    end_time = time.time() + 30

    while time.time() < end_time:

        try:
            body = text_content(
                driver.find_element(
                    By.TAG_NAME,
                    "body"
                )
            )

            if (
                "Obtaining data" not in body
                and "{{item}}" not in body
            ):
                break

        except StaleElementReferenceException:
            pass

        time.sleep(0.25)

    time.sleep(1)


# ============================================================
# CATEGORY SELECT
# ============================================================

def find_category_select(driver):

    selects = driver.find_elements(
        By.TAG_NAME,
        "select"
    )

    print(
        f"\nSelect elements found: "
        f"{len(selects)}"
    )

    best = None
    best_score = -1

    for number, element in enumerate(
        selects,
        start=1
    ):

        try:
            selector = Select(element)

            options = [
                clean_text(option.text)
                for option in selector.options
            ]

        except Exception:
            continue

        print(
            f"\nDropdown {number} options:"
        )

        for option in options:
            print(
                f"  {repr(option)}"
            )

        joined = (
            " ".join(options)
            .lower()
        )

        score = 0

        for keyword in [
            "esd",
            "tvs",
            "zener",
            "mosfet",
            "diode",
        ]:

            if keyword in joined:
                score += 1

        if score > best_score:

            best = element
            best_score = score

    if (
        best is None
        or best_score <= 0
    ):
        raise RuntimeError(
            "Could not locate NichTek "
            "category dropdown."
        )

    print(
        "\nProduct category dropdown found."
    )

    return best


def choose_product_category(
    driver,
    product
):

    element = find_category_select(
        driver
    )

    selector = Select(element)

    aliases = [
        value.lower()
        for value
        in product["aliases"]
    ]

    candidates = []

    for option in selector.options:

        label = clean_text(
            option.text
        )

        lower = label.lower()

        score = 0

        for alias in aliases:

            if lower == alias:
                score += 100

            elif alias in lower:
                score += 10

        if score:

            candidates.append(
                (
                    score,
                    label,
                    clean_text(
                        option.get_attribute(
                            "value"
                        )
                    ),
                )
            )

    if not candidates:

        raise RuntimeError(
            f"Could not find category "
            f"{product['name']}."
        )

    candidates.sort(
        key=lambda x: x[0],
        reverse=True
    )

    _, label, value = (
        candidates[0]
    )

    print(
        f"\nSelecting category: "
        f"{label}"
    )

    if value:

        selector.select_by_value(
            value
        )

    else:

        selector.select_by_visible_text(
            label
        )

    driver.execute_script(
        """
        arguments[0].dispatchEvent(
            new Event(
                'change',
                {bubbles:true}
            )
        );
        """,
        element
    )

    time.sleep(1)

    wait_for_page(driver)


# ============================================================
# FIND DATA TABLE
# ============================================================

def find_data_table(
    driver,
    verbose=False
):

    tables = driver.find_elements(
        By.TAG_NAME,
        "table"
    )

    if verbose:

        print(
            f"\nTables found: "
            f"{len(tables)}"
        )

    best = None
    best_score = -1

    for number, table in enumerate(
        tables,
        start=1
    ):

        try:

            rows = table.find_elements(
                By.CSS_SELECTOR,
                "tr"
            )

            headers = table.find_elements(
                By.TAG_NAME,
                "th"
            )

            max_cells = 0

            for row in rows:

                try:

                    cells = row.find_elements(
                        By.XPATH,
                        "./td"
                    )

                    max_cells = max(
                        max_cells,
                        len(cells)
                    )

                except StaleElementReferenceException:
                    pass

            score = (
                len(rows) * 20
                + max_cells
                + len(headers)
            )

            if verbose:

                print(
                    f"  Table {number}: "
                    f"{len(rows)} rows, "
                    f"{len(headers)} headers, "
                    f"max {max_cells} cells"
                )

            if score > best_score:

                best_score = score
                best = table

        except StaleElementReferenceException:
            continue

    if best is None:

        raise RuntimeError(
            "Could not locate NichTek "
            "parametric table."
        )

    return best


# ============================================================
# HEADERS
# ============================================================

def get_headers(table):

    try:

        rows = table.find_elements(
            By.CSS_SELECTOR,
            "thead tr"
        )

        if not rows:

            rows = table.find_elements(
                By.CSS_SELECTOR,
                "tr"
            )[:5]

    except StaleElementReferenceException:

        return []

    best = []

    for row in rows:

        try:

            cells = row.find_elements(
                By.XPATH,
                "./th | ./td"
            )

        except StaleElementReferenceException:
            continue

        values = [
            text_content(cell)
            for cell in cells
        ]

        if (
            len(values) > len(best)
            and any(values)
        ):

            best = values

    final = []

    seen = {}

    for index, header in enumerate(
        best
    ):

        header = clean_text(
            header
        )

        if not header:

            if index == 0:
                header = "Select"

            elif index == 1:
                header = "Data Sheet"

            else:
                header = (
                    f"Column {index + 1}"
                )

        if header in seen:

            seen[header] += 1

            header = (
                f"{header} "
                f"{seen[header]}"
            )

        else:

            seen[header] = 1

        final.append(
            header
        )

    print(
        "\nDetected headers:"
    )

    for number, header in enumerate(
        final,
        start=1
    ):

        print(
            f"  {number}: {header}"
        )

    return final


# ============================================================
# PART NUMBER COLUMN
# ============================================================

def get_part_number_index(headers):

    for index, header in enumerate(
        headers
    ):

        h = (
            clean_text(header)
            .lower()
            .replace(" ", "")
        )

        if any(
            key in h
            for key in [
                "nichtekp/n",
                "partnumber",
                "partno",
            ]
        ):
            return index

    # NichTek layout:
    #
    # Select
    # Data Sheet
    # Nichtek P/N
    if len(headers) >= 3:
        return 2

    return None


# ============================================================
# "VIEWING X TO Y OF Z ROWS"
# ============================================================

VIEWING_PATTERN = re.compile(
    r"Viewing\s+"
    r"(\d+)\s+"
    r"to\s+"
    r"(\d+)\s+"
    r"of\s+"
    r"(\d+)\s+"
    r"rows",
    re.IGNORECASE
)


def find_viewing_element(driver):
    """
    Finds the smallest visible element whose own text contains:
        Viewing X to Y of Z rows
    """

    candidates = driver.find_elements(
        By.XPATH,
        "//*[contains("
        "translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),"
        "'viewing')]"
    )

    matches = []

    for element in candidates:

        try:

            if not element.is_displayed():
                continue

            text = clean_text(
                element.text
            )

            match = (
                VIEWING_PATTERN.search(
                    text
                )
            )

            if not match:
                continue

            # Prefer smallest element rather than BODY / giant wrapper.
            matches.append(
                (
                    len(text),
                    element,
                    match,
                )
            )

        except StaleElementReferenceException:
            continue

    if not matches:
        return None

    matches.sort(
        key=lambda item: item[0]
    )

    return matches[0][1]


def get_viewing_state(driver):
    """
    Returns:
        {
            "start": 1,
            "end": 100,
            "total": 283,
            "text": "Viewing 1 to 100 of 283 rows..."
        }
    """

    try:

        element = find_viewing_element(
            driver
        )

        if element is None:
            return None

        text = clean_text(
            element.text
        )

        match = (
            VIEWING_PATTERN.search(
                text
            )
        )

        if not match:
            return None

        return {
            "start": int(
                match.group(1)
            ),
            "end": int(
                match.group(2)
            ),
            "total": int(
                match.group(3)
            ),
            "text": text,
        }

    except (
        StaleElementReferenceException,
        UnexpectedAlertPresentException,
    ):

        dismiss_alert_if_present(
            driver
        )

        return None


def wait_for_viewing_state(
    driver,
    timeout=20
):

    end_time = (
        time.time() + timeout
    )

    while time.time() < end_time:

        state = get_viewing_state(
            driver
        )

        if state is not None:

            return state

        time.sleep(0.25)

    return None


# ============================================================
# FIND ROWS-PER-PAGE INPUT
# ============================================================

def find_rows_input(driver):

    inputs = driver.find_elements(
        By.TAG_NAME,
        "input"
    )

    candidates = []

    for inp in inputs:

        try:

            if not inp.is_displayed():
                continue

            placeholder = clean_text(
                inp.get_attribute(
                    "placeholder"
                )
            )

            # Do not grab the top part-number search field.
            if placeholder:
                continue

            input_type = clean_text(
                inp.get_attribute(
                    "type"
                )
            ).lower()

            if input_type not in (
                "",
                "number",
                "text",
            ):
                continue

            value = clean_text(
                inp.get_attribute(
                    "value"
                )
            )

            if not re.fullmatch(
                r"\d+",
                value
            ):
                continue

            rect = driver.execute_script(
                """
                const r =
                    arguments[0]
                    .getBoundingClientRect();

                return {
                    x:
                        (r.left + r.right) / 2,
                    y:
                        (r.top + r.bottom) / 2
                };
                """,
                inp
            )

            labels = driver.find_elements(
                By.XPATH,
                "//*[normalize-space("
                "translate(text(),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                "'abcdefghijklmnopqrstuvwxyz'))"
                "='rows']"
            )

            closest = None

            for label in labels:

                try:

                    if not label.is_displayed():
                        continue

                    lrect = (
                        driver.execute_script(
                            """
                            const r =
                                arguments[0]
                                .getBoundingClientRect();

                            return {
                                x:
                                    (r.left+r.right)/2,
                                y:
                                    (r.top+r.bottom)/2
                            };
                            """,
                            label
                        )
                    )

                    distance = (
                        (
                            lrect["x"]
                            - rect["x"]
                        ) ** 2
                        +
                        (
                            lrect["y"]
                            - rect["y"]
                        ) ** 2
                    ) ** 0.5

                    if (
                        closest is None
                        or distance < closest
                    ):

                        closest = distance

                except Exception:
                    continue

            if (
                closest is not None
                and closest < 200
            ):

                candidates.append(
                    (
                        closest,
                        inp,
                    )
                )

        except Exception:
            continue

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x[0]
    )

    return candidates[0][1]


# ============================================================
# SET 100 ROWS PER PAGE
# ============================================================

def set_rows_per_page(driver):

    print(
        "\nSetting rows per page "
        "to 100..."
    )

    rows_input = find_rows_input(
        driver
    )

    if rows_input is None:
        raise RuntimeError(
            "Could not locate bottom "
            "rows-per-page input."
        )

    current = clean_text(
        rows_input.get_attribute(
            "value"
        )
    )

    print(
        f"Current rows/page: {current}"
    )

    # ========================================================
    # CHANGE 10 -> 100
    # ========================================================

    if current != "100":

        try:

            driver.execute_script(
                """
                const input = arguments[0];

                const setter =
                    Object.getOwnPropertyDescriptor(
                        HTMLInputElement.prototype,
                        'value'
                    ).set;

                setter.call(input, '100');

                input.dispatchEvent(
                    new Event(
                        'input',
                        {bubbles: true}
                    )
                );

                input.dispatchEvent(
                    new Event(
                        'change',
                        {bubbles: true}
                    )
                );

                input.dispatchEvent(
                    new Event(
                        'blur',
                        {bubbles: true}
                    )
                );

                input.blur();
                """,
                rows_input
            )

        except UnexpectedAlertPresentException:

            dismiss_alert_if_present(
                driver
            )

            raise

    # ========================================================
    # WAIT FOR ACTUAL TABLE TO REBUILD
    # ========================================================

    print(
        "Waiting for 100-row table "
        "to finish loading..."
    )

    # TVS/Zener are much larger than ESD.
    # Give the backend/browser enough time.
    end_time = time.time() + 120

    last_state = None
    stable_count = 0
    last_report_time = 0

    while time.time() < end_time:

        dismiss_alert_if_present(
            driver
        )

        try:

            state = get_viewing_state(
                driver
            )

            # ------------------------------------------------
            # Print progress every ~5 seconds
            # ------------------------------------------------

            if (
                time.time()
                - last_report_time
                >= 5
            ):

                if state:

                    print(
                        "  Current site state: "
                        f"{state['text']}"
                    )

                else:

                    print(
                        "  Still waiting for "
                        "Viewing range..."
                    )

                last_report_time = (
                    time.time()
                )

            if state is None:

                time.sleep(0.25)
                continue

            # For a category containing fewer than 100:
            expected_end = min(
                100,
                state["total"]
            )

            # Success means the SITE itself says
            # we're showing rows 1..100.
            if (
                state["start"] == 1
                and state["end"]
                == expected_end
            ):

                signature = (
                    state["start"],
                    state["end"],
                    state["total"],
                )

                if signature == last_state:
                    stable_count += 1
                else:
                    last_state = signature
                    stable_count = 1

                if stable_count >= 2:

                    print(
                        "Rows/page ready: "
                        f"{state['text']}"
                    )

                    return state

            else:

                stable_count = 0

        except (
            StaleElementReferenceException,
            UnexpectedAlertPresentException,
        ):

            dismiss_alert_if_present(
                driver
            )

        time.sleep(0.25)

    # ========================================================
    # FAILURE DEBUG
    # ========================================================

    try:

        final_input = find_rows_input(
            driver
        )

        final_value = (
            clean_text(
                final_input.get_attribute(
                    "value"
                )
            )
            if final_input is not None
            else "NOT FOUND"
        )

    except Exception:

        final_value = "UNKNOWN"

    state = get_viewing_state(
        driver
    )

    print(
        f"Final rows-input value: "
        f"{final_value}"
    )

    if state:

        print(
            f"Final Viewing state: "
            f"{state['text']}"
        )

    raise RuntimeError(
        "Timed out waiting for "
        "100-row table."
    )
# ============================================================
# FILTER/JUNK ROW DETECTION
# ============================================================

def row_is_filter_row(
    row,
    cells,
    headers
):

    # Select dropdowns in the row = filter row.
    try:

        if row.find_elements(
            By.TAG_NAME,
            "select"
        ):
            return True

    except StaleElementReferenceException:
        return True

    pn_index = (
        get_part_number_index(
            headers
        )
    )

    if (
        pn_index is None
        or pn_index >= len(cells)
    ):

        return True

    pn_cell = cells[
        pn_index
    ]

    # Text/search/number input in PN cell = filter row.
    try:

        for inp in pn_cell.find_elements(
            By.TAG_NAME,
            "input"
        ):

            input_type = clean_text(
                inp.get_attribute(
                    "type"
                )
            ).lower()

            if input_type in (
                "",
                "text",
                "search",
                "number",
            ):

                return True

    except StaleElementReferenceException:
        return True

    return False


# ============================================================
# EXTRACT CURRENT TABLE
# ============================================================

def extract_current_page(
    driver,
    headers
):

    table = find_data_table(
        driver,
        verbose=False
    )

    pn_index = (
        get_part_number_index(
            headers
        )
    )

    records = []

    try:

        rows = table.find_elements(
            By.CSS_SELECTOR,
            "tbody tr"
        )

        if not rows:

            rows = table.find_elements(
                By.CSS_SELECTOR,
                "tr"
            )

    except StaleElementReferenceException:

        return []

    for row in rows:

        try:

            cells = row.find_elements(
                By.XPATH,
                "./td"
            )

            if not cells:
                continue

            if row_is_filter_row(
                row,
                cells,
                headers
            ):
                continue

            values = []

            for index, cell in enumerate(
                cells
            ):

                value = text_content(
                    cell
                )

                # --------------------------------------------
                # PDF LINK
                # --------------------------------------------

                links = cell.find_elements(
                    By.TAG_NAME,
                    "a"
                )

                pdf = ""

                for link in links:

                    href = clean_text(
                        link.get_attribute(
                            "href"
                        )
                    )

                    if not href:
                        continue

                    lower = href.lower()

                    if (
                        ".pdf" in lower
                        or "download" in lower
                    ):

                        pdf = urljoin(
                            BASE_URL,
                            href
                        )

                        break

                if (
                    pdf
                    and not value
                ):

                    value = pdf

                values.append(
                    value
                )

            # --------------------------------------------
            # WIDTH
            # --------------------------------------------

            if len(values) < len(headers):

                values.extend(
                    [""] * (
                        len(headers)
                        - len(values)
                    )
                )

            elif len(values) > len(headers):

                values = values[
                    :len(headers)
                ]

            # --------------------------------------------
            # VALIDATE PART NUMBER
            # --------------------------------------------

            if (
                pn_index is None
                or pn_index >= len(values)
            ):
                continue

            pn = clean_text(
                values[pn_index]
            )

            if not pn:
                continue

            pn_lower = pn.lower()

            if pn_lower in (
                "nichtek p/n",
                "part number",
                "p/n",
            ):
                continue

            # THIS REMOVES THE HUGE CONCATENATED FILTER ROW.
            #
            # Real NichTek PNs are short. A 2000-character
            # "part number" is actually the dropdown's entire
            # option list flattened by textContent.
            if len(pn) > 100:
                continue

            # No whitespace-heavy paragraph-like values.
            if pn.count(" ") > 3:
                continue

            records.append(
                values
            )

        except StaleElementReferenceException:
            continue

    return records


# ============================================================
# WAIT FOR PAGE TABLE TO MATCH VIEWING STATE
# ============================================================

def wait_for_page_records(
    driver,
    headers,
    state,
    timeout=30
):

    expected_count = (
        state["end"]
        - state["start"]
        + 1
    )

    end_time = (
        time.time() + timeout
    )

    best = []

    while time.time() < end_time:

        try:

            records = extract_current_page(
                driver,
                headers
            )

        except StaleElementReferenceException:

            time.sleep(0.2)
            continue

        if len(records) > len(best):

            best = records

        # Once we have exactly the number the site says
        # should be visible, we're done.
        if len(records) == expected_count:

            # Small delay + one verification pass
            time.sleep(0.3)

            try:

                verify = extract_current_page(
                    driver,
                    headers
                )

                if len(verify) == expected_count:

                    return verify

            except StaleElementReferenceException:
                pass

        time.sleep(0.25)

    if len(best) != expected_count:

        print(
            f"WARNING: site says "
            f"{expected_count} rows should be "
            f"visible, scraper found {len(best)}."
        )

    return best
# ============================================================
# FIND REAL PAGINATION BUTTONS
# ============================================================

def get_pagination_candidates(
    driver
):
    """
    Find clickable controls physically associated with
    the 'Viewing X to Y...' status.

    This deliberately does NOT use the rows input as an
    anchor anymore.
    """

    viewing = find_viewing_element(
        driver
    )

    if viewing is None:
        return []

    try:

        view_rect = (
            driver.execute_script(
                """
                const r =
                    arguments[0]
                    .getBoundingClientRect();

                return {
                    left: r.left,
                    right: r.right,
                    top: r.top,
                    bottom: r.bottom,
                    x: (r.left+r.right)/2,
                    y: (r.top+r.bottom)/2
                };
                """,
                viewing
            )
        )

    except StaleElementReferenceException:

        return []

    # Go up a few parents and collect actual clickable descendants.
    clickable = driver.execute_script(
        """
        const viewing = arguments[0];

        let root = viewing;

        // Walk upward until we get a reasonably-sized
        // pagination/footer container.
        for (let i = 0; i < 5 && root.parentElement; i++) {

            const parent =
                root.parentElement;

            const rect =
                parent.getBoundingClientRect();

            if (
                rect.height < 250
                && rect.width > 200
            ) {
                root = parent;
            } else {
                break;
            }
        }

        const els = Array.from(
            root.querySelectorAll(
                'button, a, [role="button"], [onclick]'
            )
        );

        return els;
        """,
        viewing
    )

    candidates = []

    seen = set()

    for element in clickable:

        try:

            if element.id in seen:
                continue

            seen.add(
                element.id
            )

            if not element.is_displayed():
                continue

            rect = driver.execute_script(
                """
                const r =
                    arguments[0]
                    .getBoundingClientRect();

                return {
                    x: (r.left+r.right)/2,
                    y: (r.top+r.bottom)/2,
                    width: r.width,
                    height: r.height
                };
                """,
                element
            )

            # Pagination controls should be in roughly
            # the same horizontal band as Viewing text.
            vertical_distance = abs(
                rect["y"]
                - view_rect["y"]
            )

            if vertical_distance > 120:
                continue

            text = clean_text(
                element.text
            )

            title = clean_text(
                element.get_attribute(
                    "title"
                )
            )

            aria = clean_text(
                element.get_attribute(
                    "aria-label"
                )
            )

            cls = clean_text(
                element.get_attribute(
                    "class"
                )
            )

            disabled = (
                element.get_attribute(
                    "disabled"
                )
            )

            aria_disabled = clean_text(
                element.get_attribute(
                    "aria-disabled"
                )
            ).lower()

            combined = (
                f"{text} "
                f"{title} "
                f"{aria} "
                f"{cls}"
            ).lower()

            if (
                disabled is not None
                or aria_disabled == "true"
                or "disabled" in combined
            ):
                continue

            # Give preference to anything that explicitly
            # looks like next/right/forward.
            score = 0

            if "next" in combined:
                score += 100

            if "right" in combined:
                score += 50

            if "forward" in combined:
                score += 50

            if "chevron-right" in combined:
                score += 100

            if text in (
                ">",
                "›",
                "»",
                "→",
            ):
                score += 100

            # Prefer controls to the right side of Viewing.
            if rect["x"] > view_rect["x"]:
                score += 20

            candidates.append(
                (
                    score,
                    rect["x"],
                    element,
                    combined,
                )
            )

        except StaleElementReferenceException:
            continue

    # Best likely "next" control first.
    candidates.sort(
        key=lambda item: (
            item[0],
            item[1],
        ),
        reverse=True
    )

    return candidates


# ============================================================
# ADVANCE USING VIEWING STATE
# ============================================================

def advance_to_next_page(
    driver,
    old_state,
    headers
):
    """
    Critical rule:

    A candidate button is ONLY considered the true Next button
    when clicking it changes:

        Viewing old_start ... -> Viewing NEW_START ...

    with NEW_START > old_start.

    So we never trust the arrow itself.
    """

    old_start = (
        old_state["start"]
    )

    old_end = (
        old_state["end"]
    )

    total = (
        old_state["total"]
    )

    # Already final page.
    if old_end >= total:
        return None

    candidates = (
        get_pagination_candidates(
            driver
        )
    )

    print(
        f"  Pagination candidates: "
        f"{len(candidates)}"
    )

    if not candidates:

        raise RuntimeError(
            "Could not find pagination "
            "controls near Viewing status."
        )

    for index, (
        score,
        x,
        element,
        description,
    ) in enumerate(
        candidates,
        start=1
    ):

        try:

            # Reacquire state before trying.
            current_state = (
                get_viewing_state(
                    driver
                )
            )

            if (
                current_state is None
                or current_state["start"]
                != old_start
            ):
                # Something already changed.
                if (
                    current_state
                    and current_state["start"]
                    > old_start
                ):
                    return current_state

                continue

            print(
                f"  Trying pagination "
                f"candidate {index}..."
            )

            driver.execute_script(
                """
                arguments[0].scrollIntoView({
                    block:'center',
                    inline:'center'
                });
                """,
                element
            )

            time.sleep(0.1)

            driver.execute_script(
                "arguments[0].click();",
                element
            )

        except StaleElementReferenceException:
            continue

        except UnexpectedAlertPresentException:

            dismiss_alert_if_present(
                driver
            )

            continue

        # --------------------------------------------
        # VERIFY USING VIEWING STATUS
        # --------------------------------------------

        verify_end = (
            time.time() + 8
        )

        while time.time() < verify_end:

            state = get_viewing_state(
                driver
            )

            if state is None:

                time.sleep(0.2)
                continue

            # SUCCESS.
            if (
                state["start"]
                > old_start
            ):

                print(
                    f"  Advanced: "
                    f"{old_state['start']}"
                    f"-{old_state['end']} "
                    f"-> "
                    f"{state['start']}"
                    f"-{state['end']}"
                )

                return state

            time.sleep(0.2)

        # Candidate did not advance.
        #
        # If it changed some irrelevant setting,
        # we rely on Viewing status and simply move
        # to the next candidate.
        print(
            "    Candidate did not "
            "advance Viewing range."
        )

    raise RuntimeError(
        "Found pagination controls, "
        "but none advanced the Viewing range."
    )


# ============================================================
# SCRAPE ALL PAGES
# ============================================================

def scrape_all_pages(
    driver,
    headers
):

    collected = {}

    pn_index = (
        get_part_number_index(
            headers
        )
    )

    state = wait_for_viewing_state(
        driver
    )

    if state is None:

        raise RuntimeError(
            "Could not read NichTek "
            "'Viewing X to Y...' status."
        )

    page_number = 1

    while True:

        print()
        print(
            f"Reading page "
            f"{page_number}: "
            f"{state['start']} to "
            f"{state['end']} of "
            f"{state['total']}..."
        )

        records = (
            wait_for_page_records(
                driver,
                headers,
                state
            )
        )

        expected = (
            state["end"]
            - state["start"]
            + 1
        )

        print(
            f"  Expected rows: "
            f"{expected}"
        )

        print(
            f"  Parsed product rows: "
            f"{len(records)}"
        )

        before = len(
            collected
        )

        for row in records:

            if (
                pn_index is not None
                and pn_index
                < len(row)
            ):

                pn = clean_text(
                    row[pn_index]
                )

            else:

                pn = ""

            # Preserve every distinct NichTek table row.
            #
            # Some NichTek P/Ns may legitimately occur more than once
            # with different package/spec/datasheet information, so do
            # NOT dedupe solely by part number.
            key = tuple(
                clean_text(value)
                for value in row
            )

            collected[key] = row

        added = (
            len(collected)
            - before
        )

        print(
            f"  New unique parts: "
            f"{added}"
        )

        print(
            f"  Total unique parts: "
            f"{len(collected)}"
        )

        # ----------------------------------------------------
        # FINAL PAGE?
        # ----------------------------------------------------

        if (
            state["end"]
            >= state["total"]
        ):

            print(
                "Reached final page "
                "according to Viewing status."
            )

            break

        # ----------------------------------------------------
        # NEXT
        # ----------------------------------------------------

        state = advance_to_next_page(
            driver,
            state,
            headers
        )

        if state is None:
            break

        page_number += 1

        if page_number > 1000:

            raise RuntimeError(
                "Pagination safety limit "
                "reached."
            )

    print()
    print(
        f"Viewing status reported "
        f"{state['total']} total products."
    )

    print(
        f"Unique products collected: "
        f"{len(collected)}"
    )

    if (
        len(collected)
        != state["total"]
    ):

        print(
            "WARNING: collected count "
            "does not match site total."
        )

    return list(
        collected.values()
    )


# ============================================================
# NICHTEK CHANNEL COUNT
# ============================================================

def nichtek_channel_count(part_number, package):
    pn = str(part_number or "").upper().strip()
    pkg = str(package or "").upper().strip()

    # ========================================================
    # VERIFIED MULTI-CHANNEL ARRAYS
    # ========================================================

    # D2550 family / DFN2510 array
    # Datasheet shows four protected I/O lines
    if pn.startswith("D2550"):
        return 4

    # D2519 family is also a DFN2510 multi-line array
    if pn.startswith("D2519"):
        return 4

    # ========================================================
    # VERIFIED 1-CHANNEL DISCRETE TVS / ESD
    # ========================================================

    # Tiny 2-terminal DFN discretes
    if pkg in (
        "DFN0603",
        "DFN1006-2L",
        "DFN1006-2",
    ):
        return 1

    # SOD discrete TVS packages
    if pkg.startswith((
        "SOD-323",
        "SOD-523",
        "SOD-882",
        "SOD-923",
    )):
        return 1

    # Standard axial / SMA / SMB / SMC discrete TVS
    if pkg.startswith((
        "DO-15",
        "DO-41",
        "DO-214AA",
        "DO-214AB",
        "DO-214AC",
        "SMA",
        "SMB",
        "SMC",
    )):
        return 1

    # Known tiny discrete families
    if pn.startswith((
        "D1215",
        "D1219",
        "D1250",
        "D2315",
    )):
        return 1

    # ========================================================
    # MULTI-PIN PACKAGES: DON'T INFER BLINDLY
    # ========================================================

    return None
# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    # Remove blank columns.
    blank_columns = []

    for column in df.columns:

        values = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        if values.eq("").all():

            blank_columns.append(
                column
            )

    if blank_columns:

        df = df.drop(
            columns=blank_columns
        )

    df = (
        df
        .drop_duplicates()
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # REMOVE ANY REMAINING FAKE PN ROWS
    # --------------------------------------------------------

    pn_column = None

    for column in df.columns:

        h = (
            clean_text(column)
            .lower()
            .replace(" ", "")
        )

        if any(
            marker in h
            for marker in [
                "nichtekp/n",
                "partnumber",
                "partno",
            ]
        ):

            pn_column = column
            break

    if pn_column:

        pn = (
            df[pn_column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        valid = (
            pn.ne("")
            &
            pn.str.len().le(100)
            &
            ~pn.str.lower().isin(
                [
                    "nichtek p/n",
                    "part number",
                    "p/n",
                ]
            )
        )

        df = (
            df[valid]
            .copy()
        )

        df = (
            df
            .drop_duplicates()
            .reset_index(drop=True)
        )

    # ========================================================
    # ADD CHANNEL COUNT
    # ========================================================

    pn_column = None
    package_column = None

    for column in df.columns:
        normalized = (
            clean_text(column)
            .lower()
            .replace(" ", "")
        )

        if any(
            marker in normalized
            for marker in [
                "nichtekp/n",
                "partnumber",
                "partno",
            ]
        ):
            pn_column = column

        if normalized == "package":
            package_column = column

    if pn_column is not None and package_column is not None:

        df["Channels"] = df.apply(
            lambda row: nichtek_channel_count(
                row[pn_column],
                row[package_column]
            ),
            axis=1
        )

    return df
# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(
    df,
    filename,
    sheet_name
):

    with pd.ExcelWriter(
        filename,
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

        for column_cells in (
            ws.columns
        ):

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
                    len(value)
                )

            ws.column_dimensions[
                letter
            ].width = min(
                max(
                    max_length + 2,
                    12
                ),
                60
            )


# ============================================================
# SCRAPE ONE CATEGORY
# ============================================================

def scrape_product(
    driver,
    product
):

    print()
    print(
        "=" * 70
    )

    print(
        f"SCRAPING NICHTEK "
        f"{product['name']}"
    )

    print(
        "=" * 70
    )

    driver.get(
        BASE_URL
    )

    wait_for_page(
        driver
    )

    choose_product_category(
        driver,
        product
    )

    # --------------------------------------------------------
    # SET 100 PER PAGE
    # --------------------------------------------------------

    initial_state = (
        set_rows_per_page(
            driver
        )
    )

    print(
        f"Site reports: "
        f"{initial_state['text']}"
    )

    # --------------------------------------------------------
    # TABLE + HEADERS
    # --------------------------------------------------------

    table = find_data_table(
        driver,
        verbose=True
    )

    headers = get_headers(
        table
    )

    if not headers:

        raise RuntimeError(
            "No NichTek table "
            "headers detected."
        )

    # --------------------------------------------------------
    # ALL PAGES
    # --------------------------------------------------------

    records = scrape_all_pages(
        driver,
        headers
    )

    if not records:

        raise RuntimeError(
            f"No {product['name']} "
            f"products extracted."
        )

    # --------------------------------------------------------
    # DATAFRAME
    # --------------------------------------------------------

    df = pd.DataFrame(
        records,
        columns=headers
    )

    df = clean_dataframe(
        df
    )

    print()
    print(
        f"NICHTEK "
        f"{product['name']}: "
        f"{len(df)} unique parts"
    )

    print(
        f"Columns "
        f"({len(df.columns)}):"
    )

    for column in df.columns:

        print(
            f"  {column}"
        )

    print(
        "\nFirst actual part:"
    )

    for column in df.columns:

        print(
            f"  {column}: "
            f"{df.iloc[0][column]}"
        )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    save_excel(
        df,
        product["output"],
        product["sheet"]
    )

    print()
    print(
        f"Saved: "
        f"{product['output']}"
    )

    return df


# ============================================================
# MAIN
# ============================================================

def main():

    driver = make_driver()

    results = {}

    try:

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
                    f"ERROR SCRAPING "
                    f"{product['name']}:"
                )

                print(
                    exc
                )

                dismiss_alert_if_present(
                    driver
                )

                results[
                    product["name"]
                ] = "FAILED"

    finally:

        driver.quit()

    print()
    print(
        "=" * 70
    )

    print(
        "NICHTEK SCRAPE COMPLETE"
    )

    print(
        "=" * 70
    )

    for category, count in (
        results.items()
    ):

        print(
            f"{category}: {count}"
        )

    print()
    print(
        "Output files:"
    )

    print(
        "  nichtek_esd_specs.xlsx"
    )

    print(
        "  nichtek_tvs_specs.xlsx"
    )

    print(
        "  nichtek_zener_specs.xlsx"
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

    Nichtek reuses names across headers - "VZ(V)" sits beside "VZ@IZT(mA)",
    "VRWM(V)" beside "VBR(V)" - so a plain substring search returns whichever
    column happens to sit first. Anchoring at the start of the header keeps
    each parameter on its own column.
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
    The TVS sheet carries two peak pulse power columns that differ only by
    their waveform ("PPP(W)@8/20us" vs "PPP(W)@10/1000us").
    """
    for tokens in token_sets:
        for col_header in row.index:
            header = _norm_header(col_header)
            if all(_norm_header(token) in header for token in tokens):
                return col_header
    return None


def _nichtek_text(value):
    """Cell text, with pandas blanks reduced to an empty string."""
    if value is None:
        return ""
    if not isinstance(value, str) and pd.isna(value):
        return ""
    return clean_text(value)


def _numeric_or_none(value):
    """
    First number in a cell, or None if the cell is blank or non-numeric.

    Nichtek writes an unstated parameter as "-", and qualifies some clamping
    figures in place ("12 /Max."), so the number has to be pulled out of the
    cell rather than handed straight to float().
    """
    text = _nichtek_text(value)
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


def _nichtek_grade(part_row):
    """All three sheets carry an Automotive column holding Yes or No."""
    auto = _nichtek_text(_value_by_prefix(part_row, "Automotive")).upper()
    return "Automotive" if auto.startswith("Y") else "Commercial"


def _nichtek_package(part_row):
    return _nichtek_text(_value_by_prefix(part_row, "Package")) or "-"


def _nichtek_direction(part_row):
    """The Uni./Bi. column states the directionality as "Uni-" or "Bi-"."""
    config = _nichtek_text(_value_by_prefix(part_row, "Uni./Bi.", "Uni")).lower()
    config = re.sub(r'[^a-z]', '', config)
    if config.startswith("uni"):
        return "Unidirectional"
    if config.startswith("bi"):
        return "Bidirectional"
    return "-"


def _parse_nichtek_protection_row(part_row, found_df_name, part_number):
    """
    Parses a row from a Nichtek ESD or TVS datasheet. The two sheets share a
    header set apart from the peak pulse power waveform, so one parser covers
    both.

    Columns: Data Sheet, Nichtek P/N, Package, VRWM(V), VBR(V),
    Clamping Volt.(VCL), PPP(W)@8/20us, PPP(W)@10/1000us (TVS only),
    Uni./Bi., Automotive.

    The sheets state no capacitance, no surge current and no ESD contact
    rating, so those stay at their defaults.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _nichtek_grade(part_row),
        "Direction": _nichtek_direction(part_row),
        "Channels": "-",
        "Package": _nichtek_package(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    # ========================================================
    # CHANNEL COUNT
    # ========================================================

    # First try the Channels column from the scraped Excel.
    channel_val = _numeric_or_none(
        _value_by_prefix(
            part_row,
            "Channels"
        )
    )

    # If the Excel does not have it, infer from NichTek
    # part number / package.
    if channel_val is None:
        channel_val = nichtek_channel_count(
            part_number,
            _nichtek_package(part_row)
        )

    if channel_val is not None:
        specs_result["Channels"] = str(
            int(channel_val)
        )
    vrwm_val = _numeric_or_none(_value_by_prefix(part_row, "VRWM"))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    vcl_val = _numeric_or_none(_value_by_prefix(part_row, "Clamping"))
    if vcl_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vcl_val:g} V"

    # The ESD sheet rates peak pulse power at 8/20us; the TVS sheet carries
    # both columns and leaves 8/20 blank on the parts rated at 10/1000.
    ppp_val = _numeric_or_none(_value_by_tokens(part_row, ("PPP", "8/20")))
    if ppp_val is None:
        ppp_val = _numeric_or_none(_value_by_tokens(part_row, ("PPP", "10/1000"), ("PPP",)))
    if ppp_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp_val:g} W"

    return specs_result


def _parse_nichtek_esd_row(part_row, found_df_name, part_number):
    """Parses a row from nichtek_esd_specs."""
    return _parse_nichtek_protection_row(part_row, found_df_name, part_number)


def _parse_nichtek_tvs_row(part_row, found_df_name, part_number):
    """Parses a row from nichtek_tvs_specs."""
    return _parse_nichtek_protection_row(part_row, found_df_name, part_number)


def _parse_nichtek_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row from nichtek_zener_specs.

    Columns: Data Sheet, Nichtek P/N, Package, VZ(V), VZ@IZT(mA),
    Tolerance(%), PD(mW), Automotive.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _nichtek_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",  # not present in the Zener file headers
        "Package": _nichtek_package(part_row),
    }

    
    # "VZ(" keeps this off "VZ@IZT(mA)", which is the test current.
    vz_val = _numeric_or_none(_value_by_prefix(part_row, "VZ("))
    if vz_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_val:g} V"

    tol_val = _numeric_or_none(_value_by_prefix(part_row, "Tolerance"))
    if tol_val is not None:
        specs_result["Tolerance"] = f"±{tol_val:g}%"

    # PD is stated in mW.
    pd_val = _numeric_or_none(_value_by_prefix(part_row, "PD("))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val / 1000.0:g} W"

    return specs_result


def fetch_nichtek_specs_from_excel(part_number, nichtek_dfs):
    """
    Searches across Nichtek DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS vs. ESD).
    """
    if part_number is None or not nichtek_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["Zener", "ESD", "TVS"]

    all_df_keys = list(nichtek_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["p/n", "part number", "mfr part"]

    for df_name in ordered_search_keys:
        df = nichtek_dfs[df_name]
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
        print(f"Part '{part_number}' not found in any Nichtek database.")
        return None

    print(f"Found specs for Part '{part_number}' in Nichtek database '{found_df_name}'.")

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_nichtek_zener_row(part_row, found_df_name, part_number)
    elif "tvs" in name_lower:
        return _parse_nichtek_tvs_row(part_row, found_df_name, part_number)
    else:
        return _parse_nichtek_esd_row(part_row, found_df_name, part_number)
    
if __name__ == "__main__":
    main()