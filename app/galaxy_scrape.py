# galaxy_scrape.py

import re
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


# ============================================================
# GALAXY PAGES / OUTPUTS
# ============================================================

PAGES = [
    {
        "kind": "esd",
        "name": "Galaxy ESD",
        "url": "https://en.gmesemi.com/products-en/protection-diodes/esd-protection-diodes/",
        "output": "galaxy_esd_specs.xlsx",
    },
    {
        "kind": "tvs",
        "name": "Galaxy TVS",
        "url": "https://en.gmesemi.com/products-en/protection-diodes/tvs/",
        "output": "galaxy_tvs_specs.xlsx",
    },
    {
        "kind": "zener",
        "name": "Galaxy Zener",
        "url": "https://en.gmesemi.com/products-en/protection-diodes/zener-diodes/",
        "output": "galaxy_zener_specs.xlsx",
    },
]


TVS_ESD_HEADERS = [
    "Id",
    "Product",
    "Package",
    "Compliance",
    "DataSheet",
    "Marketing Status",
    "Configuration",
    "Condition",
    "P_PK(W)",
    "V_RWM(V) max.",
    "V_BR(V) min.",
    "V_BR(V) max.",
    "V_C(V) max.",
    "I_PP(A)",
    "I_R(uA) max.",
    "C (pF) max.",
    "T_J(°C) max.",
    "T_J(°C) min.",
    "ECCN(US)",
]


ZENER_HEADERS = [
    "Id",
    "Product",
    "Package",
    "Compliance",
    "DataSheet",
    "Marketing Status",
    "Configuration",
    "Tolerance",
    "V_Z(V) nom.",
    "I_T(mA)",
    "I_R(uA) max.",
    "P_D(mW) max.",
    "T_J(°C) max.",
    "T_J(°C) min.",
    "Pin",
    "ECCN(US)",
]


# ============================================================
# DRIVER
# ============================================================

def make_driver():
    options = webdriver.ChromeOptions()

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")

    service = Service(
        ChromeDriverManager().install()
    )

    return webdriver.Chrome(
        service=service,
        options=options
    )


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    return " ".join(
        str(value)
        .replace("\n", " ")
        .split()
    ).strip()


def find_results_dropdown(driver):
    print("Looking for Show results dropdown...")

    selects = driver.find_elements(
        By.TAG_NAME,
        "select"
    )

    print(
        f"Select elements found: {len(selects)}"
    )

    for i, element in enumerate(
        selects,
        start=1
    ):

        try:
            select = Select(element)

            options = [
                clean_text(o.text)
                for o in select.options
            ]

            print(
                f"Dropdown {i}: {options}"
            )

            if any(
                x.lower() == "all"
                for x in options
            ):
                print(
                    f"Found results dropdown: "
                    f"dropdown {i}"
                )

                return element

        except Exception:
            continue

    return None


def select_all_results(driver):
    """
    Select Galaxy's Show -> All and wait
    for the real rows to finish loading.
    """

    print("Selecting Show -> All...")

    dropdown = find_results_dropdown(
        driver
    )

    if dropdown is None:
        raise RuntimeError(
            "Could not find Galaxy "
            "Show results dropdown."
        )

    select = Select(dropdown)

    all_option = None

    for option in select.options:

        if (
            clean_text(option.text).lower()
            == "all"
        ):
            all_option = option
            break

    if all_option is None:
        raise RuntimeError(
            "Galaxy dropdown does not "
            "contain an All option."
        )

    all_value = all_option.get_attribute(
        "value"
    )

    print(
        f"Selecting All "
        f"(value={all_value!r})..."
    )

    if all_value:
        select.select_by_value(
            all_value
        )

    else:
        select.select_by_visible_text(
            all_option.text
        )

    # Galaxy temporarily switches to a
    # one-row loading state.
    time.sleep(1.5)

    print(
        "Waiting for Galaxy to finish "
        "loading All results..."
    )

    start = time.time()

    last_count = -1
    stable = 0

    while time.time() - start < 60:

        counts = driver.execute_script("""
            return Array.from(
                document.querySelectorAll('table')
            ).map(
                table =>
                    table.querySelectorAll(
                        'tbody tr'
                    ).length
            );
        """)

        max_count = (
            max(counts)
            if counts
            else 0
        )

        print(
            f"\rLargest table currently has "
            f"{max_count} rows",
            end="",
            flush=True,
        )

        # Ignore temporary loading row
        if max_count > 1:

            if max_count == last_count:
                stable += 1

            else:
                stable = 0

            last_count = max_count

            # Stable for several seconds
            if stable >= 3:

                print()

                print(
                    f"All results loaded: "
                    f"{max_count} rows"
                )

                return max_count

        time.sleep(1)

    print()

    raise RuntimeError(
        "Galaxy did not finish loading "
        "All results within 60 seconds."
    )


# ============================================================
# RAW TABLE EXTRACTION
# ============================================================

def get_raw_tables(driver):
    """
    Pull every Galaxy candidate table in ONE JavaScript
    round trip.

    Galaxy renders several split/duplicate tables.

    We preserve:
        - text
        - links
        - compliance badges
        - CSS classes
        - attributes

    so the Python parser can repair rows without relying
    completely on raw TD positions.
    """

    return driver.execute_script("""
        const tables = Array.from(
            document.querySelectorAll('table')
        );

        function attrObject(el) {

            const out = {};

            if (!el || !el.attributes) {
                return out;
            }

            for (const a of el.attributes) {
                out[a.name] = a.value;
            }

            return out;
        }

        return tables.map(
            (table, tableIndex) => {

                const rows = Array.from(
                    table.querySelectorAll(
                        'tbody tr'
                    )
                );

                const data = rows
                    .map(row => {

                        return Array.from(
                            row.querySelectorAll(
                                'td'
                            )
                        ).map(cell => {

                            const text = (
                                cell.innerText || ''
                            )
                            .replace(/\\s+/g, ' ')
                            .trim();


                            // -----------------------------
                            // LINKS
                            // -----------------------------

                            const links = Array.from(
                                cell.querySelectorAll('a')
                            )
                            .map(a => a.href || '')
                            .filter(Boolean);


                            // -----------------------------
                            // COMPLIANCE BADGES
                            //
                            // Pb = Pb-free
                            // H  = Halide-free
                            // A  = AEC-qualified badge
                            // -----------------------------

                            const badgeTexts = [];

                            for (
                                const el of
                                cell.querySelectorAll(
                                    'a, span, div'
                                )
                            ) {

                                const t = (
                                    el.innerText || ''
                                )
                                .replace(/\\s+/g, ' ')
                                .trim();

                                if (
                                    (
                                        t === 'Pb' ||
                                        t === 'H' ||
                                        t === 'A'
                                    )
                                    &&
                                    !badgeTexts.includes(t)
                                ) {
                                    badgeTexts.push(t);
                                }
                            }


                            return {
                                text: text,
                                links: links,
                                badges: badgeTexts,
                                cls:
                                    cell.className || '',
                                attrs:
                                    attrObject(cell),
                            };
                        });

                    })
                    .filter(
                        row => row.length > 0
                    );


                return {
                    tableIndex:
                        tableIndex + 1,

                    rowCount:
                        data.length,

                    data:
                        data,

                    cls:
                        table.className || '',

                    attrs:
                        attrObject(table),
                };
            }
        );
    """)


def cell_value(cell):
    """
    Convert a raw Galaxy cell object
    into its useful value.
    """

    badges = (
        cell.get("badges")
        or []
    )

    if badges:

        # Preserve compliance badges:
        # Pb, H, A
        return " ".join(
            dict.fromkeys(badges)
        )

    text = clean_text(
        cell.get(
            "text",
            ""
        )
    )

    links = (
        cell.get("links")
        or []
    )

    # Datasheet icons often contain
    # no text, only an href.
    if not text and links:

        pdf = next(
            (
                u
                for u in links
                if (
                    ".pdf" in u.lower()
                    or
                    "products_pdf"
                    in u.lower()
                )
            ),
            None,
        )

        if pdf:
            return pdf

        return links[0]

    return text


def choose_product_table(
    tables_data
):

    valid = [
        t
        for t in tables_data
        if t.get(
            "rowCount",
            0
        ) > 1
    ]

    if not valid:
        raise RuntimeError(
            "No Galaxy product rows found."
        )

    for table in tables_data:

        print(
            f"  Table "
            f"{table['tableIndex']}: "
            f"{table['rowCount']} rows"
        )

    # Galaxy generally exposes several
    # duplicate copies of the product data.
    #
    # Prefer the one with the most rows,
    # then the most cells per row.

    best = max(
        valid,
        key=lambda t: (
            t.get(
                "rowCount",
                0
            ),
            max(
                (
                    len(r)
                    for r
                    in t.get(
                        "data",
                        []
                    )
                ),
                default=0,
            ),
        ),
    )

    print(
        f"Using table "
        f"{best['tableIndex']} "
        f"with "
        f"{best['rowCount']} rows."
    )

    return best["data"]


# ============================================================
# FIELD RECOGNITION
# ============================================================

def is_pdf(value):

    s = clean_text(
        value
    ).lower()

    return (
        s.startswith("http")
        and
        (
            ".pdf" in s
            or
            "products_pdf" in s
        )
    )


def is_compliance(value):
    """
    Valid Galaxy compliance values.

    Examples:
        Pb
        H
        A
        Pb H
        Pb H A
        Pb A
    """

    words = set(
        clean_text(value)
        .replace(",", " ")
        .split()
    )

    return (
        bool(words)
        and
        words.issubset(
            {
                "Pb",
                "H",
                "A",
            }
        )
    )


def normalize_compliance(value):
    """
    Normalize badge order consistently:
        Pb -> H -> A
    """

    words = (
        clean_text(value)
        .replace(",", " ")
        .split()
    )

    out = []

    if "Pb" in words:
        out.append("Pb")

    if "H" in words:
        out.append("H")

    if "A" in words:
        out.append("A")

    return " ".join(out)


def is_marketing_status(value):

    s = clean_text(
        value
    ).lower()

    return s in {
        "active",
        "discontinued",
        "not for new design",
        "not recommended for new design",
        "active and preferred",
    }


def is_tvs_esd_configuration(
    value
):

    s = clean_text(
        value
    ).lower()

    return s in {
        "single",
        "dual",
        "uni-directional",
        "unidirectional",
        "bi-directional",
        "bidirectional",
    }


def is_zener_configuration(
    value
):

    s = clean_text(
        value
    ).lower()

    return s in {
        "single",
        "dual",
        "dual,common cathode",
        "dual, common cathode",
        "dual,two single",
        "dual, two single",
        "two duals,common anode",
        "two duals, common anode",
    }


def is_condition(value):

    s = (
        clean_text(value)
        .lower()
        .replace(
            "×",
            "x"
        )
        .replace(
            "µ",
            "u"
        )
        .replace(
            "μ",
            "u"
        )
    )

    s = s.replace(
        " ",
        ""
    )

    return bool(
        re.fullmatch(
            r"\d+(?:\.\d+)?[/x]"
            r"\d+(?:\.\d+)?us",
            s,
        )
    )


def is_eccn(value):

    s = clean_text(
        value
    ).upper()

    if not s:
        return False

    return (
        s == "EAR99"
        or
        bool(
            re.fullmatch(
                r"[0-9][A-Z][0-9]{3}"
                r"(?:\.[A-Z0-9]+)?",
                s,
            )
        )
    )


def looks_like_description(
    value,
    kind
):
    """
    Hidden Galaxy description text that must
    never become an electrical spec.
    """

    s = clean_text(
        value
    ).lower()

    if not s:
        return False

    phrases = [
        "surface mount tvs",
        "surface mount zener",
        "zener diode",
        "esd protection",
        "transient voltage suppressor",
        "transient voltage suppressors",
    ]

    if any(
        p in s
        for p in phrases
    ):
        return True

    # Examples:
    # 70V,Surface Mount TVS
    # 5V,100W,Surface Mount TVS

    if (
        kind in {
            "esd",
            "tvs"
        }
        and
        "tvs" in s
        and
        "," in s
    ):
        return True

    if (
        kind == "zener"
        and
        "zener" in s
        and
        "," in s
    ):
        return True

    return False


def clean_configuration(value):

    s = clean_text(
        value
    )

    low = s.lower()

    if low in {
        "unidirectional",
        "uni-directional",
    }:
        return "Uni-directional"

    if low in {
        "bidirectional",
        "bi-directional",
    }:
        return "Bi-directional"

    return s


# ============================================================
# TVS / ESD ROW PARSER
# ============================================================

def parse_tvs_esd_row(
    raw_cells
):
    """
    Parse Galaxy ESD/TVS rows by field meaning,
    rather than blindly trusting raw TD position.

    This prevents missing DataSheet / Configuration
    cells from shifting electrical specs.
    """

    values = [
        cell_value(c)
        for c in raw_cells
    ]

    result = {
        h: ""
        for h
        in TVS_ESD_HEADERS
    }

    # --------------------------------------------------------
    # FIRST THREE FIELDS
    # --------------------------------------------------------

    if len(values) > 0:
        result["Id"] = values[0]

    if len(values) > 1:
        result["Product"] = values[1]

    if len(values) > 2:
        result["Package"] = values[2]

    used = {
        i
        for i
        in range(
            min(
                3,
                len(values)
            )
        )
    }


    # --------------------------------------------------------
    # COMPLIANCE
    #
    # Prefer the actual badge-bearing cell.
    # --------------------------------------------------------

    for i, cell in enumerate(
        raw_cells
    ):

        if i in used:
            continue

        badges = (
            cell.get("badges")
            or []
        )

        if badges:

            result[
                "Compliance"
            ] = normalize_compliance(
                " ".join(badges)
            )

            used.add(i)

            break


    # --------------------------------------------------------
    # RECOGNIZABLE TEXT FIELDS
    # --------------------------------------------------------

    for i, value in enumerate(
        values
    ):

        if i in used:
            continue

        if (
            not result["Compliance"]
            and
            is_compliance(value)
        ):

            result[
                "Compliance"
            ] = normalize_compliance(
                value
            )

            used.add(i)


        elif (
            not result["DataSheet"]
            and
            is_pdf(value)
        ):

            result[
                "DataSheet"
            ] = value

            used.add(i)


        elif (
            not result[
                "Marketing Status"
            ]
            and
            is_marketing_status(
                value
            )
        ):

            result[
                "Marketing Status"
            ] = value

            used.add(i)


        elif (
            not result[
                "Configuration"
            ]
            and
            is_tvs_esd_configuration(
                value
            )
        ):

            result[
                "Configuration"
            ] = clean_configuration(
                value
            )

            used.add(i)


        elif (
            not result[
                "Condition"
            ]
            and
            is_condition(
                value
            )
        ):

            result[
                "Condition"
            ] = value

            used.add(i)


        elif (
            not result[
                "ECCN(US)"
            ]
            and
            is_eccn(
                value
            )
        ):

            result[
                "ECCN(US)"
            ] = value

            used.add(i)


    # --------------------------------------------------------
    # REMAINING ELECTRICAL SPECS
    # --------------------------------------------------------

    remaining = []

    for i, value in enumerate(
        values
    ):

        if i in used:
            continue

        value = clean_text(
            value
        )

        if not value:
            continue

        if looks_like_description(
            value,
            "tvs"
        ):
            continue

        if is_compliance(
            value
        ):
            continue

        if is_pdf(
            value
        ):
            continue

        if is_marketing_status(
            value
        ):
            continue

        if is_tvs_esd_configuration(
            value
        ):
            continue

        if is_condition(
            value
        ):
            continue

        if is_eccn(
            value
        ):
            continue

        remaining.append(
            value
        )


    param_headers = [
        "P_PK(W)",
        "V_RWM(V) max.",
        "V_BR(V) min.",
        "V_BR(V) max.",
        "V_C(V) max.",
        "I_PP(A)",
        "I_R(uA) max.",
        "C (pF) max.",
        "T_J(°C) max.",
        "T_J(°C) min.",
    ]


    for header, value in zip(
        param_headers,
        remaining
    ):

        result[
            header
        ] = value


    return result


# ============================================================
# ZENER ROW PARSER
# ============================================================

def parse_zener_row(
    raw_cells
):
    """
    Parse Galaxy Zener rows.

    Published spec sequence:
        Tolerance
        V_Z(V) nom.
        I_T(mA)
        I_R(uA) max.
        P_D(mW) max.
        T_J max.
        T_J min.
        Pin
        ECCN
    """

    values = [
        cell_value(c)
        for c in raw_cells
    ]

    result = {
        h: ""
        for h
        in ZENER_HEADERS
    }


    # --------------------------------------------------------
    # FIRST THREE FIELDS
    # --------------------------------------------------------

    if len(values) > 0:
        result["Id"] = values[0]

    if len(values) > 1:
        result["Product"] = values[1]

    if len(values) > 2:
        result["Package"] = values[2]


    used = {
        i
        for i
        in range(
            min(
                3,
                len(values)
            )
        )
    }


    # --------------------------------------------------------
    # COMPLIANCE
    # --------------------------------------------------------

    for i, cell in enumerate(
        raw_cells
    ):

        if i in used:
            continue

        badges = (
            cell.get("badges")
            or []
        )

        if badges:

            result[
                "Compliance"
            ] = normalize_compliance(
                " ".join(badges)
            )

            used.add(i)

            break


    # --------------------------------------------------------
    # RECOGNIZABLE FIELDS
    # --------------------------------------------------------

    for i, value in enumerate(
        values
    ):

        if i in used:
            continue


        if (
            not result[
                "Compliance"
            ]
            and
            is_compliance(
                value
            )
        ):

            result[
                "Compliance"
            ] = normalize_compliance(
                value
            )

            used.add(i)


        elif (
            not result[
                "DataSheet"
            ]
            and
            is_pdf(
                value
            )
        ):

            result[
                "DataSheet"
            ] = value

            used.add(i)


        elif (
            not result[
                "Marketing Status"
            ]
            and
            is_marketing_status(
                value
            )
        ):

            result[
                "Marketing Status"
            ] = value

            used.add(i)


        elif (
            not result[
                "Configuration"
            ]
            and
            is_zener_configuration(
                value
            )
        ):

            result[
                "Configuration"
            ] = clean_configuration(
                value
            )

            used.add(i)


        elif (
            not result[
                "ECCN(US)"
            ]
            and
            is_eccn(
                value
            )
        ):

            result[
                "ECCN(US)"
            ] = value

            used.add(i)


    # --------------------------------------------------------
    # REMAINING ZENER SPECS
    # --------------------------------------------------------

    remaining = []

    for i, value in enumerate(
        values
    ):

        if i in used:
            continue

        value = clean_text(
            value
        )

        if not value:
            continue

        if looks_like_description(
            value,
            "zener"
        ):
            continue

        if is_compliance(
            value
        ):
            continue

        if is_pdf(
            value
        ):
            continue

        if is_marketing_status(
            value
        ):
            continue

        if is_zener_configuration(
            value
        ):
            continue

        if is_eccn(
            value
        ):
            continue

        remaining.append(
            value
        )


    param_headers = [
        "Tolerance",
        "V_Z(V) nom.",
        "I_T(mA)",
        "I_R(uA) max.",
        "P_D(mW) max.",
        "T_J(°C) max.",
        "T_J(°C) min.",
        "Pin",
    ]


    for header, value in zip(
        param_headers,
        remaining
    ):

        result[
            header
        ] = value


    return result


# ============================================================
# EXTRACT ONE PAGE
# ============================================================

def extract_table(
    driver,
    kind
):

    print(
        "\nReading Galaxy product table..."
    )

    tables_data = get_raw_tables(
        driver
    )

    records = choose_product_table(
        tables_data
    )

    parsed = []

    for raw_cells in records:

        if kind in {
            "esd",
            "tvs"
        }:

            row = parse_tvs_esd_row(
                raw_cells
            )

        elif kind == "zener":

            row = parse_zener_row(
                raw_cells
            )

        else:

            raise ValueError(
                f"Unknown Galaxy page "
                f"kind: {kind}"
            )


        if clean_text(
            row.get(
                "Product",
                ""
            )
        ):

            parsed.append(
                row
            )


    headers = (
        TVS_ESD_HEADERS
        if kind in {
            "esd",
            "tvs"
        }
        else
        ZENER_HEADERS
    )


    df = pd.DataFrame(
        parsed,
        columns=headers
    )

    print(
        f"Rows extracted: "
        f"{len(df)}"
    )


    # --------------------------------------------------------
    # DIAGNOSTIC:
    # Catch the previous DataSheet/status shift problem.
    # --------------------------------------------------------

    bad_status = df[
        df[
            "DataSheet"
        ]
        .astype(str)
        .str.lower()
        .isin(
            [
                "active",
                "discontinued",
                "not for new design",
            ]
        )
    ]


    if not bad_status.empty:

        print(
            f"WARNING: "
            f"{len(bad_status)} rows "
            f"still appear to have a "
            f"status inside DataSheet. "
            f"Check: "
            +
            ", ".join(
                bad_status[
                    "Product"
                ]
                .astype(str)
                .head(5)
            )
        )


    return df


def derive_config_and_channels(df):
    """
    Galaxy states directionality for only some parts. Where the Configuration
    cell says uni or bi outright, that wins. Otherwise the suffix carries it:
    a part ending in C or CW is 2-channel bidirectional, CA is 1-channel
    bidirectional. Anything else is left blank so it is not scored on.
    """
    if "Configuration" not in df.columns or "Product" not in df.columns:
        return df

    directions = []
    channels = []

    for _, row in df.iterrows():

        stated = clean_text(
            row.get("Configuration", "")
        ).lower().replace("-", "")

        direction = ""

        if stated.startswith("uni"):
            direction = "Uni-directional"
        elif stated.startswith("bi"):
            direction = "Bi-directional"

        product = clean_text(
            row.get("Product", "")
        ).upper()

        if product.endswith("CA"):
            suffix_direction = "Bi-directional"
            channel_count = "1"
        elif product.endswith("CW") or product.endswith("C"):
            suffix_direction = "Bi-directional"
            channel_count = "2"
        else:
            suffix_direction = ""
            channel_count = ""

        directions.append(direction or suffix_direction)
        channels.append(channel_count)

    df["Configuration"] = directions
    df["Channels"] = channels

    return df

# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(
    df,
    output_file,
    sheet_name
):

    print(
        "\nCleaning collected data..."
    )

    # Keep empty strings because blank specs
    # are meaningful.
    #
    # Only remove fully duplicate rows.

    before = len(df)

    df = (
        df
        .drop_duplicates()
        .reset_index(
            drop=True
        )
    )

    after = len(df)


    if before != after:

        print(
            f"Removed "
            f"{before - after} "
            f"duplicate rows."
        )


    print(
        f"Final unique rows: "
        f"{after}"
    )


    # --------------------------------------------------------
    # WRITE EXCEL
    # --------------------------------------------------------

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


        # ----------------------------------------------------
        # COLUMN WIDTHS
        # ----------------------------------------------------

        for column_cells in ws.columns:

            max_length = 0

            column_letter = (
                column_cells[
                    0
                ].column_letter
            )

            for cell in column_cells:

                value = (
                    ""
                    if cell.value
                    is None
                    else
                    str(
                        cell.value
                    )
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
                    10
                ),
                45
            )


    print(
        f"Saved successfully as: "
        f"{output_file}"
    )


# ============================================================
# SCRAPE ONE GALAXY CATEGORY
# ============================================================

# ============================================================
# PAGINATION HELPERS
# ============================================================

def get_visible_page_numbers(driver):
    """
    Return visible numeric pagination buttons.

    Example:
        1 2 3 4 5

    Returns:
        [1, 2, 3, 4, 5]

    If there is no pagination, returns [].
    """

    page_numbers = driver.execute_script("""
        const candidates = Array.from(
            document.querySelectorAll(
                'a, button, li'
            )
        );

        const nums = [];

        for (const el of candidates) {

            const text = (
                el.innerText || ''
            ).trim();

            if (!/^\\d+$/.test(text)) {
                continue;
            }

            const n = parseInt(text, 10);

            // Avoid accidentally capturing random numbers
            // elsewhere on the page.
            const cls = (
                (
                    el.className || ''
                )
                +
                ' '
                +
                (
                    el.parentElement
                    ? el.parentElement.className || ''
                    : ''
                )
            ).toLowerCase();

            const nearPager =
                cls.includes('page') ||
                cls.includes('pagination') ||
                cls.includes('pager');

            if (nearPager) {
                nums.push(n);
            }
        }

        return [...new Set(nums)].sort(
            (a, b) => a - b
        );
    """)

    return page_numbers


def get_current_product_signature(driver):
    """
    Grab a small signature from the current product table.

    Used to verify that clicking Next actually changed
    the displayed page.
    """

    return driver.execute_script("""
        const tables = Array.from(
            document.querySelectorAll('table')
        );

        let best = null;

        for (const table of tables) {

            const rows = Array.from(
                table.querySelectorAll(
                    'tbody tr'
                )
            );

            if (
                !best ||
                rows.length > best.rows.length
            ) {
                best = {
                    table: table,
                    rows: rows
                };
            }
        }

        if (
            !best ||
            best.rows.length === 0
        ) {
            return '';
        }

        const first = (
            best.rows[0].innerText || ''
        )
        .replace(/\\s+/g, ' ')
        .trim();

        const last = (
            best.rows[
                best.rows.length - 1
            ].innerText || ''
        )
        .replace(/\\s+/g, ' ')
        .trim();

        return (
            best.rows.length +
            '|' +
            first +
            '|' +
            last
        );
    """)


def find_next_button(driver):
    """
    Find Galaxy's active Next / > pagination button.

    Returns a Selenium element or None.
    """

    candidates = driver.find_elements(
        By.XPATH,
        """
        //a[
            normalize-space(.)='Next'
            or normalize-space(.)='>'
            or normalize-space(.)='›'
            or normalize-space(.)='»'
        ]
        |
        //button[
            normalize-space(.)='Next'
            or normalize-space(.)='>'
            or normalize-space(.)='›'
            or normalize-space(.)='»'
        ]
        """
    )

    for candidate in candidates:

        try:

            if not candidate.is_displayed():
                continue

            cls = (
                candidate.get_attribute(
                    "class"
                )
                or ""
            ).lower()

            aria_disabled = (
                candidate.get_attribute(
                    "aria-disabled"
                )
                or ""
            ).lower()

            disabled = (
                candidate.get_attribute(
                    "disabled"
                )
            )

            parent_cls = ""

            try:
                parent_cls = (
                    candidate.find_element(
                        By.XPATH,
                        ".."
                    )
                    .get_attribute(
                        "class"
                    )
                    or ""
                ).lower()

            except Exception:
                pass

            if (
                disabled
                or aria_disabled == "true"
                or "disabled" in cls
                or "disabled" in parent_cls
            ):
                continue

            return candidate

        except Exception:
            continue

    return None


def scrape_all_visible_pages(
    driver,
    kind
):
    """
    Extract the current page, then keep clicking Next
    until there are no more pages.

    This is used EVEN AFTER Show -> All, because Galaxy
    may still paginate the result set.
    """

    all_frames = []

    page_number = 1
    seen_signatures = set()

    while True:

        print(
            f"\nParsing displayed page "
            f"{page_number}..."
        )

        signature = (
            get_current_product_signature(
                driver
            )
        )

        if (
            signature
            and
            signature in seen_signatures
        ):
            print(
                "Current page was already "
                "seen. Stopping pagination."
            )
            break

        if signature:
            seen_signatures.add(
                signature
            )


        # ----------------------------------------------------
        # EXTRACT CURRENT PAGE
        # ----------------------------------------------------

        df_page = extract_table(
            driver,
            kind
        )

        print(
            f"Rows parsed on page "
            f"{page_number}: "
            f"{len(df_page)}"
        )

        all_frames.append(
            df_page
        )


        # ----------------------------------------------------
        # CHECK WHETHER ANOTHER PAGE EXISTS
        # ----------------------------------------------------

        next_button = find_next_button(
            driver
        )

        if next_button is None:

            print(
                "No active Next button found. "
                "Reached final page."
            )

            break


        old_signature = signature

        print(
            "Additional page detected. "
            "Clicking Next..."
        )


        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center'
            });
            """,
            next_button
        )

        time.sleep(
            0.3
        )

        driver.execute_script(
            "arguments[0].click();",
            next_button
        )


        # ----------------------------------------------------
        # WAIT FOR PAGE TO CHANGE
        # ----------------------------------------------------

        changed = False

        start = time.time()

        while (
            time.time() - start
            < 20
        ):

            time.sleep(
                0.5
            )

            new_signature = (
                get_current_product_signature(
                    driver
                )
            )

            if (
                new_signature
                and
                new_signature != old_signature
            ):
                changed = True
                break


        if not changed:

            print(
                "Next was clicked, but the "
                "product table did not change. "
                "Assuming final page."
            )

            break


        page_number += 1


    if not all_frames:

        raise RuntimeError(
            "No Galaxy rows were collected."
        )


    # --------------------------------------------------------
    # COMBINE ALL PAGES
    # --------------------------------------------------------

    df = pd.concat(
        all_frames,
        ignore_index=True
    )


    before = len(df)

    df = (
        df
        .drop_duplicates()
        .reset_index(
            drop=True
        )
    )

    after = len(df)


    if before != after:

        print(
            f"Removed "
            f"{before - after} "
            f"duplicate rows after "
            f"combining pages."
        )


    print(
        f"\nTotal rows after "
        f"all displayed pages: "
        f"{after}"
    )


    return df


# ============================================================
# SCRAPE ONE GALAXY CATEGORY
# ============================================================

def scrape_category(
    driver,
    page
):

    print(
        "\n"
        +
        "=" * 70
    )

    print(
        f"Opening "
        f"{page['name']} page..."
    )

    print(
        page["url"]
    )

    print(
        "=" * 70
    )


    driver.get(
        page["url"]
    )


    wait = WebDriverWait(
        driver,
        30
    )

    wait.until(
        EC.presence_of_element_located(
            (
                By.TAG_NAME,
                "table"
            )
        )
    )


    time.sleep(
        2.5
    )


    print(
        "Page loaded."
    )


    # --------------------------------------------------------
    # SHOW -> ALL
    # --------------------------------------------------------

    select_all_results(
        driver
    )


    # --------------------------------------------------------
    # CHECK PAGINATION AFTER "ALL"
    # --------------------------------------------------------

    page_numbers = (
        get_visible_page_numbers(
            driver
        )
    )


    if page_numbers:

        print(
            "Pagination still detected "
            "after selecting All: "
            +
            ", ".join(
                str(x)
                for x
                in page_numbers
            )
        )

    else:

        print(
            "No numeric pagination "
            "detected after selecting All."
        )


    # --------------------------------------------------------
    # SCRAPE CURRENT PAGE + ANY REMAINING PAGES
    # --------------------------------------------------------

    print(
        f"\nReading all "
        f"{page['name']} "
        f"product data..."
    )


    df = scrape_all_visible_pages(
        driver,
        page["kind"]
    )


    print(
        f"\nTotal unique rows collected: "
        f"{len(df)}"
    )


    # --------------------------------------------------------
    # PREVIEW
    # --------------------------------------------------------

    if not df.empty:

        print(
            "\nFirst parsed record:"
        )

        for column in df.columns:

            print(
                f"  "
                f"{column}: "
                f"{df.iloc[0][column]}"
            )

    if page["kind"] in ("esd", "tvs"):
        df = derive_config_and_channels(df)

    save_excel(
        df,
        page["output"],
        page["name"]
    )
    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    save_excel(
        df,
        page["output"],
        page["name"]
    )


    return df
# ============================================================
# MAIN
# ============================================================

def main():

    driver = make_driver()

    try:

        results = {}


        for page in PAGES:

            try:

                results[
                    page["kind"]
                ] = scrape_category(
                    driver,
                    page
                )

            except Exception as exc:

                print(
                    f"\nERROR while scraping "
                    f"{page['name']}: "
                    f"{exc}"
                )

                print(
                    "Continuing to the next "
                    "Galaxy category..."
                )


        # ----------------------------------------------------
        # FINAL SUMMARY
        # ----------------------------------------------------

        print(
            "\n"
            +
            "=" * 70
        )

        print(
            "GALAXY SCRAPE COMPLETE"
        )

        print(
            "=" * 70
        )


        for page in PAGES:

            df = results.get(
                page["kind"]
            )

            if df is not None:

                print(
                    f"{page['output']}: "
                    f"{len(df)} rows"
                )

            else:

                print(
                    f"{page['output']}: "
                    f"FAILED"
                )


    finally:

        print(
            "\nClosing browser..."
        )

        driver.quit()

# ============================================================
# SPEC LOOKUP (used by applesauce, not by the scraper)
# ============================================================

def _norm_header(col_header):
    """Header text with spaces, underscores and case removed, for matching."""
    return re.sub(r'[\s_]+', '', str(col_header)).upper()


def _get_header_by_prefix(row, *prefixes):
    """
    First column header whose normalized text starts with one of the prefixes.

    Galaxy reuses letters across headers - "Compliance" and "C (pF) max." both
    begin with C, "I_PP(A)" and "I_R(uA) max." both begin with I - so a plain
    substring search returns whichever column happens to sit first. Anchoring
    at the start of the header keeps each parameter on its own column.
    """
    for prefix in prefixes:
        target = _norm_header(prefix)
        for col_header in row.index:
            if _norm_header(col_header).startswith(target):
                return col_header
    return None


def _value_by_prefix(row, *prefixes):
    """Cell value for the first header matching one of the prefixes."""
    header = _get_header_by_prefix(row, *prefixes)
    return row.get(header) if header is not None else None


def _galaxy_text(value):
    """Cell text, with pandas blanks reduced to an empty string."""
    if value is None:
        return ""
    if not isinstance(value, str) and pd.isna(value):
        return ""
    return clean_text(value)


def _numeric_or_none(value):
    """First number in a cell, or None if the cell is blank or non-numeric."""
    text = _galaxy_text(value)
    if text in ("", "/", "-", "--", "N/A", "NA"):
        return None
    numeric_match = re.search(r"[\d.]+", text)
    if not numeric_match:
        return None
    try:
        return float(numeric_match.group(0))
    except (ValueError, TypeError):
        return None


def galaxy_config_and_channels(part_row, part_number):
    """
    Galaxy states directionality for only some parts.

    The Configuration column wins where it says uni or bi outright, and the
    channel count then stays unstated. Otherwise the part-number suffix
    carries both: ending in C or CW is a 2-channel bidirectional, CA is a
    1-channel bidirectional. Anything else leaves both at "-" so neither is
    scored on.
    """
    config = _galaxy_text(_value_by_prefix(part_row, "Configuration")).lower()
    config = re.sub(r'[^a-z]', '', config)

    if config.startswith("uni"):
        return "Unidirectional", "-"
    if config.startswith("bi"):
        return "Bidirectional", "-"

    pn = str(part_number).strip().upper()

    if pn.endswith("CA"):
        return "Bidirectional", "1"
    if pn.endswith("CW") or pn.endswith("C"):
        return "Bidirectional", "2"

    return "-", "-"

def _galaxy_grade(part_row):
    """
    The Compliance column carries space-separated marks ("Pb H A"), where A is
    the automotive qualification. It is matched as a whole token so an A
    inside another word cannot promote a commercial part.
    """
    compliance = _galaxy_text(_value_by_prefix(part_row, "Compliance")).upper()
    if not compliance:
        return "Commercial"
    tokens = re.split(r'[\s,/]+', compliance)
    return "Automotive" if "A" in tokens else "Commercial"

def _parse_galaxy_tvs_esd_row(part_row, found_df_name, part_number):
    """
    Parses a row written with TVS_ESD_HEADERS. The ESD and TVS sheets share
    that header set, so one parser covers both.

    Compliance carries only the Pb and H marks, so there is no automotive
    qualification to read, and the sheet has no ESD contact rating.
    """
    direction, channels = galaxy_config_and_channels(part_row, part_number)

    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _galaxy_grade(part_row),
        "Direction": direction,
        "Channels": channels,
        "Package": _galaxy_text(_value_by_prefix(part_row, "Package")) or "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Leakage Current (Ir)": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    ppk_val = _numeric_or_none(_value_by_prefix(part_row, "P_PK"))
    if ppk_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppk_val:g} W"

    vrwm_val = _numeric_or_none(_value_by_prefix(part_row, "V_RWM"))
    if vrwm_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm_val:g} V"

    vc_val = _numeric_or_none(_value_by_prefix(part_row, "V_C("))
    if vc_val is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{vc_val:g} V"

    ipp_val = _numeric_or_none(_value_by_prefix(part_row, "I_PP"))
    if ipp_val is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp_val:g} A (8/20µs)"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "I_R("))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    cap_val = _numeric_or_none(_value_by_prefix(part_row, "C (pF"))
    if cap_val is not None:
        specs_result["Capacitance"] = f"{cap_val:.2f} pF"

    return specs_result


def _parse_galaxy_zener_row(part_row, found_df_name, part_number):
    """
    Parses a row written with ZENER_HEADERS.

    P_D is stated in mW here, and Tolerance is a real column rather than
    something derived from a VZ window.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": _galaxy_grade(part_row),
        "Voltage - Reverse Standoff (Typ)": "-",  # holds Vz for comparison
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",  # not present in the Zener file headers
        "Package": _galaxy_text(_value_by_prefix(part_row, "Package")) or "-",
    }

    vz_val = _numeric_or_none(_value_by_prefix(part_row, "V_Z("))
    if vz_val is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vz_val:g} V"

    tol_val = _numeric_or_none(_value_by_prefix(part_row, "Tolerance"))
    if tol_val is not None:
        specs_result["Tolerance"] = f"±{tol_val:g}%"

    pd_val = _numeric_or_none(_value_by_prefix(part_row, "P_D("))
    if pd_val is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_val / 1000.0:g} W"

    ir_val = _numeric_or_none(_value_by_prefix(part_row, "I_R("))
    if ir_val is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_val:g} uA"

    return specs_result


def fetch_galaxy_specs_from_excel(part_number, galaxy_dfs):
    """
    Searches across Galaxy DataFrames, finds the part, and dispatches to the
    correct parser based on the source file type (Zener vs. TVS/ESD).
    """
    if part_number is None or not galaxy_dfs:
        return None

    part_row = None
    found_df_name = ""

    search_priority = ["Zener", "ESD", "TVS"]

    all_df_keys = list(galaxy_dfs.keys())
    ordered_search_keys = [key for key in search_priority if key in all_df_keys]
    ordered_search_keys.extend([key for key in all_df_keys if key not in ordered_search_keys])

    part_num_col_keywords = ["product", "part number", "mfr part"]

    for df_name in ordered_search_keys:
        df = galaxy_dfs[df_name]
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
        print(f"Part '{part_number}' not found in any Galaxy database.")
        return None

    print(f"Found specs for Part '{part_number}' in Galaxy database '{found_df_name}'.")

    if "zener" in found_df_name.lower():
        return _parse_galaxy_zener_row(part_row, found_df_name, part_number)
    else:
        return _parse_galaxy_tvs_esd_row(part_row, found_df_name, part_number)
    
if __name__ == "__main__":
    main()