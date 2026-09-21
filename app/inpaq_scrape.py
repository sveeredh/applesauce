"""
INPAQ scraper
=============

Outputs:

1. inpaq_auto_esd_specs.xlsx
2. inpaq_esd_specs.xlsx
3. inpaq_auto_tvs_specs.xlsx
4. inpaq_tvs_specs.xlsx

For categories made from multiple INPAQ pages:
- all rows are combined
- all unique headers are preserved
- missing values stay blank
- duplicate part numbers are merged using nonblank values

INPAQ only allows 5 products at a time through its manual download interface,
but the scraper bypasses that limitation by reading the underlying product
table directly.
"""

from __future__ import annotations

import re
import sys
import time
from copy import copy
from io import StringIO
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup


# =====================================================================
# CONFIGURATION
# =====================================================================

SCRAPE_GROUPS = [
    {
        "name": "INPAQ Automotive ESD",
        "sheet_name": "INPAQ Auto ESD",
        "output": Path("inpaq_auto_esd_specs.xlsx"),
        "urls": [
            "https://www.inpaqgp.com/2023-product.php?"
            "k4=HBNUaGWPEJrMbKk4yGKomg%3D%3D"
        ],
    },

    {
        "name": "INPAQ ESD",
        "sheet_name": "INPAQ ESD",
        "output": Path("inpaq_esd_specs.xlsx"),
        "urls": [
            "https://www.inpaqgp.com/2023-product.php?"
            "k4=POGIATj5TEZ8zmHNy5SoyQ%3D%3D",

            "https://www.inpaqgp.com/2023-product.php?"
            "k4=B03dGrQ9Nu4rVaA3o1wA2A%3D%3D",
        ],
    },

    {
        "name": "INPAQ Automotive TVS",
        "sheet_name": "INPAQ Auto TVS",
        "output": Path("inpaq_auto_tvs_specs.xlsx"),

        # INPAQ's English 2023-product iframe for this category is broken.
        # Use the working Power TVS AM product-family page instead.
        "urls": [
            "https://www.inpaqgp.com/jp/rw_0f65fb1a9fba06c0f54b25221669cf8b"
        ],

        "special_auto_tvs": True,
    },

    {
        "name": "INPAQ TVS",
        "sheet_name": "INPAQ TVS",
        "output": Path("inpaq_tvs_specs.xlsx"),
        "urls": [
            "https://www.inpaqgp.com/2023-product.php?"
            "k4=1AtCCFMba9iiZ0mpLfMKXQ%3D%3D",

            "https://www.inpaqgp.com/2023-product.php?"
            "k4=X310iEXEFUkRWbqU8Y8OwA%3D%3D",
        ],
    },
]


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/152.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


# =====================================================================
# GENERAL HELPERS
# =====================================================================

def clean_text(value: object) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)

    return session


def get(
    session: requests.Session,
    url: str,
    *,
    timeout: int = 30,
) -> requests.Response:

    last_error: Exception | None = None

    for attempt in range(1, 4):

        try:

            response = session.get(
                url,
                timeout=timeout,
            )

            response.raise_for_status()

            return response

        except requests.RequestException as exc:

            last_error = exc

            if attempt < 3:

                print(
                    f"  Request failed. Retrying ({attempt}/3)..."
                )

                time.sleep(attempt * 1.5)

    raise RuntimeError(
        f"Could not load {url}: {last_error}"
    )


# =====================================================================
# TABLE URL DISCOVERY
# =====================================================================

def make_fallback_table_url(page_url: str) -> str:

    match = re.search(
        r"[?&]k4=([^&]+)",
        page_url,
    )

    if not match:
        raise RuntimeError(
            f"Could not determine k4 parameter from {page_url}"
        )

    raw_k4 = match.group(1)

    return (
        "https://www.inpaqgp.com/product-table.php?"
        f"k4={raw_k4}&productName="
    )


def discover_table_url(
    session: requests.Session,
    page_url: str,
) -> str:

    print("  Opening category page...")

    # Keep original category k4.
    original_match = re.search(
        r"[?&]k4=([^&]+)",
        page_url,
    )

    original_k4 = (
        original_match.group(1)
        if original_match
        else None
    )

    try:

        response = get(
            session,
            page_url,
        )

        soup = BeautifulSoup(
            response.text,
            "html.parser",
        )

        for iframe in soup.find_all("iframe"):

            src = clean_text(
                iframe.get("src")
            )

            if "product-table.php" not in src:
                continue

            table_url = urljoin(
                page_url,
                src,
            )

            print(
                "  Found product table iframe:"
            )

            print(
                f"    {table_url}"
            )

            # =========================================================
            # AUTOMOTIVE TVS SPECIAL HANDLING
            # =========================================================
            #
            # INPAQ currently returns a different/broken k4 in the
            # Automotive TVS iframe:
            #
            # requested:
            # SF+RPxWXByBMDbagnOhEwg%3D%3D
            #
            # iframe:
            # bApmxlokWrOK0ZTdjxZ6Ig%3D%3D
            #
            # That iframe can return:
            #
            # 無效的資料 menuId
            #
            # So for this category we ignore the altered iframe k4
            # and try the original category k4 directly.
            # =========================================================

            if (
                original_k4
                and original_k4.startswith(
                    "SF+RPxWXByBMDbagnOhEwg"
                )
            ):

                print(
                    "  Automotive TVS category detected."
                )

                print(
                    "  INPAQ changed the iframe k4."
                )

                print(
                    "  Trying original Automotive TVS category k4 directly..."
                )

                return (
                    "https://www.inpaqgp.com/product-table.php?"
                    f"k4={original_k4}"
                    "&productName="
                )

            return table_url

        print(
            "  Product iframe not visible in outer HTML."
        )

    except Exception as exc:

        print(
            f"  Outer page discovery warning: {exc}"
        )

    fallback = make_fallback_table_url(
        page_url
    )

    print(
        "  Using direct product table endpoint:"
    )

    print(
        f"    {fallback}"
    )

    return fallback


# =====================================================================
# TABLE HEADER HANDLING
# =====================================================================

def make_unique_headers(
    headers: list[str],
) -> list[str]:

    output = []
    counts = {}

    for header in headers:

        header = clean_text(header)

        if not header:
            header = "Unnamed"

        counts[header] = (
            counts.get(header, 0) + 1
        )

        if counts[header] == 1:
            output.append(header)

        else:
            output.append(
                f"{header}_{counts[header]}"
            )

    return output


def normalize_header(header: str) -> str:

    header = clean_text(header)

    header = re.sub(
        r"\s+\(",
        " (",
        header,
    )

    return header


# =====================================================================
# TABLE SELECTION
# =====================================================================

def table_score(table) -> int:

    text = clean_text(
        table.get_text(
            " ",
            strip=True,
        )
    ).lower()

    score = 0

    keywords = [
        "part no",
        "part number",
        "vrms",
        "vrwm",
        "vdc",
        "capacitance",
        "esd",
        "voltage",
        "current",
        "package",
        "size",
        "ipp",
        "ppp",
    ]

    for keyword in keywords:

        if keyword in text:
            score += 1

    rows = table.find_all("tr")

    score += min(
        len(rows) // 5,
        10,
    )

    return score


def identify_header_row(rows) -> tuple[int, list[str]]:

    best_index = None
    best_headers = None
    best_score = -1

    for index, row in enumerate(
        rows[:15]
    ):

        cells = row.find_all(
            ["th", "td"]
        )

        if not cells:
            continue

        values = [
            clean_text(
                cell.get_text(
                    " ",
                    strip=True,
                )
            )
            for cell in cells
        ]

        if not values:
            continue

        text = " ".join(
            values
        ).lower()

        score = 0

        if "part no" in text:
            score += 10

        if "part number" in text:
            score += 10

        if "part" in text:
            score += 3

        if "voltage" in text:
            score += 1

        if "esd" in text:
            score += 1

        if "size" in text:
            score += 1

        if "vrwm" in text:
            score += 1

        if "ipp" in text:
            score += 1

        score += (
            sum(bool(v) for v in values)
            / 10
        )

        if score > best_score:

            best_score = score
            best_index = index
            best_headers = values

    if (
        best_index is None
        or best_headers is None
        or best_score < 3
    ):

        raise RuntimeError(
            "Could not identify table header row."
        )

    return (
        best_index,
        make_unique_headers(
            [
                normalize_header(h)
                for h in best_headers
            ]
        ),
    )


def find_part_column(
    columns: list[str],
) -> str | None:

    preferred = [
        "Part No.",
        "Part No",
        "Part Number",
        "Part number",
        "Part Number.",
    ]

    for name in preferred:

        if name in columns:
            return name

    for column in columns:

        normalized = re.sub(
            r"[^a-z0-9]",
            "",
            column.lower(),
        )

        if normalized in {
            "partno",
            "partnumber",
            "part",
        }:

            return column

    return None


# =====================================================================
# BEAUTIFULSOUP PARSER
# =====================================================================

def extract_with_beautifulsoup(
    html: str,
    table_url: str,
) -> pd.DataFrame:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    tables = soup.find_all(
        "table"
    )

    if not tables:

        # -------------------------------------------------------------
        # Detect actual INPAQ error page rather than pretending that
        # pandas/html5lib is the problem.
        # -------------------------------------------------------------

        page_text = clean_text(
            soup.get_text(
                " ",
                strip=True,
            )
        )

        if (
            "menuId" in page_text
            or "無效的資料" in page_text
        ):

            raise RuntimeError(
                "INPAQ returned an invalid menuId instead of a product table."
            )

        raise RuntimeError(
            "No HTML tables were found."
        )

    best_table = max(
        tables,
        key=table_score,
    )

    rows = best_table.find_all(
        "tr"
    )

    if not rows:

        raise RuntimeError(
            "Product table contains no rows."
        )

    (
        header_row_index,
        headers,
    ) = identify_header_row(
        rows
    )

    records = []
    product_urls = []

    for row in rows[
        header_row_index + 1 :
    ]:

        cells = row.find_all(
            "td"
        )

        if not cells:
            continue

        values = [
            clean_text(
                cell.get_text(
                    " ",
                    strip=True,
                )
            )
            for cell in cells
        ]

        if not any(values):
            continue

        # INPAQ sometimes adds checkbox/control columns
        # before the actual specification columns.
        if len(values) > len(headers):

            difference = (
                len(values)
                - len(headers)
            )

            values = values[
                difference:
            ]

        elif len(values) < len(headers):

            values += [""] * (
                len(headers)
                - len(values)
            )

        part_anchor = None

        for anchor in row.find_all(
            "a",
            href=True,
        ):

            href = anchor.get(
                "href",
                "",
            )

            if (
                href
                and not href.lower().startswith(
                    "javascript:"
                )
            ):

                part_anchor = anchor
                break

        if part_anchor:

            product_url = urljoin(
                table_url,
                part_anchor["href"],
            )

        else:

            product_url = ""

        records.append(
            values[:len(headers)]
        )

        product_urls.append(
            product_url
        )

    if not records:

        raise RuntimeError(
            "Product table was found but contained no data rows."
        )

    df = pd.DataFrame(
        records,
        columns=headers,
    )

    part_column = find_part_column(
        list(df.columns)
    )

    if part_column is None:

        raise RuntimeError(
            "Could not identify the Part Number column. "
            f"Columns found: {list(df.columns)}"
        )

    if part_column != "Part No.":

        df = df.rename(
            columns={
                part_column: "Part No."
            }
        )

    if (
        len(product_urls) == len(df)
        and any(product_urls)
    ):

        df["Product URL"] = (
            product_urls
        )

    return df


# =====================================================================
# PANDAS FALLBACK
# =====================================================================

def extract_with_pandas(
    html: str,
) -> pd.DataFrame:

    # -------------------------------------------------------------
    # BEFORE pandas.read_html()
    #
    # Check if INPAQ actually returned an error page.
    #
    # This prevents the misleading:
    #
    #   Import html5lib failed
    #
    # message when there wasn't actually a table to parse anyway.
    # -------------------------------------------------------------

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    page_text = clean_text(
        soup.get_text(
            " ",
            strip=True,
        )
    )

    if (
        "menuId" in page_text
        or "無效的資料" in page_text
    ):

        raise RuntimeError(
            "INPAQ returned an invalid menuId instead of a product table."
        )

    tables = pd.read_html(
        StringIO(html)
    )

    if not tables:

        raise RuntimeError(
            "pandas.read_html found no tables."
        )

    candidates = []

    for table in tables:

        columns = [
            clean_text(c)
            for c in table.columns
        ]

        score = 0

        for column in columns:

            c = column.lower()

            if "part" in c:
                score += 10

            if "voltage" in c:
                score += 1

            if "esd" in c:
                score += 1

            if "size" in c:
                score += 1

            if "vrwm" in c:
                score += 1

            if "ipp" in c:
                score += 1

        score += min(
            len(table) / 10,
            10,
        )

        candidates.append(
            (
                score,
                table,
            )
        )

    _, df = max(
        candidates,
        key=lambda item: item[0],
    )

    df = df.copy()

    df.columns = make_unique_headers(
        [
            normalize_header(
                clean_text(c)
            )
            for c in df.columns
        ]
    )

    part_column = find_part_column(
        list(df.columns)
    )

    if part_column is None:

        raise RuntimeError(
            "pandas fallback could not identify Part Number column."
        )

    if part_column != "Part No.":

        df = df.rename(
            columns={
                part_column: "Part No."
            }
        )

    return df


def scrape_auto_tvs_family_page(
    session: requests.Session,
    page_url: str,
) -> pd.DataFrame:

    print("  Loading Automotive Power TVS family table...")

    response = get(
        session,
        page_url,
    )

    soup = BeautifulSoup(
        response.text,
        "html.parser",
    )

    tables = soup.find_all("table")

    if not tables:
        raise RuntimeError(
            "No Automotive TVS tables found on Power TVS AM page."
        )

    # Find the table containing the actual Power TVS part list.
    product_table = None

    for table in tables:

        text = clean_text(
            table.get_text(
                " ",
                strip=True,
            )
        ).lower()

        if (
            "part number" in text
            and "aec" in text
            and "peak" in text
            and (
                "breakdown" in text
                or "clamping" in text
            )
        ):
            product_table = table
            break

    if product_table is None:
        raise RuntimeError(
            "Could not locate Automotive TVS specification table."
        )

    rows = product_table.find_all("tr")

    if not rows:
        raise RuntimeError(
            "Automotive TVS specification table is empty."
        )

    # -------------------------------------------------------------
    # Build headers.
    #
    # The INPAQ Power TVS table uses multi-level headers.
    # Instead of trusting HTML colspan layout, map the final useful
    # fields into clean columns.
    # -------------------------------------------------------------

    output_columns = [
        "Package",
        "Part No.",
        "AEC Q101",
        "Direction",
        "Peak Power",
        "VRWM (V)",
        "VBR Min (V)",
        "VBR Max (V)",
        "IT (mA)",
        "Vc (V) @ Ipp",
        "Ipp (A)",
        "IR @ VR (uA)",
    ]

    records = []

    for row in rows:

        cells = row.find_all("td")

        if not cells:
            continue

        values = [
            clean_text(
                cell.get_text(
                    " ",
                    strip=True,
                )
            )
            for cell in cells
        ]

        # Ignore headings / malformed rows.
        if len(values) < 12:
            continue

        # The useful rows on this page are:
        #
        # Package
        # Part Number
        # AEC Q101
        # Direction
        # Peak Power
        # VR
        # VBR MIN
        # VBR MAX
        # IT
        # Vc
        # Ipp
        # IR
        #
        values = values[:12]

        part_number = values[1]

        if not part_number:
            continue

        # Filter out header rows accidentally represented as <td>.
        if "part number" in part_number.lower():
            continue

        record = dict(
            zip(
                output_columns,
                values,
            )
        )

        records.append(
            record
        )

    if not records:
        raise RuntimeError(
            "Automotive TVS table was found, but no product rows were parsed."
        )

    df = pd.DataFrame(
        records,
        columns=output_columns,
    )

    for column in df.columns:
        df[column] = df[column].map(
            clean_text
        )

    df = (
        df[
            df["Part No."].ne("")
        ]
        .drop_duplicates(
            subset=["Part No."],
            keep="first",
        )
        .reset_index(
            drop=True
        )
    )

    print(
        f"  Rows scraped from Automotive TVS page: {len(df)}"
    )

    print(
        "  Columns found:"
    )

    for column in df.columns:
        print(
            f"    {column}"
        )

    return df

# =====================================================================
# SCRAPE ONE PAGE
# =====================================================================

def scrape_page(
    session: requests.Session,
    page_url: str,
) -> pd.DataFrame:

    table_url = discover_table_url(
        session,
        page_url,
    )

    print(
        "  Loading complete product table..."
    )

    response = get(
        session,
        table_url,
    )

    # -------------------------------------------------------------
    # Check for INPAQ invalid-menu response immediately.
    # -------------------------------------------------------------

    response_text = clean_text(
        response.text
    )

    if (
        "無效的資料" in response_text
        or "menuId" in response_text
    ):

        raise RuntimeError(
            "INPAQ returned an invalid menuId for this category. "
            "This is an INPAQ page/menu mapping issue, not an "
            "html5lib installation problem."
        )

    try:

        df = extract_with_beautifulsoup(
            response.text,
            table_url,
        )

    except Exception as bs_exc:

        # If we already know it is an INPAQ menuId problem,
        # don't try another HTML parser.
        if (
            "menuId" in str(bs_exc)
            or "無效的資料" in str(bs_exc)
        ):

            raise bs_exc

        print(
            "  BeautifulSoup parser fallback triggered:"
        )

        print(
            f"    {bs_exc}"
        )

        df = extract_with_pandas(
            response.text
        )

    # -------------------------------------------------------------
    # CLEAN DATA
    # -------------------------------------------------------------

    for column in df.columns:

        df[column] = df[column].map(
            clean_text
        )

    if "Part No." not in df.columns:

        raise RuntimeError(
            "Parsed table does not contain Part No.\n"
            f"Columns: {list(df.columns)}"
        )

    part_series = (
        df["Part No."]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    bad_patterns = (
        r"^(?:"
        r"Part No\.?"
        r"|Part Number"
        r"|Show Selection"
        r"|Download ZIP"
        r"|Select 5 items"
        r")$"
    )

    keep_mask = (
        part_series.ne("")
        & ~part_series.str.match(
            bad_patterns,
            case=False,
            na=False,
        )
        & ~part_series.str.contains(
            r"Show Selection|Download ZIP|Select 5 items",
            case=False,
            regex=True,
            na=False,
        )
    )

    df = df[
        keep_mask
    ].copy()

    df = df.reset_index(
        drop=True
    )

    print(
        f"  Rows scraped from this page: {len(df)}"
    )

    print(
        "  Columns found:"
    )

    for column in df.columns:

        print(
            f"    {column}"
        )

    return df


# =====================================================================
# MULTI-PAGE COMBINATION
# =====================================================================

def union_column_order(
    dataframes: list[pd.DataFrame],
) -> list[str]:

    columns = []

    for df in dataframes:

        for column in df.columns:

            if column not in columns:

                columns.append(
                    column
                )

    if "Part No." in columns:

        columns.remove(
            "Part No."
        )

        columns.insert(
            0,
            "Part No.",
        )

    if "Product URL" in columns:

        columns.remove(
            "Product URL"
        )

        columns.append(
            "Product URL"
        )

    return columns


def first_nonempty(
    series: pd.Series,
) -> str:

    for value in series:

        value = clean_text(
            value
        )

        if value:
            return value

    return ""


def combine_dataframes(
    dataframes: list[pd.DataFrame],
) -> pd.DataFrame:

    if not dataframes:

        return pd.DataFrame()

    column_order = union_column_order(
        dataframes
    )

    normalized_frames = []

    for df in dataframes:

        temp = df.copy()

        for column in column_order:

            if column not in temp.columns:

                temp[column] = ""

        temp = temp[
            column_order
        ]

        normalized_frames.append(
            temp
        )

    combined = pd.concat(
        normalized_frames,
        ignore_index=True,
        sort=False,
    )

    for column in combined.columns:

        combined[column] = (
            combined[column]
            .map(clean_text)
        )

    if "Part No." in combined.columns:

        combined = (
            combined
            .groupby(
                "Part No.",
                sort=False,
                as_index=False,
            )
            .agg(first_nonempty)
        )

    final_order = [
        c
        for c in column_order
        if c in combined.columns
    ]

    combined = combined[
        final_order
    ]

    return combined.reset_index(
        drop=True
    )


# =====================================================================
# EXCEL OUTPUT
# =====================================================================

def save_excel(
    df: pd.DataFrame,
    output_path: Path,
    sheet_name: str,
) -> None:

    sheet_name = sheet_name[:31]

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
        )

        ws = writer.book[
            sheet_name
        ]

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = (
            ws.dimensions
        )

        # No deprecated cell.font.copy()
        for cell in ws[1]:

            new_font = copy(
                cell.font
            )

            new_font.bold = True

            cell.font = new_font

        for column_cells in ws.columns:

            max_len = 0

            for cell in column_cells:

                value = (
                    ""
                    if cell.value is None
                    else str(cell.value)
                )

                max_len = max(
                    max_len,
                    len(value),
                )

            width = min(
                max(
                    max_len + 2,
                    12,
                ),
                45,
            )

            ws.column_dimensions[
                column_cells[0].column_letter
            ].width = width


# =====================================================================
# SCRAPE ONE OUTPUT GROUP
# =====================================================================

def scrape_group(
    session: requests.Session,
    group: dict,
) -> pd.DataFrame:

    print()

    print(
        "=" * 72
    )

    print(
        f"SCRAPING {group['name']}"
    )

    print(
        "=" * 72
    )

    frames = []

    for index, page_url in enumerate(
        group["urls"],
        start=1,
    ):

        print()

        print(
            f"Source page {index}/{len(group['urls'])}"
        )

        print(
            page_url
        )

        try:

            if group.get(
                "special_auto_tvs",
                False,
            ):

                df = scrape_auto_tvs_family_page(
                    session,
                    page_url,
                )

            else:

                df = scrape_page(
                    session,
                    page_url,
                )

            frames.append(
                df
            )
        except Exception as exc:

            print()

            print(
                f"  ERROR scraping source page {index}:"
            )

            print(
                f"    {exc}"
            )

            continue

    if not frames:

        raise RuntimeError(
            f"No data was successfully scraped for {group['name']}."
        )

    combined = combine_dataframes(
        frames
    )

    print()

    print(
        f"Combined rows: {len(combined)}"
    )

    if "Part No." in combined.columns:

        print(
            f"Unique parts: "
            f"{combined['Part No.'].nunique()}"
        )

    print(
        f"Final columns: {len(combined.columns)}"
    )

    print()

    print(
        "Final header list:"
    )

    for column in combined.columns:

        print(
            f"  {column}"
        )

    print()

    print(
        "Parts:"
    )

    if "Part No." in combined.columns:

        for part in combined[
            "Part No."
        ].tolist():

            print(
                f"  {part}"
            )

    return combined


# =====================================================================
# MAIN
# =====================================================================

def main() -> int:

    session = make_session()

    successful = []
    failed = []

    print()

    print(
        "=" * 72
    )

    print(
        "INPAQ PRODUCT SCRAPER"
    )

    print(
        "=" * 72
    )

    for group in SCRAPE_GROUPS:

        try:

            df = scrape_group(
                session,
                group,
            )

            save_excel(
                df,
                group["output"],
                group["sheet_name"],
            )

            output_path = (
                group["output"].resolve()
            )

            successful.append(
                str(output_path)
            )

            print()

            print(
                "Saved compiled file:"
            )

            print(
                f"  {output_path}"
            )

        except Exception as exc:

            failed.append(
                (
                    group["name"],
                    str(exc),
                )
            )

            print()

            print(
                f"ERROR processing {group['name']}:"
            )

            print(
                f"  {exc}"
            )

    # =================================================================
    # SUMMARY
    # =================================================================

    print()

    print(
        "=" * 72
    )

    print(
        "SCRAPE COMPLETE"
    )

    print(
        "=" * 72
    )

    if successful:

        print()

        print(
            "Successfully created:"
        )

        for filename in successful:

            print(
                f"  {filename}"
            )

    if failed:

        print()

        print(
            "Failed categories:"
        )

        for name, error in failed:

            print(
                f"  {name}: {error}"
            )

    print()

    if failed:
        return 1

    return 0

# =====================================================================
# CROSS-REFERENCE PARSING
# =====================================================================

def _inpaq_safe_strip(value):
    """Return a clean string, or '-' for an empty/NaN value."""
    if value is None:
        return "-"

    try:
        if pd.isna(value):
            return "-"
    except Exception:
        pass

    value = str(value).strip()
    return value if value and value.lower() != "nan" else "-"


def _inpaq_get_value(row, keywords):
    """
    Finds a column by header keyword.

    First tries an exact normalized header match, then falls back to
    substring matching.
    """
    normalized_keywords = [
        re.sub(r"\s+", " ", str(k).strip()).lower()
        for k in keywords
    ]

    # Exact match first
    for col in row.index:
        normalized_col = re.sub(
            r"\s+",
            " ",
            str(col).strip()
        ).lower()

        if normalized_col in normalized_keywords:
            return row[col]

    # Substring fallback
    for col in row.index:
        normalized_col = re.sub(
            r"\s+",
            " ",
            str(col).strip()
        ).lower()

        for keyword in normalized_keywords:
            if keyword in normalized_col:
                return row[col]

    return None


def _inpaq_numeric(value):
    """Extract first numeric value from an INPAQ cell."""
    value = _inpaq_safe_strip(value)

    if value == "-":
        return None

    match = re.search(
        r"[-+]?\d+(?:\.\d+)?",
        value
    )

    if not match:
        return None

    try:
        return float(match.group())
    except ValueError:
        return None


def _inpaq_direction(value):
    """Normalize INPAQ direction into the format used by applesauce."""
    value = _inpaq_safe_strip(value)

    if value == "-":
        return "-"

    text = value.lower().replace("-", " ")

    if (
        "bi direction" in text
        or "bidirection" in text
        or text.strip() == "bi"
    ):
        return "Bidirectional"

    if (
        "uni direction" in text
        or "unidirection" in text
        or text.strip() == "uni"
    ):
        return "Unidirectional"

    return value


def _inpaq_package_from_size(value):
    """
    INPAQ ESD/standard TVS sheets often give:
        0402 (1005)
        0603 (1608)

    Keep that as the package value.
    """
    value = _inpaq_safe_strip(value)

    if value == "-":
        return "-"

    return value


def _parse_inpaq_esd_row(
    part_row,
    found_df_name,
    part_number,
):
    """
    Parse INPAQ ESD / Automotive ESD.

    Headers include:
        Part No.
        L size(mm)
        W size(mm)
        T size(mm)
        Size codeInch (mm)
        VRMS (V) max
        IL (A) TYP
        Vt (V) TYP
        Vc (V) TYP
        Cp Condition (KHZ)
        CP (pF) TYP
        ESD Contact discharge (kV) TYP
        ESD Air discharge (kV) TYP
        ESD pulse withstand (pulses)
    """

    is_auto = "auto" in found_df_name.lower()

    specs = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Automotive" if is_auto else "Non-Automotive",

        "Package": "-",
        "Package Info": {},

        "Direction": "-",
        "Channels": "1",

        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",

        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",

        "Capacitance": "-",

        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
    }

    # ---------------------------------------------------------
    # PACKAGE
    # ---------------------------------------------------------

    package_raw = _inpaq_get_value(
        part_row,
        [
            "Size codeInch (mm)",
            "Size code Inch (mm)",
        ],
    )

    package = _inpaq_package_from_size(
        package_raw
    )

    specs["Package"] = package

    specs["Package Info"] = {
        "name": package,
    }

    # ---------------------------------------------------------
    # VRWM
    #
    # INPAQ calls this VRMS in this particular ESD table.
    # ---------------------------------------------------------

    vrwm = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "VRMS (V) max",
                "VRMS (V) Max",
            ],
        )
    )

    if vrwm is not None:
        specs[
            "Voltage - Reverse Standoff (Typ)"
        ] = f"{vrwm:g} V"

    # ---------------------------------------------------------
    # CLAMPING VOLTAGE
    # ---------------------------------------------------------

    vc = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Vc (V) TYP",
                "VC (V) TYP",
            ],
        )
    )

    if vc is not None:
        specs[
            "Voltage - Clamping (Max) @ Ipp"
        ] = f"{vc:g} V"

    # ---------------------------------------------------------
    # CAPACITANCE
    #
    # Prefer max if a future/general sheet contains it.
    # Otherwise use typical.
    # ---------------------------------------------------------

    cap = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Cp (pF) max",
                "CP (pF) max",
                "Cp (pF) Max",
            ],
        )
    )

    if cap is None:
        cap = _inpaq_numeric(
            _inpaq_get_value(
                part_row,
                [
                    "CP (pF) TYP",
                    "Cp (pF) TYP",
                    "Cp (pF) Typical",
                ],
            )
        )

    if cap is not None:
        specs["Capacitance"] = f"{cap:g} pF"

    # ---------------------------------------------------------
    # IEC 61000-4-2 CONTACT ESD
    # ---------------------------------------------------------

    esd = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "ESD Contact discharge (kV) TYP",
                "ESD Contact Discharge (kV)",
                "ESD Contact discharge (kV)",
            ],
        )
    )

    if esd is not None:
        specs[
            "IEC 61000-4-2"
        ] = f"±{esd:g} kV"

    # ---------------------------------------------------------
    # LEAKAGE
    #
    # INPAQ ESD calls this IL.
    # Keep the value if present.
    # ---------------------------------------------------------

    leakage = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "IL (A) TYP",
                "IL (A) Typ",
            ],
        )
    )

    if leakage is not None:
        specs[
            "Leakage Current (Ir)"
        ] = f"{leakage:g} A"

    return specs


def _parse_inpaq_tvs_row(
    part_row,
    found_df_name,
    part_number,
):
    """
    Parse regular INPAQ TVS.

    Handles headers shown in inpaq_tvs_specs.xlsx, including:
        Part No.
        L size(mm)
        W size(mm)
        T size(mm)
        Size codeInch (mm)
        Direction Bi/Uni
        Channel
        VRWM (V) Max
        Cp (pF) Typical
        Cp (pF) Max
        Ipp (A)
        ESD Contact Discharge
        ESD Air Discharge
        VB(V) Typical
        VB(V) Max
        Ppp (W)
    """

    specs = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Non-Automotive",

        "Package": "-",
        "Package Info": {},

        "Direction": "-",
        "Channels": "1",

        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",

        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",

        "Capacitance": "-",

        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
    }

    # Package
    package = _inpaq_package_from_size(
        _inpaq_get_value(
            part_row,
            [
                "Size codeInch (mm)",
                "Size code Inch (mm)",
            ],
        )
    )

    specs["Package"] = package
    specs["Package Info"] = {
        "name": package,
    }

    # Direction
    specs["Direction"] = _inpaq_direction(
        _inpaq_get_value(
            part_row,
            [
                "Direction Bi/Uni",
                "Direction",
            ],
        )
    )

    # Channels
    channels = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Channel",
                "Channels",
            ],
        )
    )

    if channels is not None:
        specs["Channels"] = f"{channels:g}"

    # VRWM
    vrwm = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "VRWM (V) Max",
                "VRWM (V) MAX",
                "VRWM (V)",
            ],
        )
    )

    if vrwm is not None:
        specs[
            "Voltage - Reverse Standoff (Typ)"
        ] = f"{vrwm:g} V"

    # Capacitance - prefer max
    cap = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Cp (pF) Max",
                "CP (pF) Max",
            ],
        )
    )

    if cap is None:
        cap = _inpaq_numeric(
            _inpaq_get_value(
                part_row,
                [
                    "Cp (pF) Typical",
                    "CP (pF) Typical",
                ],
            )
        )

    if cap is not None:
        specs["Capacitance"] = f"{cap:g} pF"

    # Ipp / surge current
    ipp = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Ipp (A)",
                "IPP (A)",
            ],
        )
    )

    if ipp is not None:
        specs[
            "IEC 61000-4-5"
        ] = f"{ipp:g} A (8/20µs)"

    # ESD contact
    esd = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "ESD Contact Discharge (V)",
                "ESD Contact Discharge (kV)",
                "ESD Contact discharge (kV) TYP",
            ],
        )
    )

    if esd is not None:
        # Screenshot values such as 20 are ratings in kV despite
        # INPAQ's exported header sometimes saying "(V)".
        specs[
            "IEC 61000-4-2"
        ] = f"±{esd:g} kV"

    # Peak pulse power
    power = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Ppp (W)",
                "PPP (W)",
                "Peak Power",
            ],
        )
    )

    if power is not None:
        specs[
            "Power Dissipation (Pd)"
        ] = f"{power:g} W"

    # IMPORTANT:
    # VB is breakdown voltage, NOT clamping voltage.
    # Do not map VB(V) Typical / Max to Vclamp.

    return specs


def _parse_inpaq_auto_tvs_row(
    part_row,
    found_df_name,
    part_number,
):
    """
    Parse INPAQ Automotive TVS.

    Headers shown in inpaq_auto_tvs_specs.xlsx:
        Part No.
        Package
        AEC Q101
        Direction
        Peak Power
        VRWM (V)
        VBR Min (V)
        VBR Max (V)
        IT (mA)
        Vc (V) @ Ipp
        Ipp (A)
        IR @ VR (uA)
    """

    specs = {
        "Device Name": part_number,
        "Source File": found_df_name,
        "Grade": "Automotive",

        "Package": "-",
        "Package Info": {},

        "Direction": "-",
        "Channels": "1",

        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",

        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",

        "Capacitance": "-",

        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
    }

    # Package
    package = _inpaq_safe_strip(
        _inpaq_get_value(
            part_row,
            ["Package"],
        )
    )

    specs["Package"] = package
    specs["Package Info"] = {
        "name": package,
    }

    # Direction
    specs["Direction"] = _inpaq_direction(
        _inpaq_get_value(
            part_row,
            ["Direction"],
        )
    )

    # VRWM
    vrwm = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "VRWM (V)",
                "VRWM (V) Max",
            ],
        )
    )

    if vrwm is not None:
        specs[
            "Voltage - Reverse Standoff (Typ)"
        ] = f"{vrwm:g} V"

    # Clamping voltage
    vc = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Vc (V) @ Ipp",
                "VC (V) @ Ipp",
            ],
        )
    )

    if vc is not None:
        specs[
            "Voltage - Clamping (Max) @ Ipp"
        ] = f"{vc:g} V"

    # Ipp / surge
    ipp = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Ipp (A)",
                "IPP (A)",
            ],
        )
    )

    if ipp is not None:
        specs[
            "IEC 61000-4-5"
        ] = f"{ipp:g} A (8/20µs)"

    # Peak power
    power = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "Peak Power",
                "Peak Power (W)",
                "Ppp (W)",
            ],
        )
    )

    if power is not None:
        specs[
            "Power Dissipation (Pd)"
        ] = f"{power:g} W"

    # Leakage
    leakage = _inpaq_numeric(
        _inpaq_get_value(
            part_row,
            [
                "IR @ VR (uA)",
                "IR @ VR (µA)",
            ],
        )
    )

    if leakage is not None:
        specs[
            "Leakage Current (Ir)"
        ] = f"{leakage:g} uA"

    return specs


def fetch_inpaq_specs_from_excel(
    part_number,
    inpaq_dfs,
):
    """
    Search all four INPAQ databases and return the standardized
    competitor-spec dictionary expected by applesauce.py.

    Expected dictionary keys:
        ESD
        Auto_ESD
        TVS
        Auto_TVS
    """

    if part_number is None or not inpaq_dfs:
        return None

    wanted_part = str(
        part_number
    ).strip().lower()

    # Automotive first so an automotive part cannot accidentally
    # resolve against a commercial sheet if there is overlap.
    search_priority = [
        "Auto_ESD",
        "Auto_TVS",
        "ESD",
        "TVS",
    ]

    ordered_keys = [
        key
        for key in search_priority
        if key in inpaq_dfs
    ]

    ordered_keys.extend(
        key
        for key in inpaq_dfs
        if key not in ordered_keys
    )

    for df_name in ordered_keys:

        df = inpaq_dfs[df_name]

        if df is None or df.empty:
            continue

        # Find Part No. column
        part_col = None

        for col in df.columns:
            normalized = re.sub(
                r"[^a-z0-9]",
                "",
                str(col).lower(),
            )

            if normalized in {
                "partno",
                "partnumber",
            }:
                part_col = col
                break

        if part_col is None:
            continue

        matches = df[
            df[part_col]
            .astype(str)
            .str.strip()
            .str.lower()
            == wanted_part
        ]

        if matches.empty:
            continue

        part_row = matches.iloc[0]

        exact_part_number = _inpaq_safe_strip(
            part_row[part_col]
        )

        print(
            f"Found specs for Part '{exact_part_number}' "
            f"in INPAQ database '{df_name}'."
        )

        # ---------------------------------------------------------
        # Dispatch to the correct parser
        # ---------------------------------------------------------

        df_name_lower = df_name.lower()

        if (
            "auto" in df_name_lower
            and "tvs" in df_name_lower
        ):
            return _parse_inpaq_auto_tvs_row(
                part_row,
                df_name,
                exact_part_number,
            )

        if "esd" in df_name_lower:
            return _parse_inpaq_esd_row(
                part_row,
                df_name,
                exact_part_number,
            )

        if "tvs" in df_name_lower:
            return _parse_inpaq_tvs_row(
                part_row,
                df_name,
                exact_part_number,
            )

    print(
        f"Part '{part_number}' not found in any INPAQ database."
    )

    return None

if __name__ == "__main__":

    raise SystemExit(
        main()
    )