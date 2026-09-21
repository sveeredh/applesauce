"""
build_batch_input.py -- assemble a mode-2 batch input file from every
competitor spec sheet in a folder.

Scans for files whose name contains "specs", skips the DigiKey and TI sheets,
finds each file's part-number column, and writes a two-column Excel file:

    Competitor Parts | Competitor Name

which is exactly what applesauce.py's "(2) Batch File Cross" mode expects.

Usage:
    python build_batch_input.py                     # scan ., write batch_input.xlsx
    python build_batch_input.py --dir C:\\path\\to\\data
    python build_batch_input.py --out all_parts.xlsx
    python build_batch_input.py --dry-run           # report only, write nothing

The manufacturer is taken from the filename prefix (nexperia_specs_tvs.xls ->
"nexperia"), which is the same token get_competitor_specs_leniently matches on,
so the output feeds straight back into the cross.
"""

import argparse
import os
import re
import sys
import glob
import pandas as pd
import contextlib
import io


# JJM's workbooks carry a styles.xml openpyxl rejects, which kills pd.read_excel
# before any cell is read. jjm_scrape has a reader that goes straight to the
# sheet XML; imported softly so this script still runs without that module.
try:
    from jjm_scrape import read_xlsx_raw as _read_xlsx_raw
except ImportError:
    _read_xlsx_raw = None

# Some sheets weld the datasheet hyperlink's text onto the part number, so the
# cell reads "JEB03CX-AU<a run of spaces>DataSheet". Left alone it blows past
# MAX_PART_LENGTH and the part is dropped.
DATASHEET_SUFFIX_RE = re.compile(r"\s{2,}data\s*sheet\s*$", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Which files to read
# ---------------------------------------------------------------------------

SPEC_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv")

# A file is skipped if its manufacturer token is one of these. "ti" is matched
# as a whole token, not a substring -- otherwise "vishay_specs_protection"
# (protec-TI-on) and anything else containing "ti" would be dropped too.
EXCLUDED_MANUFACTURERS = {"ti", "digikey", "digi"}

# ...and skipped if any of these appear anywhere in the filename.
EXCLUDED_NAME_FRAGMENTS = ("digikey", "digi_", "_digi")


# ---------------------------------------------------------------------------
# Which column holds the part number
# ---------------------------------------------------------------------------

# Ordered most specific first -- the first keyword that matches a header wins.
# Add to this list if a new competitor sheet uses a heading not covered here.
PART_COL_KEYWORDS = [
    "eaton part number",
    "manufacturer part number",
    "mfr part",
    "type number",              # Nexperia
    "article number",           # Diotec
    "article no",
    "product or part number",   # TI-style
    "part number",
    "part no",
    "part #",
    "orderable part",
    "device",
    "device name",
    "model number",
    "opn",
    "product group", 
    "产品名称", 
    "Nichtek P/N",
    "product name"
]

# These are too generic to match as substrings -- Semtech's column is exactly
# "Parts" and AOS's is exactly "Product", but "Replacement Parts" and "Product
# type" would also contain them. They are only accepted on an exact match, and
# only after every keyword above has failed.
PART_COL_EXACT_KEYWORDS = [
    "parts",                    # Semtech
    "product",                  # AOS
    "part",
    "type",
]

# Headers that contain a keyword but are not the manufacturer's own part number.
PART_COL_BLOCKLIST = [
    "product type", "product family", "product line", "product group",
    "product category", "product status", "part status", "number of",
    "nr of", "# of", "part marking", "marking code",
    # cross-reference columns -- these hold somebody else's part numbers
    "replacement", "competitor", "cross", "equivalent", "alternative",
    "similar", "obsolete", "replaces", "superseded",
]

# Per-file overrides, if auto-detection ever picks the wrong column. Key is the
# filename (or any substring of it), value is the exact column heading.
#   COLUMN_OVERRIDES = {"amazing_specs.xlsx": "Part Number"}
COLUMN_OVERRIDES = {
    "onsemi_esd_specs.csv": "Product Group",
    "onsemi_zener_specs.csv": "Product Group",
}

# Manufacturer name overrides, keyed on the filename prefix. Only needed when
# the prefix is not what you want written into the Competitor Name column.
MANUFACTURER_OVERRIDES = {}

MAX_HEADER_SCAN_ROWS = 20   # spec sheets bury headers up to ~row 10
MAX_PART_LENGTH = 60        # anything longer is a description, not a part


def _norm(text):
    """Lowercase with all whitespace removed, for header comparison."""
    return re.sub(r'[\s\u00a0]+', '', str(text)).lower()


def manufacturer_from_filename(filename):
    """nexperia_specs_tvs.xls -> 'nexperia'; central_zener_specs.xlsx -> 'central'."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    token = re.split(r'[_\-\s]+', stem)[0].lower()
    return MANUFACTURER_OVERRIDES.get(token, token)


def should_skip(filepath):
    """(skip?, reason). Applies the specs/DigiKey/TI filtering rules."""
    name = os.path.basename(filepath)
    lower = name.lower()

    if not lower.endswith(SPEC_EXTENSIONS):
        return True, "not a spreadsheet"
    if "specs" not in lower:
        return True, "no 'specs' in filename"
    if any(frag in lower for frag in EXCLUDED_NAME_FRAGMENTS):
        return True, "DigiKey file"

    mfr = manufacturer_from_filename(name)
    if mfr in EXCLUDED_MANUFACTURERS:
        return True, f"excluded manufacturer '{mfr}'"
    if lower.startswith("digikey") or lower.startswith("digi"):
        return True, "DigiKey file"

    return False, None


def find_part_column(columns, filename=""):
    """Pick the part-number column from a list of headers. None if none fit."""
    for override_key, override_col in COLUMN_OVERRIDES.items():
        if override_key.lower() in filename.lower() and override_col in columns:
            return override_col

    normalized = [(col, _norm(col)) for col in columns]
    allowed = [(col, norm_col) for col, norm_col in normalized
               if not any(_norm(bad) in norm_col for bad in PART_COL_BLOCKLIST)]

    # Pass 1: specific headings, matched as substrings.
    for keyword in PART_COL_KEYWORDS:
        key = _norm(keyword)
        for col, norm_col in allowed:
            if key in norm_col:
                return col

    # Pass 2: generic headings, exact match only.
    for keyword in PART_COL_EXACT_KEYWORDS:
        key = _norm(keyword)
        for col, norm_col in allowed:
            if key == norm_col:
                return col

    return None


def _read_raw(filepath, sheet_name=0):
    """Read a sheet with no header applied, so the header row can be located."""
    if filepath.lower().endswith(".csv"):
        return pd.read_csv(
            filepath,
            header=None,
            dtype=str,
            on_bad_lines="skip",
            encoding="latin-1",
            low_memory=False,
        )

    # Panjit .xls files have minor workbook corruption that xlrd
    # can safely ignore.
    if (
        (
            "panjit_" in os.path.basename(filepath).lower()
            or "jiangsu_" in os.path.basename(filepath).lower() 
            or "yangjie_" in os.path.basename(filepath).lower()

        )
        and filepath.lower().endswith(".xls")
    ):
        with contextlib.redirect_stdout(io.StringIO()):
            return pd.read_excel(
                filepath,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                engine="xlrd",
                engine_kwargs={
                    "ignore_workbook_corruption": True
                },
            )

    try:
        return pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
        )
    except Exception:
        # A workbook whose styles.xml openpyxl won't parse still has perfectly
        # good cell data. Read it without the styles rather than losing the file.
        if _read_xlsx_raw is None or not filepath.lower().endswith((".xlsx", ".xlsm")):
            raise
        return _read_xlsx_raw(filepath, header_row=None).astype(object)

def _sheet_names(filepath):
    if filepath.lower().endswith(".csv"):
        return [0]

    try:
        if (
            (
                "panjit_" in os.path.basename(filepath).lower()
                or "jiangsu_" in os.path.basename(filepath).lower()
            )
            and filepath.lower().endswith(".xls")
        ):
            with contextlib.redirect_stdout(io.StringIO()):
                return pd.ExcelFile(
                    filepath,
                    engine="xlrd",
                    engine_kwargs={
                        "ignore_workbook_corruption": True
                    },
                ).sheet_names

        return pd.ExcelFile(
            filepath
        ).sheet_names

    except Exception:
        return [0]


def extract_parts_from_sheet(raw, filename):
    """
    Locate the header row, then the part column, then return the parts.
    Returns (parts, header_row_index, column_name) or (None, None, None).
    """
    if raw is None or raw.empty:
        return None, None, None

    for header_row in range(min(MAX_HEADER_SCAN_ROWS, len(raw))):
        row = raw.iloc[header_row]
        if row.notna().sum() < 2:
            continue
        part_col = find_part_column(list(row.values), filename)
        if part_col is None:
            continue

        col_index = list(row.values).index(part_col)
        series = raw.iloc[header_row + 1:, col_index]

        parts = []
        for value in series:
            if pd.isna(value):
                continue
            text = DATASHEET_SUFFIX_RE.sub("", str(value)).strip()
            if not text or text in ("-", "nan", "None", "N/A", "--"):
                continue
            if len(text) > MAX_PART_LENGTH:
                continue
            if _norm(text) == _norm(part_col):     # repeated header row
                continue
            parts.append(text)

        if parts:
            return parts, header_row, str(part_col)

    return None, None, None


def extract_parts(filepath):
    """Pull parts from every sheet in a file. Returns (parts, notes)."""
    all_parts = []
    notes = []

    for sheet in _sheet_names(filepath):
        try:
            raw = _read_raw(filepath, sheet)
        except Exception as exc:
            notes.append(f"sheet '{sheet}' unreadable ({exc})")
            continue

        parts, header_row, part_col = extract_parts_from_sheet(raw, filepath)
        if parts:
            all_parts.extend(parts)
            label = "" if sheet in (0, None) else f"[{sheet}] "
            notes.append(f"{label}col '{part_col}' (header row {header_row}), {len(parts)} parts")
        else:
            label = "" if sheet in (0, None) else f"sheet '{sheet}': "
            notes.append(f"{label}no part-number column found")

    return all_parts, notes


def collect(directory):
    """Walk the directory and return (rows, report)."""
    rows = []
    report = []

    candidates = sorted(
        p for p in glob.glob(os.path.join(directory, "*"))
        if os.path.isfile(p)
    )

    # If a stem exists as both .csv and .xlsx, keep only the .xlsx -- the main
    # script converts CSV to XLSX on first run and they would double up.
    by_stem = {}
    for path in candidates:
        skip, _ = should_skip(path)
        if skip:
            continue
        stem = os.path.splitext(os.path.basename(path))[0].lower()
        current = by_stem.get(stem)
        if current is None or (current.lower().endswith(".csv") and not path.lower().endswith(".csv")):
            by_stem[stem] = path
    preferred = set(by_stem.values())

    for path in candidates:
        name = os.path.basename(path)
        skip, reason = should_skip(path)
        if skip:
            if "specs" in name.lower():
                report.append((name, "SKIPPED", reason))
            continue
        if path not in preferred:
            report.append((name, "SKIPPED", "superseded by the .xlsx of the same name"))
            continue

        mfr = manufacturer_from_filename(name)
        parts, notes = extract_parts(path)
        if not parts:
            report.append((name, "NO PARTS", "; ".join(notes) or "empty"))
            continue

        for part in parts:
            rows.append({"Competitor Parts": part, "Competitor Name": mfr})
        report.append((name, f"{len(parts)} parts -> {mfr}", "; ".join(notes)))

    return rows, report


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dir", default=".", help="folder holding the spec files (default: .)")
    parser.add_argument("--out", default="batch_input.xlsx", help="output file (default: batch_input.xlsx)")
    parser.add_argument("--dry-run", action="store_true", help="report what would be pulled, write nothing")
    parser.add_argument("--keep-duplicates", action="store_true", help="do not drop repeated part/manufacturer pairs")
    args = parser.parse_args()

    if not os.path.isdir(args.dir):
        print(f"ERROR: '{args.dir}' is not a folder.")
        sys.exit(1)

    rows, report = collect(args.dir)

    print(f"\nScanned '{os.path.abspath(args.dir)}'\n")
    width = max([len(name) for name, _, _ in report] + [12])
    for name, status, detail in report:
        print(f"  {name:<{width}}  {status}")
        if detail:
            print(f"  {'':<{width}}    {detail}")

    if not rows:
        print("\nNo parts found. Check that the spec files are in this folder.")
        sys.exit(1)

    df = pd.DataFrame(rows, columns=["Competitor Parts", "Competitor Name"])
    total = len(df)
    if not args.keep_duplicates:
        df = df.drop_duplicates().reset_index(drop=True)

    print(f"\nTotal: {total} parts"
          + (f" ({total - len(df)} duplicates dropped, {len(df)} written)" if total != len(df) else ""))
    print("By manufacturer:")
    for mfr, count in df["Competitor Name"].value_counts().items():
        print(f"  {mfr:<20} {count}")

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return

    df.to_excel(args.out, index=False)
    print(f"\nWrote {len(df)} rows to '{os.path.abspath(args.out)}'")


if __name__ == "__main__":
    main()