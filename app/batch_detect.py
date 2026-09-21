"""
batch_detect.py -- read an uploaded batch file without insisting on its layout.

A batch file does not have to say "Competitor Parts" / "Competitor Name". When
those headers are missing, the two columns are found by what is in them: the
competitor column is the one whose cells name companies in the alias table,
and the part column is the one whose cells turn up in that competitor's
database. Title rows, header rows with any wording, no header row at all,
extra columns and any column order are all fine.

Used by api.py:
    batch_df, detected = read_batch_file(request.files["file"], all_dfs)
"""

import contextlib
import io
import re

import pandas as pd

from applesauce import get_competitor_specs_leniently, _COMPETITOR_NAME_ALIASES

BATCH_PART_COL = "Competitor Parts"
BATCH_NAME_COL = "Competitor Name"

_DETECT_SCAN_ROWS = 500      # rows examined when scoring columns
_PART_SAMPLE_SIZE = 12       # database lookups per candidate part column
_MIN_NAME_RATIO = 0.5        # share of cells that must name a known competitor
_MIN_PART_RATIO = 0.25       # share of sampled cells that must be real parts

# Header text accepted as-is, so a file in the documented format skips detection.
_PART_HEADERS = {"competitor parts", "competitor part", "part names", "part name"}
_NAME_HEADERS = {"competitor name", "competitor", "manufacturer"}

_ALIAS_TO_KEY = {}
for _canonical, _aliases in _COMPETITOR_NAME_ALIASES.items():
    _ALIAS_TO_KEY[_canonical.lower()] = _aliases[0]
    for _alias in _aliases:
        _ALIAS_TO_KEY[_alias.lower()] = _aliases[0]

# Searched inside longer text ("Littelfuse Inc." -> littelfuse). Short aliases
# like "lf", "az" and "yj" are exact-match only: inside other text they turn up
# in part numbers ("SMBJ5.0A-LF") and would make a part column look like a
# competitor column.
_ALIAS_SEARCH = sorted(
    ((alias, key) for alias, key in _ALIAS_TO_KEY.items() if len(alias) >= 4),
    key=lambda pair: -len(pair[0]),
)

_PART_LIKE_RE = re.compile(r"^(?=.*\d)[A-Za-z0-9][A-Za-z0-9.\-_/+#,() ]{1,39}$")


def competitor_key(value):
    """Cell text -> the lookup key of the competitor it names, or None."""
    text = re.sub(r"\s+", " ", str(value if value is not None else "")).strip().lower()
    if not text or text == "nan":
        return None
    if text in _ALIAS_TO_KEY:
        return _ALIAS_TO_KEY[text]
    for alias, key in _ALIAS_SEARCH:
        if re.search(r"(?<!\w)" + re.escape(alias) + r"(?!\w)", text):
            return key
    return None


def _looks_like_part(value):
    text = str(value if value is not None else "").strip()
    return bool(_PART_LIKE_RE.match(text)) and text.count(" ") <= 2


def _norm_pn(text):
    return re.sub(r"[\W_]+", "", str(text)).upper()


def _is_known_part(part, key, all_dfs):
    """
    Whether the part is in that competitor's database. The dispatcher matches
    leniently (partial, fuzzy), so a hit only counts when the part it resolved
    to is the input or a prefix/extension of it -- "SMBJ5.0A-13-F" against
    "SMBJ5.0A" is a hit, a description column fuzzy-matching some random part
    is not.
    """
    try:
        # The dispatcher narrates every lookup; a dozen misses per column would
        # bury the real log. redirect_stdout is process-wide, so a batch job's
        # prints in another thread go quiet for the second this takes.
        with contextlib.redirect_stdout(io.StringIO()):
            specs = get_competitor_specs_leniently(part, key, all_dfs)
    except Exception:
        return False
    if not specs:
        return False
    found, wanted = _norm_pn(specs.get("Device Name", "")), _norm_pn(part)
    return bool(found) and bool(wanted) and (wanted.startswith(found) or found.startswith(wanted))


def _column_label(raw, col, header_row):
    """What to call a column in a message: its header text, else its letter."""
    if header_row is not None:
        text = str(raw.iat[header_row, col]).strip()
        if text and text.lower() != "nan":
            return f"'{text}'"
    letters, n = "", col + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"column {letters}"


def _detect_batch_columns(raw, all_dfs):
    """
    raw: the sheet read with header=None and dtype=str.
    Returns (part_col, name_col, first_data_row, header_row) as positions, or
    raises ValueError saying what could not be found.
    """
    cell = lambda r, c: str(raw.iat[r, c]).strip() if pd.notna(raw.iat[r, c]) else ""

    # Fast path: the documented headers, anywhere in the first few rows.
    for r in range(min(10, len(raw))):
        headers = {cell(r, c).lower(): c for c in range(raw.shape[1]) if cell(r, c)}
        part_col = next((headers[h] for h in _PART_HEADERS if h in headers), None)
        name_col = next((headers[h] for h in _NAME_HEADERS if h in headers), None)
        if part_col is not None and name_col is not None and part_col != name_col:
            return part_col, name_col, r + 1, r

    scan = range(min(_DETECT_SCAN_ROWS, len(raw)))

    # Competitor column: the one whose cells most often name a known competitor.
    best_name = None
    for c in range(raw.shape[1]):
        filled = [cell(r, c) for r in scan if cell(r, c)]
        if not filled:
            continue
        hits = sum(1 for v in filled if competitor_key(v))
        ratio = hits / len(filled)
        if hits and ratio >= _MIN_NAME_RATIO and (best_name is None or (ratio, hits) > best_name[1:]):
            best_name = (c, ratio, hits)
    if best_name is None:
        raise ValueError(
            "Couldn't find a column of competitor names. Add a column naming the "
            "manufacturer of each part (e.g. 'Nexperia', 'onsemi', 'lf')."
        )
    name_col = best_name[0]

    # Data starts at the first row that names a competitor; anything above it
    # is a title or header row.
    first_data_row = next(r for r in range(len(raw)) if competitor_key(cell(r, name_col)))
    header_row = first_data_row - 1 if first_data_row > 0 else None

    rows = [r for r in range(first_data_row, len(raw)) if competitor_key(cell(r, name_col))]

    # Part column: of the columns whose cells look like part numbers, the one
    # whose cells are actually found in the named competitor's database.
    # Sampled evenly across the file, not just the top, so a block of odd
    # rows at the start can't decide it.
    best_part = None
    for c in range(raw.shape[1]):
        if c == name_col:
            continue
        candidates = [r for r in rows if _looks_like_part(cell(r, c))]
        if len(candidates) < max(1, len(rows) // 2):
            continue
        step = max(1, len(candidates) // _PART_SAMPLE_SIZE)
        sample = candidates[::step][:_PART_SAMPLE_SIZE]
        hits = sum(1 for r in sample
                   if _is_known_part(cell(r, c), competitor_key(cell(r, name_col)), all_dfs))
        ratio = hits / len(sample)
        if hits and ratio >= _MIN_PART_RATIO and (best_part is None or (ratio, hits) > best_part[1:]):
            best_part = (c, ratio, hits)
    if best_part is None:
        raise ValueError(
            f"Found competitor names in {_column_label(raw, name_col, header_row)}, but no "
            "column whose values match parts in those competitors' databases."
        )

    return best_part[0], name_col, first_data_row, header_row


def read_batch_file(upload, all_dfs):
    """Uploaded file -> (DataFrame with the two standard columns, detection note)."""
    data = upload.read()
    if upload.filename.lower().endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(data), header=None, dtype=str,
                          encoding="utf-8-sig", on_bad_lines="skip")
    else:
        raw = pd.read_excel(io.BytesIO(data), header=None, dtype=str)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    raw.columns = range(raw.shape[1])
    if raw.empty:
        raise ValueError("The file is empty.")

    part_col, name_col, start, header_row = _detect_batch_columns(raw, all_dfs)

    body = raw.iloc[start:]
    batch_df = pd.DataFrame({
        BATCH_PART_COL: body[part_col].fillna("").astype(str).str.strip(),
        BATCH_NAME_COL: body[name_col].fillna("").astype(str).str.strip(),
    })
    batch_df = batch_df[(batch_df[BATCH_PART_COL] != "") | (batch_df[BATCH_NAME_COL] != "")]
    batch_df = batch_df.reset_index(drop=True)

    detected = {
        "part_column": _column_label(raw, part_col, header_row),
        "competitor_column": _column_label(raw, name_col, header_row),
    }
    return batch_df, detected