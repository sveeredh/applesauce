"""
JJM (Jiangsu JieJie Microelectronics) catalogue parser and loader.

The six workbooks all share a layout: row 1 is a merged title banner ("Zener
Diode", "Transient Voltage Suppressor", "TVS Diode Array") and row 2 carries
the real headers, so they load with header_row=1. Column A is "Product Name"
in every file and column B is "JJM Package".

Two things these files do that the normal loader can't cope with, both handled
by load_jjm_xlsx below:

1. Their styles.xml carries a fill definition openpyxl's schema rejects, so
   pandas' default engine dies with "expected <class
   'openpyxl.styles.fills.Fill'>" before a single row is read. Nothing is wrong
   with the cell data - only the formatting - so the fallback reads the sheet
   XML straight out of the zip and never looks at styles.xml at all.

2. Every "Product Name" cell has the datasheet link's text welded onto it, so
   the value arrives as "JEB03DFP<a run of spaces>DataSheet". That is stripped
   at load, so every downstream lookup sees a clean part number.

Three sheet shapes, one parser each:
  Zener  Product Name | JJM Package | Status | Tolerance | PD_Max (W) |
         VZ_Min (V) | VZ_Max (V) | @ IZT (mA) | ZZT_Max | ZZK_Max | IZK |
         IR_Max | @ VR (V)
  TVS    Product Name | JJM Package | Uni-Polar / Bi-Polar | Status |
         VR_Typ (V) | IR_Max | @ VR | VBR_Min | VBR_Max | @ IT |
         PPP_Max (10/1000us)(W) | VC_Max/@IPP for each waveform
  ESD    Product Name | JJM Package | Polarity | Status | IR_Max |
         @ VRWM (V) | VBR_Min | VESD_Max Contact-type (kV) |
         PPP_Max/VC_Typ/VC_Max/@IPP (1.2/50us&8/20us @ 2ohm) |
         CJ_Typ (pF) | CJ_Max (pF)

applesauce.py imports both entry points from here:
    from jjm_scrape import fetch_jjm_specs_from_excel, load_jjm_xlsx
"""

import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

from parsing import safe_strip


# ==========================================================================
# Loading
# ==========================================================================

_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# Trailing link text on a part-number cell: two or more spaces (Excel pads the
# gap) followed by "datasheet"/"data sheet" and nothing else. Anchored to the
# end so a part number that legitimately contains the word is untouched.
DATASHEET_SUFFIX_RE = re.compile(r"\s{2,}data\s*sheet\s*$", re.IGNORECASE)


def clean_jjm_cell(value):
    """
    Strips the welded-on datasheet link text and collapses whitespace.
    "JEB03DFP                    DataSheet" -> "JEB03DFP"
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return value
    if not isinstance(value, str):
        return value
    cleaned = DATASHEET_SUFFIX_RE.sub("", value)
    # Some rows pad without the link text; some use a non-breaking space.
    cleaned = cleaned.replace("\xa0", " ")
    return re.sub(r"\s+", " ", cleaned).strip()


def _col_index(cell_ref):
    """'C12' -> 2. Column letters only; the row number is ignored."""
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(t.text or "" for t in si.iter(f"{{{_NS['m']}}}t")) for si in root]


def _first_sheet_path(zf):
    """
    Resolves the first sheet through workbook.xml.rels rather than assuming
    sheet1.xml, which is not always the first sheet in tab order.
    """
    try:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        first = wb.find("m:sheets/m:sheet", _NS)
        rid = first.get(f"{rel_ns}id")
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target").lstrip("/")
                return target if target.startswith("xl/") else f"xl/{target}"
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


def read_xlsx_raw(path, header_row=0):
    """
    Reads a worksheet straight out of the xlsx zip, touching only
    sharedStrings.xml and the sheet XML. Styles, themes and charts are never
    parsed, so a malformed styles.xml cannot stop the load.
    """
    with zipfile.ZipFile(path) as zf:
        strings = _shared_strings(zf)
        root = ET.fromstring(zf.read(_first_sheet_path(zf)))

    rows = []
    width = 0
    for row_el in root.iter(f"{{{_NS['m']}}}row"):
        cells = {}
        for c in row_el.findall("m:c", _NS):
            ref = c.get("r") or ""
            idx = _col_index(ref) if ref else len(cells)
            ctype = c.get("t")
            if ctype == "inlineStr":
                is_el = c.find("m:is", _NS)
                value = "".join(t.text or "" for t in is_el.iter(f"{{{_NS['m']}}}t")) if is_el is not None else None
            else:
                v = c.find("m:v", _NS)
                if v is None or v.text is None:
                    value = None
                elif ctype == "s":
                    si = int(v.text)
                    value = strings[si] if 0 <= si < len(strings) else None
                elif ctype in ("str", "e"):
                    value = v.text
                else:
                    try:
                        value = float(v.text)
                        if value.is_integer():
                            value = int(value)
                    except (ValueError, TypeError):
                        value = v.text
            if value is not None:
                cells[idx] = value
        width = max(width, (max(cells) + 1) if cells else 0)
        rows.append(cells)

    table = [[row.get(i) for i in range(width)] for row in rows]
    if header_row is None:
        return pd.DataFrame(table)
    if len(table) <= header_row:
        return pd.DataFrame()

    header = [
        str(h).strip() if h is not None else f"Unnamed: {i}"
        for i, h in enumerate(table[header_row])
    ]
    return pd.DataFrame(table[header_row + 1:], columns=header)


def load_jjm_xlsx(path, header_row=1):
    """
    Normal pandas read first; the raw reader only when that fails. Either way
    the part-number column is cleaned before the frame is handed back, so the
    cached pickle holds clean values and the cleanup cost is paid once.
    """
    try:
        df = pd.read_excel(path, header=header_row)
    except Exception as exc:
        print(f"  '{path}': standard read failed ({exc}); falling back to raw XML read.")
        df = read_xlsx_raw(path, header_row=header_row)

    if df is None or df.empty:
        return pd.DataFrame()

    df.columns = [str(c).replace("\xa0", " ").strip() for c in df.columns]
    # pandas 2 gives text columns dtype object, pandas 3 gives them str, so both
    # are checked - a dtype == object test alone silently skips every column on
    # pandas 3 and the cleanup never runs.
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].map(clean_jjm_cell)

    # Drop the repeated banner/blank rows these books scatter between families.
    part_col = next((c for c in df.columns if "product name" in str(c).lower()), None)
    if part_col:
        df = df[df[part_col].notna() & (df[part_col].astype(str).str.strip() != "")]
        df = df[df[part_col].astype(str).str.lower() != "product name"]

    print(f"Loaded {len(df)} rows from '{path}'.")
    return df.reset_index(drop=True)


# ==========================================================================
# Parsing
# ==========================================================================

# JJM quotes a clamping voltage per surge waveform. Only one can go in
# "Voltage - Clamping (Max) @ Ipp", and the tool compares it against TI's own
# clamping figure, which is quoted at 8/20us -- so the 8/20 columns are tried
# first and the long 10/1000 and 10/700 waveforms are the fallback. Reorder
# this tuple to compare on a different waveform.
TVS_CLAMP_KEYWORD_PRIORITY = (
    "vc_max (8/20",
    "vc_max (1.2/50",
    "vc_max (10/1000",
    "vc_max (10/700",
)

# Same idea for the surge current that fills IEC 61000-4-5.
TVS_IPP_KEYWORD_PRIORITY = (
    "@ ipp (8/20",
    "@ ipp (1.2/50",
    "@ ipp (10/1000",
)

# Package names that mean a leaded body. Parts in these get a "Mounting Type"
# of Through Hole, which categorize_crosses reads to leave the row un-crossed.
# The DO- family splits both ways: DO-15/27/35/41/201/204 are axial, but
# DO-214/215/216/218/219 are the surface-mount SMA/SMB/SMC bodies, so those are
# excluded rather than swept up by a bare "^DO" prefix.
THROUGH_HOLE_PKG_RE = re.compile(
    r"^(?:DO-?(?!21[45689])\d|SOD-?6[47]|A-?405|R-?6|P600|GP)", re.IGNORECASE
)


def _norm_header(text):
    """Lowercased, whitespace-collapsed header text for keyword matching."""
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def _col_value(part_row, *keywords):
    """
    First value whose column header contains one of the keywords, tried in the
    order given. Unlike a per-column scan this respects keyword priority, which
    these sheets need: several columns start "VC_Max (", and which one is wanted
    depends on the waveform, not on which sits leftmost.
    """
    headers = {col: _norm_header(col) for col in part_row.index}
    for keyword in keywords:
        needle = _norm_header(keyword)
        for col, header in headers.items():
            if needle in header:
                value = part_row[col]
                if pd.notna(value) and safe_strip(value) != "":
                    return value
    return None


def _first_number(value):
    """Leading numeric value out of a cell, or None. '±5%' -> 5.0."""
    if value is None or pd.isna(value):
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", str(value))
    if not match:
        return None
    try:
        return float(match.group(1))
    except (ValueError, TypeError):
        return None


def _package_and_mounting(part_row, specs_result):
    """
    JJM package names are plain JEDEC spellings (SOD-123FL, DFN1006-2L, DO-27),
    so they go through the generic normalize_package unchanged -- no structured
    dict and no JJM-specific normalizer needed.
    """
    pkg = safe_strip(_col_value(part_row, "JJM Package", "Package"))
    specs_result["Package"] = pkg if pkg else "-"
    if pkg and THROUGH_HOLE_PKG_RE.match(pkg):
        specs_result["Mounting Type"] = "Through Hole"


def _direction_from_polarity(part_row):
    """
    'Uni-Polar' / 'Bi-Polar' in the TVS sheet, 'Polarity' in the ESD sheet.
    Checked for "bi" first: "uni" is not a substring of "bi-polar", but the
    reverse ordering has bitten this codebase before.
    """
    text = safe_strip(_col_value(part_row, "Uni-Polar / Bi-Polar", "Polarity")).lower()
    if "bi" in text:
        return "Bidirectional"
    if "uni" in text:
        return "Unidirectional"
    return "-"


def _parse_zener_row(part_row, found_df_name, part_number):
    """
    JJM states VZ_Min and VZ_Max but no nominal, so Vz is taken as the midpoint
    of the band -- for BZD27C10 that is (9.5 + 10.5) / 2 = 10 V, which is the
    nominal the part number advertises.
    """
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": "Automotive" if "auto" in found_df_name.lower() else "Non-Automotive",
        "Voltage - Reverse Standoff (Typ)": "-",   # holds Vz for the zener flow
        "Tolerance": "-",
        "Power Dissipation (Pd)": "-",
        "Leakage Current (Ir)": "-",
        "Capacitance": "-",
        "Package": "-",
        "Price ($/ku)": "-",
    }

    _package_and_mounting(part_row, specs_result)

    vz_min = _first_number(_col_value(part_row, "VZ_Min"))
    vz_max = _first_number(_col_value(part_row, "VZ_Max"))
    if vz_min is not None and vz_max is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{(vz_min + vz_max) / 2:g} V"
    elif vz_min is not None or vz_max is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{(vz_min or vz_max):g} V"

    # Tolerance is already written as "±5%". Anything below 1 without a percent
    # sign is a fraction (0.05 = 5%); a bare "5" still reads as 5%.
    tol_raw = safe_strip(_col_value(part_row, "Tolerance"))
    tol_val = _first_number(tol_raw)
    if tol_val is not None:
        if "%" not in tol_raw and tol_val < 1:
            tol_val *= 100
        specs_result["Tolerance"] = f"±{tol_val:g}%"
    elif vz_min is not None and vz_max is not None and (vz_min + vz_max) > 0:
        # No tolerance column value: derive it from the band itself.
        nominal = (vz_min + vz_max) / 2
        specs_result["Tolerance"] = f"±{(vz_max - nominal) / nominal * 100:g}%"

    pd_max = _first_number(_col_value(part_row, "PD_Max"))
    if pd_max is not None:
        specs_result["Power Dissipation (Pd)"] = f"{pd_max:g} W"

    ir_max = _first_number(_col_value(part_row, "IR_Max"))
    if ir_max is not None:
        specs_result["Leakage Current (Ir)"] = f"{ir_max:g} µA"

    return specs_result


def _parse_tvs_row(part_row, found_df_name, part_number):
    """Row from jjm_tvs_specs / jjm_auto_tvs_specs (the through-hole TVS book)."""
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": "Automotive" if "auto" in found_df_name.lower() else "Non-Automotive",
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",       # not stated in the TVS book
        "Capacitance": "-",         # not stated in the TVS book
        "Channels": "1",            # single-line parts; the book has no channel column
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    _package_and_mounting(part_row, specs_result)
    specs_result["Direction"] = _direction_from_polarity(part_row)

    channel_count = jjm_channel_count(
        part_number,
        specs_result["Package"]
    )

    if channel_count is not None:
        specs_result["Channels"] = str(channel_count)

    vr_typ = _first_number(_col_value(part_row, "VR_Typ"))
    if vr_typ is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vr_typ:g} V"

    clamp = _first_number(_col_value(part_row, *TVS_CLAMP_KEYWORD_PRIORITY))
    if clamp is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamp:g} V"

    ipp = _first_number(_col_value(part_row, *TVS_IPP_KEYWORD_PRIORITY))
    if ipp is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp:g} A (8/20µs)"

    ppp = _first_number(_col_value(part_row, "PPP_Max"))
    if ppp is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp:g} W"

    return specs_result

def jjm_channel_count(part_number, package):
    pn = str(part_number or "").upper().strip()
    pkg = str(package or "").upper().strip()
    pkg_flat = re.sub(r"[\s\-_]+", "", pkg)

    # ========================================================
    # EXPLICIT CHANNEL COUNT IN PART NUMBER
    # ========================================================

    # SRVxx-4x
    # SRV05-4U -> 4
    # SRV03-4U -> 4
    m = re.match(r"^SRV[A-Z0-9]*-(\d+)[A-Z]", pn)
    if m:
        return int(m.group(1))

    # JEUxx-4RDA style
    # Explicit 4-line array designation
    m = re.search(r"-(\d+)RDA", pn)
    if m:
        return int(m.group(1))

    # ========================================================
    # VERIFIED JEU FAMILIES
    # ========================================================

    # 2-channel
    if pn.startswith("JEU0522P"):
        return 2

    if pn.startswith("JEB3312T"):
        return 2

    # 4-channel
    if pn.startswith((
        "JEU0524P",
        "JEU0324P",
        "JEU3324P",
        "JEU3324C",
        "JEU03SC",
        "JEU2574N",
    )):
        return 4

    # 5-channel
    if pn.startswith("JEU05MFC"):
        return 5

    # ========================================================
    # OBVIOUS SINGLE-LINE DISCRETES
    # ========================================================

    if any(x in pkg_flat for x in (
        "DFN06032L",
        "DFN10062L",
        "DFN16102L",
        "SOD323",
        "SOD523",
    )):
        return 1

    # Unknown rather than making a bad assumption
    return None
   
def _parse_esd_row(part_row, found_df_name, part_number):
    """Row from jjm_esd_specs / jjm_auto_esd_specs (the TVS diode array book)."""
    specs_result = {
        "Device Name": str(part_number).upper(),
        "Source File": found_df_name,
        "Grade": "Automotive" if "auto" in found_df_name.lower() else "Non-Automotive",
        "Package": "-",
        "Direction": "-",
        "Voltage - Reverse Standoff (Typ)": "-",
        "Voltage - Clamping (Max) @ Ipp": "-",
        "IEC 61000-4-5": "-",
        "IEC 61000-4-2": "-",
        "Capacitance": "-",
        "Channels": "-",
        "Power Dissipation (Pd)": "-",
        "Price ($/ku)": "-",
    }

    _package_and_mounting(part_row, specs_result)
    specs_result["Direction"] = _direction_from_polarity(part_row)
    channel_count = jjm_channel_count(
        part_number,
        specs_result["Package"]
    )

    if channel_count is not None:
        specs_result["Channels"] = str(channel_count)

    vrwm = _first_number(_col_value(part_row, "@ VRWM"))
    if vrwm is not None:
        specs_result["Voltage - Reverse Standoff (Typ)"] = f"{vrwm:g} V"

    # VC_Max is blank for some rows, so VC_Typ backs it up.
    clamp = _first_number(_col_value(part_row, "VC_Max", "VC_Typ"))
    if clamp is not None:
        specs_result["Voltage - Clamping (Max) @ Ipp"] = f"{clamp:g} V"

    esd = _first_number(_col_value(part_row, "VESD_Max"))
    if esd is not None:
        specs_result["IEC 61000-4-2"] = f"±{esd:g} kV"

    ipp = _first_number(_col_value(part_row, "@ IPP"))
    if ipp is not None:
        specs_result["IEC 61000-4-5"] = f"{ipp:g} A (8/20µs)"

    # CJ_Typ is what a designer picks the part on; CJ_Max backs it up.
    cap = _first_number(_col_value(part_row, "CJ_Typ", "CJ_Max"))
    if cap is not None:
        specs_result["Capacitance"] = f"{cap:.2f} pF"

    ppp = _first_number(_col_value(part_row, "PPP_Max"))
    if ppp is not None:
        specs_result["Power Dissipation (Pd)"] = f"{ppp:g} W"

    return specs_result


# Search order. Zener books first, then automotive TVS/ESD, then commercial --
# a part appearing in both an automotive and a commercial book is reported as
# the automotive one.
JJM_SEARCH_PRIORITY = ("Auto_Zener", "Zener", "Auto_TVS", "Auto_ESD", "TVS", "ESD")

JJM_PART_COL_KEYWORDS = ("product name", "part number", "device")


def _find_part_col(df):
    for keyword in JJM_PART_COL_KEYWORDS:
        for col in df.columns:
            if keyword in _norm_header(col):
                return col
    return None


def fetch_jjm_specs_from_excel(part_number, jjm_dfs):
    """
    Finds the part across the JJM books and dispatches to the parser for the
    book it turned up in.
    """
    if part_number is None or not jjm_dfs:
        return None

    part_row = None
    found_df_name = ""

    all_df_keys = list(jjm_dfs.keys())
    ordered_keys = [k for k in JJM_SEARCH_PRIORITY if k in all_df_keys]
    ordered_keys.extend([k for k in all_df_keys if k not in ordered_keys])

    target = clean_jjm_cell(str(part_number)).strip().lower()

    for df_name in ordered_keys:
        df = jjm_dfs[df_name]
        if df is None or df.empty:
            continue

        part_col = _find_part_col(df)
        if not part_col:
            continue

        hits = df[df[part_col].astype(str).str.strip().str.lower() == target]
        if not hits.empty:
            part_row = hits.iloc[0]
            found_df_name = df_name
            break

    if part_row is None:
        print(f"Part '{part_number}' not found in any JJM database.")
        return None

    print(f"Found specs for Part '{part_number}' in JJM database '{found_df_name}'.")

    # The sheet's own Product Name is already cleaned at load, so it is the
    # better display name than whatever was typed at the prompt.
    display_name = safe_strip(_col_value(part_row, "Product Name"))
    if not display_name or display_name == "-":
        display_name = clean_jjm_cell(str(part_number)).strip()

    name_lower = found_df_name.lower()
    if "zener" in name_lower:
        return _parse_zener_row(part_row, found_df_name, display_name)
    if "esd" in name_lower:
        return _parse_esd_row(part_row, found_df_name, display_name)
    return _parse_tvs_row(part_row, found_df_name, display_name)