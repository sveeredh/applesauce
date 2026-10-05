"""
multi_value.py  -  turn a competitor spec cell into ONE number
==============================================================
Some competitors (Leshan and others) put several values in one cell when a
part number covers a family, e.g.

    VRWM  "15,24"      -> 24
    VBR   "17.1,25.4"  -> 25.4

Rule: when a cell holds a comma-separated list, take the HIGHEST value.

Usage
-----
    from multi_value import to_num, max_of_list_columns

    v = to_num("17.1,25.4")            # 25.4
    df = max_of_list_columns(df)       # fixes every column that has lists
    df = max_of_list_columns(df, ["Vrwm", "Vbr", "Vc"])   # or only these
"""

import re

import pandas as pd

_NUM = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
# "1,000" / "12,500,000" = a thousands separator, NOT a list of values
_THOUSANDS = re.compile(r"^\s*[-+]?\d{1,3}(,\d{3})+(\.\d+)?\s*[A-Za-z%µ]*\s*$")
_SPLIT = re.compile(r"[,;/]")


def to_num(v):
    """Number from a spec cell. Lists ('15,24', '17.1, 25.4V') -> the max.
    Units are ignored. Blank / no number -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if pd.isna(v) else float(v)
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "-", "n/a", "na"):
        return None
    if _THOUSANDS.match(s):                       # 1,000 -> 1000
        s = s.replace(",", "")
    vals = []
    for part in _SPLIT.split(s):
        m = _NUM.search(part)
        if m:
            vals.append(float(m.group()))
    return max(vals) if vals else None


def has_list(v) -> bool:
    s = str(v)
    return bool(_SPLIT.search(s)) and not _THOUSANDS.match(s) and len(_NUM.findall(s)) > 1


def _fmt(x):
    """24.0 -> '24', 25.4 -> '25.4' (text, so string-based parsers downstream still work)."""
    return "" if x is None else (str(int(x)) if x == int(x) else repr(x))


def max_of_list_columns(df: pd.DataFrame, cols=None, verbose=True, as_text=True,
                        label="") -> pd.DataFrame:
    """Replace list cells with their max. Only touches cells that actually
    contain a list, so text columns (part numbers, packages, URLs) are left
    alone unless you name them in `cols`.
    as_text=True writes the max back as text ("24"), matching the rest of the
    column, so code that parses these cells as strings keeps working."""
    df = df.copy()
    targets = cols if cols is not None else list(df.columns)
    for c in targets:
        if c not in df.columns or pd.api.types.is_numeric_dtype(df[c]):
            continue                      # already numbers -> nothing to split
        mask = df[c].map(has_list)
        if cols is None:
            # auto mode: only fix columns that are numeric apart from the lists
            rest = df.loc[~mask, c].dropna().astype(str).str.strip()
            rest = rest[rest != ""]
            if len(rest) and rest.map(lambda x: to_num(x) is None).mean() > 0.2:
                continue
        n = int(mask.sum())
        if n:
            df[c] = df[c].astype(object)          # let numbers sit in a text column
            conv = (lambda v: _fmt(to_num(v))) if as_text else to_num
            df.loc[mask, c] = df.loc[mask, c].map(conv)
            if verbose:
                print(f"  {label + ': ' if label else ''}{c}: {n:,} multi-value cells -> took the max")
    return df


def max_of_lists_everywhere(obj, label=""):
    """Apply max_of_list_columns to a DataFrame, or to every DataFrame inside a
    dict / tuple / list (the shapes all_dfs uses). Returns the same shape."""
    if isinstance(obj, pd.DataFrame):
        return max_of_list_columns(obj, label=label) if not obj.empty else obj
    if isinstance(obj, dict):
        return {k: max_of_lists_everywhere(v, f"{label} {k}".strip()) for k, v in obj.items()}
    if isinstance(obj, (tuple, list)):
        return type(obj)(max_of_lists_everywhere(v, label) for v in obj)
    return obj


if __name__ == "__main__":
    tests = {"15,24": 24, "17.1,25.4": 25.4, "0.2": 0.2, "17": 17, "15V, 24V": 24,
             "1,000": 1000, "3.3/5": 5, "": None, None: None, 12: 12.0}
    for k, want in tests.items():
        got = to_num(k)
        print(f"{str(k)!r:14} -> {got!r:8} {'ok' if got == want else '<-- WRONG'}")
    demo = pd.DataFrame({"Part": ["S-LR1LINT", "X"], "Pkg": ["SOD-323", "SOD-323"],
                         "Vrwm": ["15,24", "5"], "Vbr": ["17.1,25.4", "6.4"]})
    print(max_of_list_columns(demo))