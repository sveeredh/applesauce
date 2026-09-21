"""
fix_digikey_zener_columns.py — realign misplaced columns in digikey_zener.csv.

The DigiKey zener export writes two extra fields into most rows without adding
their headers, so everything from "Impedance (Max) (Zzt)" rightward is shifted
one position and the last two real fields (Tolerance, Vz) fall off the end.

A shifted row looks like this (header -> value actually stored there):

    Impedance (Max) (Zzt)     Automotive          <- qualification, not impedance
    Mounting Type             4 Ohms              <- impedance
    Operating Temperature     Surface Mount       <- mounting type
    Package / Case            -65C ~ 175C (TJ)    <- operating temperature
    Power - Max               DO-214AB, SMC       <- package / case
    Supplier Device Package   10 W                <- power
    Tolerance                 AEC-Q101            <- qualification standard
    Vz                        SMC                 <- supplier device package
    (Tolerance and Vz values are gone)

This script detects those rows, moves every value back under its correct
header, and recovers the lost Tolerance and Vz from the Detailed Description
(falling back to Description). The two orphaned qualification fields are kept
in two appended columns rather than discarded, since the grade logic uses them.

Correctly-aligned rows and fully-empty rows are left untouched.

Usage:
    python fix_digikey_zener_columns.py [input.csv] [output.csv]
    (defaults: digikey_zener.csv -> digikey_zener_fixed.csv)
"""

import re
import sys

import pandas as pd

# Header positions in the DigiKey export (0-based).
COL_LEAKAGE   = "Current - Reverse Leakage @ Vr"
COL_IMPEDANCE = "Impedance (Max) (Zzt)"
COL_MOUNTING  = "Mounting Type"
COL_TEMP      = "Operating Temperature"
COL_CASE      = "Package / Case"
COL_POWER     = "Power - Max"
COL_SUPPLIER  = "Supplier Device Package"
COL_TOLERANCE = "Tolerance"
COL_VZ        = "Voltage - Zener (Nom) (Vz)"

# Appended homes for the two fields the export has no header for.
COL_QUAL      = "Qualification"
COL_QUAL_STD  = "Qualification Standard"

# A row is already aligned when Tolerance looks like a tolerance and Vz like a voltage.
TOL_SHAPE_RE = re.compile(r'^\s*[±+]/?-?\s*\d')
VZ_SHAPE_RE  = re.compile(r'^\s*\d+(?:\.\d+)?\s*[mk]?V\s*$', re.I)

# Recovery patterns for the two lost values.
# Detailed Description reads: "Zener Diode 25 V 10 W ±5% Surface Mount SMC"
DD_VZ_RE   = re.compile(r'(\d+(?:\.\d+)?)\s*V\b', re.I)
DD_TOL_RE  = re.compile(r'([±+]/?-?\s*\d+(?:\.\d+)?\s*%)')
# Description reads: "DIODE ZENER 18V 150MW 603". The trailing guard keeps this
# from matching part numbers that use nVn notation (MM9Z5V1B, MM5Z9V1-AQ).
DESC_VZ_RE = re.compile(r'(\d+(?:\.\d+)?)\s*V(?![A-Za-z0-9])', re.I)

BLANKS = ("", "-", "nan", "none")


def _blank(series):
    return series.fillna("").str.strip().str.lower().isin(BLANKS)


def classify_rows(df):
    """Splits the frame into aligned / empty / shifted masks."""
    tol_ok = df[COL_TOLERANCE].fillna("").str.match(TOL_SHAPE_RE)
    vz_ok = df[COL_VZ].fillna("").str.match(VZ_SHAPE_RE)
    aligned = tol_ok & vz_ok

    param_cols = [COL_LEAKAGE, COL_IMPEDANCE, COL_MOUNTING, COL_TEMP,
                  COL_CASE, COL_POWER, COL_SUPPLIER, COL_TOLERANCE, COL_VZ]
    empty = pd.concat([_blank(df[c]) for c in param_cols], axis=1).all(axis=1)

    shifted = ~aligned & ~empty
    return aligned, empty, shifted


def recover_vz(detailed, description):
    """Vz string ('25 V') from the description text, or '' if not stated."""
    m = DD_VZ_RE.search(str(detailed or ""))
    if not m:
        m = DESC_VZ_RE.search(str(description or ""))
    return f"{m.group(1)} V" if m else ""


def recover_tolerance(detailed):
    """Tolerance string ('±5%') from the description text, or '' if not stated."""
    m = DD_TOL_RE.search(str(detailed or ""))
    return re.sub(r'\s+', '', m.group(1)) if m else ""


def fix_frame(df, verbose=True):
    """Returns (fixed_df, stats)."""
    df = df.copy()
    for col in (COL_QUAL, COL_QUAL_STD):
        if col not in df.columns:
            df[col] = pd.NA

    aligned, empty, shifted = classify_rows(df)
    idx = df.index[shifted]

    if verbose:
        print(f"Rows: {len(df)} total — {int(aligned.sum())} already aligned, "
              f"{int(empty.sum())} empty, {int(shifted.sum())} shifted.")

    # Snapshot the shifted rows before overwriting anything.
    old = {c: df.loc[idx, c].copy() for c in
           [COL_IMPEDANCE, COL_MOUNTING, COL_TEMP, COL_CASE,
            COL_POWER, COL_SUPPLIER, COL_TOLERANCE, COL_VZ]}

    # Slide every value one position back to its real header.
    df.loc[idx, COL_QUAL]      = old[COL_IMPEDANCE]   # Automotive / Military / -
    df.loc[idx, COL_IMPEDANCE] = old[COL_MOUNTING]    # 4 Ohms
    df.loc[idx, COL_MOUNTING]  = old[COL_TEMP]        # Surface Mount
    df.loc[idx, COL_TEMP]      = old[COL_CASE]        # -65C ~ 175C (TJ)
    df.loc[idx, COL_CASE]      = old[COL_POWER]       # DO-214AB, SMC
    df.loc[idx, COL_POWER]     = old[COL_SUPPLIER]    # 10 W
    df.loc[idx, COL_QUAL_STD]  = old[COL_TOLERANCE]   # AEC-Q101 / MIL-PRF-...
    df.loc[idx, COL_SUPPLIER]  = old[COL_VZ]          # SMC

    # Tolerance and Vz were pushed off the end — rebuild them from the text.
    detailed = df.loc[idx, "Detailed Description"] if "Detailed Description" in df.columns else pd.Series("", index=idx)
    desc = df.loc[idx, "Description"] if "Description" in df.columns else pd.Series("", index=idx)
    vz = [recover_vz(d, s) for d, s in zip(detailed.fillna(""), desc.fillna(""))]
    tol = [recover_tolerance(d) for d in detailed.fillna("")]
    df.loc[idx, COL_VZ] = pd.Series(vz, index=idx).replace("", pd.NA)
    df.loc[idx, COL_TOLERANCE] = pd.Series(tol, index=idx).replace("", pd.NA)

    stats = {
        "total": len(df),
        "aligned": int(aligned.sum()),
        "empty": int(empty.sum()),
        "shifted": int(shifted.sum()),
        "vz_recovered": int(sum(1 for v in vz if v)),
        "vz_missing": int(sum(1 for v in vz if not v)),
        "tol_recovered": int(sum(1 for t in tol if t)),
        "tol_missing": int(sum(1 for t in tol if not t)),
    }

    if verbose and stats["shifted"]:
        print(f"Recovered Vz for {stats['vz_recovered']}/{stats['shifted']} shifted rows "
              f"({stats['vz_missing']} descriptions don't state a voltage).")
        print(f"Recovered Tolerance for {stats['tol_recovered']}/{stats['shifted']} shifted rows "
              f"({stats['tol_missing']} descriptions don't state one).")

    return df, stats


def verify(df, verbose=True):
    """Sanity-checks the fixed frame; returns a dict of problem counts."""
    def bad(col, pattern):
        vals = df[col].fillna("")
        return int((~vals.str.match(pattern, case=False) & ~_blank(df[col])).sum())

    problems = {
        "impedance not Ohms":     bad(COL_IMPEDANCE, r'^\s*[\d.]+\s*[kM]?\s*Ohms?\b'),
        "mounting not a type":    bad(COL_MOUNTING, r'^\s*(Surface Mount|Through Hole|Stud Mount|Chassis|Panel|Bracket)'),
        "temp has no degrees":    bad(COL_TEMP, r'.*(°C|\(TJ\))'),
        "power not watts":        bad(COL_POWER, r'^\s*[\d.]+\s*[mk]?W\s*$'),
        "tolerance not a %":      bad(COL_TOLERANCE, r'^\s*[±+]/?-?\s*[\d.]+\s*%'),
        "vz not a voltage":       bad(COL_VZ, r'^\s*[\d.]+\s*[mk]?V\s*$'),
    }
    if verbose:
        print("\nPost-fix column checks (non-empty values that don't fit their header):")
        for k, v in problems.items():
            print(f"  {k:<24} {v}")
    return problems


def main():
    in_path = sys.argv[1] if len(sys.argv) > 1 else "digikey_zener.csv"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "digikey_zener_fixed.csv"

    df = pd.read_csv(in_path, encoding="utf-8-sig", low_memory=False, dtype=str)
    fixed, _ = fix_frame(df)
    verify(fixed)
    fixed.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\nWrote {out_path}")


if __name__ == "__main__":
    main()