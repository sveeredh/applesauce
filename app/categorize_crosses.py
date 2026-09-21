"""
categorize_crosses.py — Batch-only cross categorization export.

Reads a batch input file (same format as Mode 2 / Mode 4: an Excel or CSV file
with "Competitor Parts" and "Competitor Name" columns), runs each part through
the normal cross flow, and writes a categorized output file with these columns:

    competitorName        competitor name, upper-cased
    competitorGPN         competitor device name (DigiKey "Mfr Part #")
    competitorOPN         same as competitorGPN
    tiGPN                 TI 1st alternative's GPN
    replacementCodeGPN    always "P" on a crossed row
    tiOPN                 TI OPN generated from the GPN (_generate_ti_opn)
    replacementCodeOPN    S / Q / P per the rules below
    commentsGPN           "Zener" on a zener part, otherwise "cross tool"
    commentsOPN           "Comp SMx package" when the competitor package is an
                          SMx (SMA/SMB/SMC/...) or a DO- body, else "cross tool"
    externalCommentsGPN   "Shreya Cross Tool" on every row

replacementCodeOPN rules -- ZENER parts:
    S  package matches (standard package-matching rules), Vz matches exactly,
       AND the competitor's tolerance is >= TI's (TI must be at least as tight).
       A competitor tolerance that only just undercuts TI's counts as a match
       (4.99% against TI's 5%), and an unstated competitor tolerance is taken
       to be the looser of the two, so it does not block S on its own.
    Q  package matches, and Vz is within +/- 10%; also an otherwise-S cross
       where TI's tolerance is looser than the competitor's
    P  package does NOT match, but Vz is within +/- 10%
    otherwise -> the part is NOT crossed (all TI columns left blank)

replacementCodeOPN rules -- ESD/TVS parts:
    S  package matches AND specs are same-or-better (Vrwm within 10%, same
       direction, capacitance in the tighter band, ESD within 10%, surge at
       least 20% of the competitor's) AND the channel count and direction both
       match -- a different channel count or direction caps the cross at Q
    Q  package matches but specs are slightly worse (Vrwm within 10% and
       capacitance in the looser band)
    P  different package, with specs inside the existing cross thresholds
    otherwise -> the part is NOT crossed (all TI columns left blank)

The ESD/TVS comparison reuses applesauce._get_replacement_type, so it stays in
step with the thresholds already used elsewhere in the tool.

Additional rule: if the competitor specs carry a "Mounting Type" of "Through
Hole", we never have an alternate — all TI columns are left blank for that part.

Long batches auto-save every CHUNK_SAVE_SIZE parts to CHUNK_SAVE_DIR as
categorized_crosses_pt1.xlsx, _pt2.xlsx, ... so a crash near the end doesn't
throw the run away. Each file holds its own block of parts, in order, so
concatenating them rebuilds everything processed before the crash.

A part that raises anywhere in the cross flow is logged with its traceback and
left un-crossed; the batch carries on. One malformed cell in a competitor sheet
should not cost a multi-hour run.

Usage
-----
    python categorize_crosses.py                          # file dialogs
    python categorize_crosses.py in.xlsx out.xlsx         # explicit paths

Or from other code (databases already loaded):

    from categorize_crosses import CrossContext, categorize_batch_df
    ctx = CrossContext(all_dfs, ti_specs_df, ti_zener_specs_df)
    out_df = categorize_batch_df(batch_df, ctx)
"""

import os
import re
import sys
import time
import traceback
from collections import namedtuple

import pandas as pd
import contextlib
import io

# Every database this flow needs comes from applesauce.load_all_dfs, so the
# per-competitor file constants are deliberately NOT imported here — there is
# no second copy of the loading list to fall out of step with the dispatcher.
from applesauce import (
    get_competitor_specs_leniently,
    canonical_competitor_name,
    find_ti_alternatives,
    find_ti_zener_alternatives,
    _generate_ti_opn,
    _get_replacement_type,
    _canonical_to_ti_pkg_with_pins,
    _classify_digikey_package,
    load_all_dfs,
    manage_data_files,
)
from ti_scrape import fetch_ti_specs_from_excel
from ti_zener_scrape import fetch_ti_zener_specs_from_excel
from parsing import normalize_package, to_numeric_val

# ── Tunables ──────────────────────────────────────────────────────────────────
COLUMNS = [
    "competitorName",
    "competitorGPN",
    "competitorOPN",
    "tiGPN",
    "replacementCodeGPN",
    "tiOPN",
    "replacementCodeOPN",
    "commentsGPN",
    "commentsOPN",
    "externalCommentsGPN",
]

COMMENT_TEXT = "cross tool"          # fallback when a cell has nothing to say
EXTERNAL_COMMENT_TEXT = "cross tool"          # externalCommentsGPN, every row
ZENER_COMMENT_GPN = "Zener"                   # commentsGPN on a zener part
TVS_COMMENT_GPN = "ESD/TVS"                   # commentsGPN on everything else
PACKAGE_COMMENT_OPN = "Comp {} package"       # commentsOPN, normalized comp package

# Competitor packages that earn the SMX_COMMENT_OPN note. First pattern catches
# the SMx family (SMA, SMB, SMC, SMAF, SMBJ, ...); second catches DO- bodies
# (DO-214AC, DO-41, and the un-hyphenated DO214AB some sheets use).
SMX_PACKAGE_PATTERNS = (
    re.compile(r"\bSM[A-Z]{1,2}\b", re.IGNORECASE),
    re.compile(r"\bDO-?\d", re.IGNORECASE),
)
GPN_REPLACEMENT_CODE = "P"           # replacementCodeGPN on every crossed row
VZ_TOLERANCE_PERCENT = 30.0          # the +/- 10% band for Q and P
VZ_EXACT_ABS_TOLERANCE = 1e-6        # "exactly" — float-noise slack only, in volts

# How far under TI's tolerance a competitor's may sit and still count as equal
# (relative). At 5.0, a 4.99% competitor part still crosses S to a TI 5% part,
# while a genuinely tighter 4% part does not.
TOLERANCE_CLOSE_PERCENT = 5.0
BLANK_CELL = ""                      # what an un-crossed TI cell contains

# How many failed parts to list individually in the end-of-run summary.
MAX_FAILURES_LISTED = 25

# If True, the three comment columns stay filled even on rows with no cross.
# If False, un-crossed rows get blank comments too.
COMMENTS_ON_UNCROSSED_ROWS = True

# Value(s) in the DigiKey "Mounting Type" column that block any alternate.
THROUGH_HOLE_MARKERS = ("through hole",)

MOUNTING_TYPE_COL = "Mounting Type"

# Periodic auto-save: a long batch shouldn't be lost to a crash near the end.
# Every CHUNK_SAVE_SIZE parts, the block just finished is written to its own
# numbered file in CHUNK_SAVE_DIR (categorized_crosses_pt1.xlsx, _pt2.xlsx, ...).
#
# >>> CHANGE THIS NUMBER to save more or less often. Set it to 0 to switch the
# >>> auto-save off entirely.
CHUNK_SAVE_SIZE = 15000

# Folder the snapshots land in. Created if it doesn't exist. The full path is
# printed at the start of every batch and again on each save.
CHUNK_SAVE_DIR = r"C:\Users\15039\Downloads\Cross-Reference-Tool-main\Cross-Reference-Tool-main\batch_output"

# Values that mean "no data" in the DigiKey export.
BLANKS = ("", "-", "nan", "none")

# Batch input columns (same as the existing batch modes)
INPUT_PART_COL = "Competitor Parts"
INPUT_COMPETITOR_COL = "Competitor Name"


CrossContext = namedtuple(
    "CrossContext",
    ["all_dfs", "ti_specs_df", "ti_zener_specs_df"],
)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _blank_row(competitor_name, competitor_gpn, comments_gpn=None, comments_opn=None):
    """
    A row with the competitor identified but no TI cross.

    comments_gpn / comments_opn carry through whatever the caller already
    worked out about the part (zener, SMx package), so an un-crossed row still
    says why it is what it is. They default to the plain comment text for the
    early exits that give up before the part is identified.
    """
    if not COMMENTS_ON_UNCROSSED_ROWS:
        return {
            "competitorName": canonical_competitor_name(competitor_name),
            "competitorGPN": competitor_gpn,
            "competitorOPN": competitor_gpn,
            "tiGPN": BLANK_CELL,
            "replacementCodeGPN": BLANK_CELL,
            "tiOPN": BLANK_CELL,
            "replacementCodeOPN": BLANK_CELL,
            "commentsGPN": BLANK_CELL,
            "commentsOPN": BLANK_CELL,
            "externalCommentsGPN": BLANK_CELL,
        }

    return {
        "competitorName": canonical_competitor_name(competitor_name),
        "competitorGPN": competitor_gpn,
        "competitorOPN": competitor_gpn,
        "tiGPN": BLANK_CELL,
        "replacementCodeGPN": BLANK_CELL,
        "tiOPN": BLANK_CELL,
        "replacementCodeOPN": BLANK_CELL,
        "commentsGPN": comments_gpn or COMMENT_TEXT,
        "commentsOPN": comments_opn or COMMENT_TEXT,
        "externalCommentsGPN": EXTERNAL_COMMENT_TEXT,
    }


def _build_comments(is_zener, comp_package):
    """
    The (commentsGPN, commentsOPN) pair for a row.

    commentsGPN names the part family. commentsOPN names the competitor's
    package as normalized -- the value already written back into
    comp_specs["Package"] -- so it reads "Comp SOT-23-3 package".
    """
    gpn = ZENER_COMMENT_GPN if is_zener else TVS_COMMENT_GPN
    pkg = str(comp_package or "").strip()
    if not pkg or pkg.lower() in BLANKS:
        return gpn, COMMENT_TEXT
    return gpn, PACKAGE_COMMENT_OPN.format(pkg)

def _is_smx_package(*package_values):
    """
    True when any of the given package strings names an SMx body (SMA, SMB,
    SMC, ...) or a DO- body. Several values are accepted because the package
    can be read off the competitor sheet, the DigiKey 'Package / Case' column,
    or the canonical code, and only one of them may spell it out.
    """
    for value in package_values:
        s = str(value or "").strip()
        if s.lower() in BLANKS:
            continue
        if any(pattern.search(s) for pattern in SMX_PACKAGE_PATTERNS):
            return True
    return False


def _is_through_hole(comp_specs):
    """
    True when the competitor part is a through-hole part, read off the
    'Mounting Type' the competitor specs carry. A spec dict without that key
    is treated as surface-mount.
    """
    mounting = str((comp_specs or {}).get(MOUNTING_TYPE_COL, "") or "").strip().lower()
    return any(marker in mounting for marker in THROUGH_HOLE_MARKERS)


def _vz_within_tolerance(comp_vz, ti_vz, tolerance_percent=VZ_TOLERANCE_PERCENT):
    """
    +/- tolerance band around the competitor's Vz. Written in the same
    multiplicative form the zener matcher uses (comp * 0.90 <= ti <= comp * 1.10)
    so a part sitting exactly on the boundary is judged identically by both.
    """
    if comp_vz is None or ti_vz is None:
        return False
    if comp_vz == 0:
        return ti_vz == 0
    lo = comp_vz * (1.0 - tolerance_percent / 100.0)
    hi = comp_vz * (1.0 + tolerance_percent / 100.0)
    if comp_vz < 0:
        lo, hi = hi, lo
    return lo <= ti_vz <= hi


def _vz_exact(comp_vz, ti_vz):
    if comp_vz is None or ti_vz is None:
        return False
    return abs(ti_vz - comp_vz) <= VZ_EXACT_ABS_TOLERANCE


def _parse_tolerance(value):
    """
    Pulls the numeric percentage out of a tolerance string ('±5%' -> 5.0).
    Returns None when no tolerance is stated, so callers can distinguish
    "unknown" from "zero".
    """
    if value is None:
        return None
    s = str(value).strip()
    if s.lower() in BLANKS:
        return None
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


def tolerance_allows_s(comp_tol, ti_tol):
    """
    A TI part may only be an S cross when it is at least as tight as the
    competitor's: the competitor's tolerance must be >= TI's. A meaningfully
    looser TI part (comp_tol < ti_tol) caps the cross at Q.

    A competitor tolerance that only just undercuts TI's is treated as a match
    — 4.99% against TI's 5% is the same part in practice, not a tighter one —
    so anything inside TOLERANCE_CLOSE_PERCENT of TI's figure still allows S.

    An unstated tolerance is not a failure. A missing competitor tolerance is
    taken to mean TI's is the tighter of the two (so S stays open), and a
    missing TI tolerance leaves the cross to the Vz and package rules alone
    rather than demoting it on absent data.
    """
    if comp_tol is None or ti_tol is None:
        return True
    if comp_tol >= ti_tol:
        return True
    # Close enough counts: 4.99% against TI's 5% is float/rounding noise.
    return ti_tol <= comp_tol + 0.1


def _channel_count(value):
    """
    Channel count as an int, or None when it isn't stated. A zero count is
    treated as unstated — it only ever comes from an empty cell read as a float.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s.lower() in BLANKS:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return None
    count = int(m.group(0))
    return count or None


def _normalized_direction(value):
    """'Uni-Directional' / 'unidirectional' -> 'uni'. None when not stated."""
    s = str(value or "").strip().lower()
    if not s or s in BLANKS or "unknown" in s:
        return None
    if "uni" in s:
        return "uni"
    if "bi" in s:
        return "bi"
    return None


def channels_allow_s(comp_channels, ti_channels):
    """
    An S cross has to be a drop-in, so a different channel count rules it out.
    An unknown count on either side is not treated as a mismatch — there is
    nothing to compare, so the rest of the rules decide.
    """
    comp = _channel_count(comp_channels)
    ti = _channel_count(ti_channels)
    if comp is None or ti is None:
        return True
    return comp == ti


def direction_allows_s(comp_direction, ti_direction):
    """
    Same idea as channels_allow_s: a unidirectional part is not a drop-in for
    a bidirectional one. An unstated direction on either side doesn't demote.
    """
    comp = _normalized_direction(comp_direction)
    ti = _normalized_direction(ti_direction)
    if comp is None or ti is None:
        return True
    return comp == ti


def get_replacement_code_opn(comp_vz, ti_vz, is_package_match, comp_tol=None, ti_tol=None):
    """
    Zener rule.

    S  package match + exact Vz + competitor tolerance >= TI tolerance
    Q  package match + Vz within +/- 10% (or an S that a looser TI part caps)
    P  no package match + Vz within +/- 10%
    None -> do not cross this part
    """
    in_band = _vz_within_tolerance(comp_vz, ti_vz)
    if is_package_match:
        if _vz_exact(comp_vz, ti_vz):
            # Everything else lines up for S — the tolerance decides.
            return "S" if tolerance_allows_s(comp_tol, ti_tol) else "Q"
        if in_band:
            return "Q"
        return None
    if in_band:
        return "P"
    return None


def get_replacement_code_opn_tvs(comp_specs, ti_alt_specs, is_package_match,
                                 ti_channels=None, ti_direction=None):
    """
    ESD/TVS rule. Delegates to applesauce._get_replacement_type, which already
    encodes the spec comparison this categorization is meant to use:

      S  package matches AND specs are same-or-better -- Vrwm within 10%, same
         direction, capacitance inside the tighter 'S' band, ESD within 10%,
         and surge at least 20% of the competitor's
      Q  package matches but specs are slightly worse -- Vrwm still within 10%
         and capacitance inside the looser 'Q' band
      P  different package (or a package match that clears neither S nor Q)

    Reaching this function at all means find_ti_alternatives already returned
    the part, so its clamping voltage, capacitance, direction and Vrwm are
    inside the existing cross thresholds -- which is what P means here. A
    package-matching part that fails both S and Q also lands on P, matching the
    shared helper's behavior rather than dropping the cross entirely.

    On top of that, an S cross has to be a drop-in: a different channel count
    or a different direction caps it at Q, however well the rest of the specs
    line up. ti_channels / ti_direction let the caller pass values read
    straight off the TI table, since the spec dict doesn't always carry them.

    Returns None only when there are no TI specs to compare against.
    """
    if not ti_alt_specs:
        return None

    code = _get_replacement_type(comp_specs, ti_alt_specs, False, is_package_match)
    if code != "S":
        return code

    ti_chan = ti_channels if ti_channels is not None else ti_alt_specs.get("Channels")
    ti_dir = ti_direction if ti_direction is not None else ti_alt_specs.get("Direction")

    if not channels_allow_s(comp_specs.get("Channels"), ti_chan):
        print(
            f"    S -> Q: channel count differs "
            f"(competitor {comp_specs.get('Channels')}, TI {ti_chan})."
        )
        return "Q"
    if not direction_allows_s(comp_specs.get("Direction"), ti_dir):
        print(
            f"    S -> Q: direction differs "
            f"(competitor {comp_specs.get('Direction')}, TI {ti_dir})."
        )
        return "Q"
    return code


def _ti_channel_count(gpn, src_df):
    """Channel count for a TI GPN, read off the TI table ('Number of channels')."""
    if src_df is None or src_df.empty or not gpn or gpn == "-":
        return None
    pn_col = next((c for c in src_df.columns if "product or part number" in str(c).lower()), None)
    chan_col = next((c for c in src_df.columns if "number of channels" in str(c).lower()), None)
    if not pn_col or not chan_col:
        return None
    hits = src_df[src_df[pn_col].astype(str).str.strip() == str(gpn).strip()]
    if hits.empty:
        return None
    return _channel_count(hits.iloc[0][chan_col])


def _ti_pin_count(gpn, src_df):
    """Pin count for a TI GPN, used by the OPN generator."""
    if src_df is None or src_df.empty or not gpn or gpn == "-":
        return "-"
    pn_col = next((c for c in src_df.columns if "product or part number" in c.lower()), None)
    pin_col = next((c for c in src_df.columns if "pin count" in c.lower()), None)
    if not pn_col or not pin_col:
        return "-"
    hits = src_df[src_df[pn_col].astype(str).str.strip() == str(gpn).strip()]
    if hits.empty:
        return "-"
    return str(hits.iloc[0][pin_col])


def _alt_voltage(alt_dict, ti_alt_specs):
    """
    Vz for a zener alternative / Vrwm for a TVS alternative. Prefers the value
    the matcher already pulled off the TI row, falling back to the specs dict.
    """
    for key in ("vz", "vrw"):
        if key in alt_dict:
            val = to_numeric_val(alt_dict.get(key), default_if_error=None)
            if val is not None:
                return val
    if ti_alt_specs:
        return to_numeric_val(
            ti_alt_specs.get("Voltage - Reverse Standoff (Typ)"), default_if_error=None
        )
    return None

# ── Per-part categorization ───────────────────────────────────────────────────
def categorize_part(part_str, competitor_name, ctx):
    """Runs one competitor part through the cross flow and returns one output row."""
    part_str = str(part_str or "").strip()
    competitor_name = str(competitor_name or "").strip()

    if not part_str or part_str.lower() == "nan" or not competitor_name:
        return _blank_row(competitor_name, part_str)

    comp_specs = get_competitor_specs_leniently(
        part_str, competitor_name.lower(), ctx.all_dfs
    )

    if not comp_specs:
        print(f"  No specs found for '{part_str}' — leaving alternate blank.")
        return _blank_row(competitor_name, part_str)

    # DigiKey "Mfr Part #" as resolved by the lookup.
    competitor_gpn = str(comp_specs.get("Device Name", part_str) or part_str).strip()

    # Through-hole parts never get an alternate.
    if _is_through_hole(comp_specs):
        print(f"  '{competitor_gpn}' is Through Hole — no alternate.")
        return _blank_row(competitor_name, competitor_gpn)

    is_zener = "zener" in str(comp_specs.get("Source File", "")).lower()

    # Normalize the competitor package the same way the batch modes do.
    original_display_package = comp_specs.get("Package", "-")
    comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(
        comp_specs.get("Canonical Package") or _classify_digikey_package(original_display_package, "")
    ) or normalize_package(original_display_package)

    # Comment columns. Both notes come from the competitor part, so they hold
    # whether or not a TI alternative turns up; the Q wording is filled in later,
    # once the code is known. The package test reads the original display string,
    # since normalization rewrites SMx into a TI name.
    is_smx = _is_smx_package(original_display_package, comp_specs.get("Canonical Package"))
    comments_gpn, comments_opn = _build_comments(is_zener, comp_specs.get("Package"))
    if is_zener:
        alternatives = find_ti_zener_alternatives(comp_specs, ctx.ti_zener_specs_df)
    else:
        alternatives = find_ti_alternatives(comp_specs, ctx.ti_specs_df)

    if not alternatives:
        return _blank_row(competitor_name, competitor_gpn, comments_gpn, comments_opn)

    top_alt = alternatives[0]
    ti_gpn = str(top_alt.get("part_number", "")).strip()
    is_package_match = bool(top_alt.get("is_package_match", False))

    src_df = ctx.ti_zener_specs_df if is_zener else ctx.ti_specs_df
    if is_zener:
        ti_alt_specs = fetch_ti_zener_specs_from_excel(ti_gpn, ctx.ti_zener_specs_df, [])
    else:
        ti_alt_specs = fetch_ti_specs_from_excel(ti_gpn, ctx.ti_specs_df, [])

    if is_zener:
        comp_vz = to_numeric_val(
            comp_specs.get("Voltage - Reverse Standoff (Typ)"), default_if_error=None
        )
        ti_vz = _alt_voltage(top_alt, ti_alt_specs)

        # Tolerance: the matcher carries TI's raw value on the alternative; fall
        # back to the fetched spec sheet when it isn't there.
        comp_tol = _parse_tolerance(comp_specs.get("Tolerance"))
        ti_tol = _parse_tolerance(top_alt.get("tolerance"))
        if ti_tol is None and ti_alt_specs:
            ti_tol = _parse_tolerance(ti_alt_specs.get("Tolerance"))

        code_opn = get_replacement_code_opn(comp_vz, ti_vz, is_package_match, comp_tol, ti_tol)
        failure_detail = f"comp Vz: {comp_vz}, TI Vz: {ti_vz}"
    else:
        # ESD/TVS: the shared spec comparison decides S/Q/P, then the channel
        # count and direction get the final say on whether S stands.
        code_opn = get_replacement_code_opn_tvs(
            comp_specs,
            ti_alt_specs,
            is_package_match,
            ti_channels=_ti_channel_count(ti_gpn, src_df),
            ti_direction=top_alt.get("direction") or (ti_alt_specs or {}).get("Direction"),
        )
        failure_detail = "no TI specs found for the alternative"

    if code_opn is None:
        print(
            f"  '{competitor_gpn}' -> '{ti_gpn}' fails S/Q/P rules "
            f"(pkg match: {is_package_match}, {failure_detail}). Not crossed."
        )
        return _blank_row(competitor_name, competitor_gpn, comments_gpn, comments_opn)

    # OPN generation — same call the existing batch modes make.
    if ti_alt_specs:
        ti_package = ti_alt_specs.get("Package", "-")
        ti_opn = _generate_ti_opn(
            ti_gpn,
            ti_package,
            _ti_pin_count(ti_gpn, src_df),
            comp_specs.get("Package", ""),
            comp_specs.get("Canonical Package"),
        )
    else:
        ti_opn = ti_gpn


    return {
        "competitorName": canonical_competitor_name(competitor_name),
        "competitorGPN": competitor_gpn,
        "competitorOPN": competitor_gpn,
        "tiGPN": ti_gpn,
        "replacementCodeGPN": GPN_REPLACEMENT_CODE,
        "tiOPN": ti_opn,
        "replacementCodeOPN": code_opn,
        "commentsGPN": comments_gpn,
        "commentsOPN": comments_opn,
        "externalCommentsGPN": EXTERNAL_COMMENT_TEXT,
    }


# ── Batch driver ──────────────────────────────────────────────────────────────
def _drop_header_echo_rows(batch_df):
    """
    Drops rows that just repeat the header text (some batch files carry the
    header twice, or a second title row). Left in, they get processed as if
    they were a real part and land in the output as a junk row.
    """
    if batch_df is None or batch_df.empty:
        return batch_df

    header_tokens = {str(c).strip().lower() for c in batch_df.columns}
    header_tokens.update({INPUT_PART_COL.lower(), INPUT_COMPETITOR_COL.lower()})

    def _is_echo(row):
        part = str(row.get(INPUT_PART_COL, "")).strip().lower()
        comp = str(row.get(INPUT_COMPETITOR_COL, "")).strip().lower()
        return part in header_tokens or comp in header_tokens

    echo_mask = batch_df.apply(_is_echo, axis=1)
    dropped = int(echo_mask.sum())
    if dropped:
        print(f"Skipping {dropped} repeated-header row(s) in the input file.")
    return batch_df[~echo_mask].reset_index(drop=True)


def _save_chunk(rows, chunk_number, save_dir=CHUNK_SAVE_DIR):
    """
    Writes one block of finished rows to categorized_crosses_pt<N>.xlsx. Each
    file holds only its own block, so the parts stay small on a long run and
    concatenating pt1..ptN in order rebuilds everything processed so far.

    Best-effort: a failure here is reported but never interrupts the run, since
    the real output file is still coming.
    """
    try:
        os.makedirs(save_dir, exist_ok=True)
        path = os.path.join(save_dir, f"categorized_crosses_pt{chunk_number}.xlsx")
        pd.DataFrame(rows, columns=COLUMNS).to_excel(path, index=False, sheet_name="Categorized_Crosses")
        print(f"\n  [auto-save] {len(rows)} parts saved to '{path}'")
        return path
    except Exception as e:
        print(f"\n  [auto-save] Could not write progress file: {e}")
        return None


def categorize_batch_df(batch_df, ctx, show_progress=True, chunk_size=CHUNK_SAVE_SIZE,
                        chunk_dir=CHUNK_SAVE_DIR):
    """Runs a whole batch DataFrame and returns the categorized output DataFrame."""
    if INPUT_PART_COL not in batch_df.columns or INPUT_COMPETITOR_COL not in batch_df.columns:
        raise ValueError(
            f"Input file must have '{INPUT_PART_COL}' AND '{INPUT_COMPETITOR_COL}' columns."
        )

    batch_df = _drop_header_echo_rows(batch_df)
    iterator = batch_df.iterrows()
    if show_progress:
        try:
            from tqdm import tqdm
            iterator = tqdm(
                batch_df.iterrows(),
                desc="Categorizing crosses",
                unit="part",
                file=sys.stdout,
                total=len(batch_df),
            )
        except ImportError:
            print("Categorizing crosses... (for a progress bar, run: pip install tqdm)")

    if chunk_size:
        print(f"Auto-saving every {chunk_size} parts to '{chunk_dir}'.")

    rows = []
    failures = []         # (part, competitor, error) for the end-of-run summary
    saved_through = 0     # index in `rows` up to which a snapshot exists
    chunk_number = 1
    for _, row in iterator:
        part_str = row[INPUT_PART_COL]
        comp_str = row[INPUT_COMPETITOR_COL]

        # One bad cell in a competitor sheet must not end a multi-hour batch.
        # The part is left un-crossed, the traceback is printed so the actual
        # frame is identifiable, and the run continues.
        try:
            rows.append(categorize_part(part_str, comp_str, ctx))
        except Exception as e:
            failures.append((part_str, comp_str, f"{type(e).__name__}: {e}"))
            print(f"\n  [skipped] '{part_str}' ({comp_str}) raised {type(e).__name__}: {e}")
            traceback.print_exc()
            rows.append(_blank_row(comp_str, str(part_str or "").strip()))

        # Snapshot each block of chunk_size parts, but not on the final row —
        # the real output file is written moments later anyway.
        if (chunk_size and len(rows) - saved_through >= chunk_size
                and len(rows) < len(batch_df)):
            _save_chunk(rows[saved_through:], chunk_number, chunk_dir)
            saved_through = len(rows)
            chunk_number += 1

    if failures:
        print(f"\n{len(failures)} part(s) raised and were left un-crossed:")
        for part, comp, err in failures[:MAX_FAILURES_LISTED]:
            print(f"  {part} ({comp}) — {err}")
        if len(failures) > MAX_FAILURES_LISTED:
            print(f"  ...and {len(failures) - MAX_FAILURES_LISTED} more.")

    return pd.DataFrame(rows, columns=COLUMNS)


def _read_batch_input(input_path):
    if str(input_path).lower().endswith(".csv"):
        return pd.read_csv(input_path, on_bad_lines="skip", encoding="utf-8-sig", low_memory=False)
    return pd.read_excel(input_path)


def _write_to_path(output_df, output_path):
    """Single write attempt. Raises on failure."""
    if str(output_path).lower().endswith(".csv"):
        output_df.to_csv(output_path, index=False)
        return
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Categorized_Crosses")
        worksheet = writer.sheets["Categorized_Crosses"]
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2


def write_output(output_df, output_path):
    """
    Writes the output, auto-sizing columns for xlsx.

    A batch can take a long time, so a locked target file (open in Excel,
    mid-OneDrive-sync, read-only folder) must not throw the results away.
    On failure this retries against a timestamped sibling path, then the
    user's home directory, and returns whichever path actually took.
    """
    try:
        _write_to_path(output_df, output_path)
        return output_path
    except (PermissionError, OSError) as e:
        print(f"\nCould not write to '{output_path}': {e}")
        print("The file is most likely open in Excel. Saving to a new file instead.")

    stamp = time.strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(output_path)
    ext = ext or ".xlsx"
    fallbacks = [
        f"{base}_{stamp}{ext}",
        os.path.join(os.path.expanduser("~"), "Downloads", f"categorized_crosses_{stamp}{ext}"),
        os.path.join(os.path.expanduser("~"), f"categorized_crosses_{stamp}{ext}"),
    ]

    for path in fallbacks:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            _write_to_path(output_df, path)
            print(f"Saved to '{path}' instead.")
            return path
        except (PermissionError, OSError) as e:
            print(f"  ...also could not write '{path}': {e}")

    raise IOError(
        "Could not write the results anywhere. Close any open copies of the "
        "output file and re-run, or pass a different output path."
    )


def load_context(force_reload=False):
    """
    Loads every database this flow needs and returns a CrossContext.

    The databases and their packing order live in applesauce.load_all_dfs.
    Calling it here (rather than rebuilding the tuple) is what keeps this
    module working when a competitor is added on the applesauce side — the
    dispatcher unpacks a fixed number of entries, so a locally-built tuple
    goes stale the moment that number changes.
    """
    print("Loading databases...")
    manage_data_files(force_reload=force_reload)

    loaded = load_all_dfs(force_reload)

    print("All databases loaded and ready.")
    return CrossContext(
        loaded.all_dfs,
        loaded.ti_specs_df,
        loaded.ti_zener_specs_df,
    )


def run_batch(input_path, output_path, ctx=None):
    """End-to-end: read input, categorize, write output. Returns the output DataFrame."""
    if ctx is None:
        ctx = load_context()

    batch_df = _read_batch_input(input_path)
    output_df = categorize_batch_df(batch_df, ctx)
    written_path = write_output(output_df, output_path)

    crossed = int((output_df["tiGPN"].astype(str).str.strip() != "").sum())
    print(f"\nCategorized {len(output_df)} parts ({crossed} crossed, {len(output_df) - crossed} not crossed).")
    if crossed:
        counts = output_df.loc[output_df["tiGPN"].astype(str).str.strip() != "", "replacementCodeOPN"].value_counts()
        print("replacementCodeOPN breakdown: " + ", ".join(f"{k}: {v}" for k, v in counts.items()))
    print(f"Output saved to '{written_path}'")
    return output_df


def _prompt_for_paths():
    """File dialogs, matching the behavior of the other batch modes."""
    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    input_path = filedialog.askopenfilename(
        title="Select the Competitor Parts File",
        filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")],
    )
    if not input_path:
        return None, None

    output_path = filedialog.asksaveasfilename(
        title="Save Categorized Crosses As...",
        filetypes=[("Excel Files", "*.xlsx")],
        defaultextension=".xlsx",
        initialfile="categorized_crosses.xlsx",
    )
    return input_path, (output_path or None)


def main():
    start_time = time.time()

    if len(sys.argv) >= 3:
        input_path, output_path = sys.argv[1], sys.argv[2]
    elif "RUNNING_IN_GUI" in os.environ:
        print("Backend: GUI mode detected, reading file paths from input stream.")
        input_path = input().strip()
        output_path = input().strip()
    else:
        input_path, output_path = _prompt_for_paths()

    if not input_path:
        print("No input file selected. Aborting.")
        return
    if not output_path:
        print("No output file location selected. Aborting.")
        return
    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found at '{input_path}'")
        return

    try:
        run_batch(input_path, output_path)
    except ValueError as e:
        print(f"ERROR: {e}")
        return
    except Exception as e:
        print(f"ERROR: Could not complete the batch. {e}")
        return

    print(f"\nTotal execution time: {time.time() - start_time:.2f} seconds")


if __name__ == "__main__":
    main()