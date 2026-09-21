"""
api.py -- Flask backend for the Cross Reference Tool frontend.
Run with: python api.py

Database loading goes through applesauce.load_all_dfs(), the same call main()
and categorize_crosses.load_context() make. This file deliberately keeps NO
per-competitor file constants and builds NO all_dfs tuple of its own: that
tuple is positional and get_competitor_specs_leniently unpacks it by position,
so a second copy here drifts out of step the moment a competitor is added, and
every lookup after the drift silently reads the wrong database.
"""

import io
import os
import re
import threading
import uuid

import pandas as pd
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

from applesauce import (
    get_competitor_specs_leniently,
    _COMPETITOR_NAME_ALIASES,
    find_ti_alternatives,
    find_ti_zener_alternatives,
    _generate_ti_opn,
    _get_replacement_type,
    _canonical_to_ti_pkg,
    _canonical_to_ti_pkg_with_pins,
    _classify_digikey_package,
    _classify_ti_package,
    canonical_competitor_name,
    PACKAGE_DISPLAY_MAP,
    load_all_dfs,
    manage_data_files,
)
from ti_scrape import fetch_ti_specs_from_excel
from ti_zener_scrape import fetch_ti_zener_specs_from_excel
from parsing import normalize_package

app = Flask(__name__)
CORS(app)

# The frontend is served from here rather than opened as a file:// page, so it
# is always same-origin with the API and there is no second copy of index.html
# to go stale. Caching is off: an edited page should show on refresh, not on
# whenever the browser feels like revalidating.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@app.after_request
def _no_cache(response):
    if request.path == "/" or request.path.endswith(".html"):
        response.headers["Cache-Control"] = "no-store, max-age=0"
    return response


@app.route("/", methods=["GET"])
def index():
    return send_from_directory(BASE_DIR, "index.html")


jobs = {}

# -- Load databases ----------------------------------------------------------
manage_data_files(force_reload=False)
_data = load_all_dfs(force_reload=False)
all_dfs = _data.all_dfs
ti_specs_df = _data.ti_specs_df
ti_zener_specs_df = _data.ti_zener_specs_df

print("All databases loaded. API ready.")

# -- Param keys --------------------------------------------------------------
PARAM_KEYS_TVS = [
    ("Device Name",                      "Device Name"),
    ("Direction",                        "Direction"),
    ("Package",                          "Package"),
    ("Voltage - Reverse Standoff (Typ)", "Vrwm (V)"),
    ("Voltage - Clamping (Max) @ Ipp",   "Vcl @ Ipp (Max)"),
    ("Capacitance",                      "Capacitance"),
    ("Channels",                         "Channels"),
    ("IEC 61000-4-5",                    "IEC 61000-4-5"),
    ("IEC 61000-4-2",                    "IEC 61000-4-2"),
    ("Power Dissipation (Pd)",           "Peak Pulse Power"),
    ("Price ($/ku)",                     "Price ($/ku)"),
]
PARAM_KEYS_ZENER = [
    ("Device Name",                      "Device Name"),
    ("Package",                          "Package"),
    ("Voltage - Reverse Standoff (Typ)", "Vz (V)"),
    ("Tolerance",                        "Tolerance"),
    ("Power Dissipation (Pd)",           "Power (Pd)"),
    ("Price ($/ku)",                     "Price ($/ku)"),
]

# Batch input columns, matching applesauce modes 2 and 4 and the output of
# build_batch_input.py.
BATCH_PART_COL = "Competitor Parts"
BATCH_NAME_COL = "Competitor Name"

_W_VALUE_RE = re.compile(r'^\s*([\d.]+)\s*W?\s*$', re.IGNORECASE)


# -- Helpers -----------------------------------------------------------------
def _w_to_mw(val_str):
    """Convert a TI power value in W to a mW string."""
    if not val_str or val_str == "-":
        return val_str
    match = _W_VALUE_RE.match(str(val_str))
    if match:
        try:
            return f"{float(match.group(1)) * 1000:g} mW"
        except ValueError:
            pass
    return val_str


# -- Shared cross logic ------------------------------------------------------
def _run_cross(part, competitor):
    """
    One competitor part -> (comp_specs, param_keys, alt_specs_list, codes).
    Returns (None, None, [], []) when the part is not in any database.
    """
    comp_specs = get_competitor_specs_leniently(part, competitor, all_dfs)
    if not comp_specs:
        return None, None, [], []

    is_zener = "zener" in str(comp_specs.get("Source File", "")).lower()
    original_pkg = comp_specs.get("Package", "-")
    comp_specs["Package"] = _canonical_to_ti_pkg_with_pins(
        comp_specs.get("Canonical Package") or _classify_digikey_package(original_pkg, "")
    ) or normalize_package(original_pkg)

    alts = (find_ti_zener_alternatives(comp_specs, ti_zener_specs_df) if is_zener
            else find_ti_alternatives(comp_specs, ti_specs_df))

    param_keys = PARAM_KEYS_ZENER if is_zener else PARAM_KEYS_TVS
    src_df = ti_zener_specs_df if is_zener else ti_specs_df

    pn_col = pin_col = ppp_col = None
    if src_df is not None and not src_df.empty:
        pn_col = next((c for c in src_df.columns if "product or part number" in c.lower()), None)
        pin_col = next((c for c in src_df.columns if "pin count" in c.lower()), None)
        ppp_col = next((c for c in src_df.columns
                        if "peak pulse" in c.lower() or "pd (max)" in c.lower()), None)

    alt_specs_list = []
    alt_codes = []
    for alt in alts[:3]:
        gpn = alt["part_number"]
        specs = (fetch_ti_zener_specs_from_excel(gpn, ti_zener_specs_df, param_keys) if is_zener
                 else fetch_ti_specs_from_excel(gpn, ti_specs_df, param_keys))
        if not specs:
            continue

        ti_pin_str = "-"
        if pn_col:
            match_row = src_df[src_df[pn_col].astype(str).str.strip() == gpn.strip()]
            if not match_row.empty:
                if pin_col:
                    ti_pin_str = str(match_row.iloc[0][pin_col])
                if ppp_col:
                    ppp_val = str(match_row.iloc[0][ppp_col]).strip()
                    if ppp_val not in ("-", "nan", ""):
                        specs["Power Dissipation (Pd)"] = _w_to_mw(ppp_val)

        specs["OPN"] = _generate_ti_opn(
            gpn, specs.get("Package", "-"), ti_pin_str,
            comp_specs.get("Package", ""), comp_specs.get("Canonical Package"))

        # A TI row can list several bodies; show the one the competitor matched.
        pkg_str = specs.get("Package", "-")
        if pkg_str != "-":
            pkg_list = [p.strip() for p in pkg_str.split(",") if p.strip()]
            comp_canonical = comp_specs.get("Canonical Package") or next(
                (_classify_digikey_package(normalize_package(p.strip()), "")
                 for p in original_pkg.split(",")
                 if _classify_digikey_package(normalize_package(p.strip()), "")), None)
            chosen = next((p for p in pkg_list
                           if _classify_ti_package(p, None) == comp_canonical), None) or pkg_list[0]
            specs["Package"] = PACKAGE_DISPLAY_MAP.get(normalize_package(chosen), chosen)

        alt_specs_list.append(specs)
        alt_codes.append(
            _get_replacement_type(comp_specs, specs, is_zener,
                                  alt.get("is_package_match", False)) or "-"
        )

    comp_specs["Package"] = _canonical_to_ti_pkg(comp_specs.get("Canonical Package")) or original_pkg
    return comp_specs, param_keys, alt_specs_list, alt_codes


def _alt_names(alt_specs_list):
    return [s.get("OPN", s.get("Device Name", "-")) for s in alt_specs_list]


# -- Endpoints ---------------------------------------------------------------
@app.route("/competitors", methods=["GET"])
def competitors():
    """
    The competitor picker's source list. Each entry carries the display name,
    the lookup token to send back on /cross and /specs, and every alias the
    dispatcher accepts so the frontend can match on "lf" or "littelfuse" and
    still show "Littelfuse Inc".
    """
    return jsonify({
        "competitors": [
            {"name": canonical, "key": aliases[0], "aliases": list(aliases)}
            for canonical, aliases in _COMPETITOR_NAME_ALIASES.items()
        ]
    })


@app.route("/cross", methods=["POST"])
def cross():
    data = request.json or {}
    part = (data.get("part") or "").strip()
    competitor = (data.get("competitor") or "").strip().lower()
    if not part or not competitor:
        return jsonify({"error": "part and competitor are required"}), 400

    comp_specs, _, alt_specs_list, alt_codes = _run_cross(part, competitor)
    if comp_specs is None:
        return jsonify({"error": f"Could not find specs for '{part}'"}), 404

    return jsonify({
        "competitor_part": part,
        "competitor_name": canonical_competitor_name(competitor),
        "alternatives": _alt_names(alt_specs_list),
        "replacement_codes": alt_codes,
    })


@app.route("/specs", methods=["POST"])
def specs():
    data = request.json or {}
    part = (data.get("part") or "").strip()
    competitor = (data.get("competitor") or "").strip().lower()
    if not part or not competitor:
        return jsonify({"error": "part and competitor are required"}), 400

    comp_specs, param_keys, alt_specs_list, alt_codes = _run_cross(part, competitor)
    if comp_specs is None:
        return jsonify({"error": f"Could not find specs for '{part}'"}), 404

    rows = []
    for key, label in param_keys:
        comp_val = comp_specs.get(key, "-")
        if key == "Device Name":
            comp_val = str(comp_val).upper()
        row = {"label": label, "comp": comp_val}
        for i, alt in enumerate(alt_specs_list):
            row[f"alt{i + 1}"] = (alt.get("OPN", alt.get("Device Name", "-"))
                                  if key == "Device Name" else alt.get(key, "-"))
        for i in range(len(alt_specs_list), 3):
            row[f"alt{i + 1}"] = "-"
        rows.append(row)

    code_row = {"label": "Replacement Code", "comp": "-"}
    for i in range(3):
        code_row[f"alt{i + 1}"] = alt_codes[i] if i < len(alt_codes) else "-"
    rows.append(code_row)

    return jsonify({
        "competitor_part": part,
        "competitor_name": canonical_competitor_name(competitor),
        "rows": rows,
        "alt_names": _alt_names(alt_specs_list),
    })


def _process_batch_job(job_id, df):
    CHUNK_SIZE = 2000  # Change to 10000 for production
    chunk_results = []
    chunk_index = 1

    for index, row in df.iterrows():
        part = str(row[BATCH_PART_COL]).strip()
        competitor = str(row[BATCH_NAME_COL]).strip().lower()
        result = {
            "part": part,
            "competitor": canonical_competitor_name(competitor),
            "alt1": "-", "alt2": "-", "alt3": "-",
            "code1": "-", "code2": "-", "code3": "-",
        }
        if part and part.lower() != "nan" and competitor and competitor != "nan":
            try:
                _, _, alt_specs_list, alt_codes = _run_cross(part, competitor)
            except Exception as exc:
                # One bad part must not take the whole job down.
                print(f"[Batch {job_id}] '{part}' ({competitor}) failed: {exc}")
                alt_specs_list, alt_codes = [], []
            opns = _alt_names(alt_specs_list) + ["-", "-", "-"]
            codes = list(alt_codes) + ["-", "-", "-"]
            result.update({
                "alt1": opns[0], "alt2": opns[1], "alt3": opns[2],
                "code1": codes[0], "code2": codes[1], "code3": codes[2],
            })

        chunk_results.append(result)
        jobs[job_id]["results"].append(result)
        jobs[job_id]["progress"] = index + 1

        if len(chunk_results) == CHUNK_SIZE:
            _save_chunk(job_id, chunk_results, chunk_index)
            chunk_index += 1
            chunk_results = []

    if chunk_results:
        _save_chunk(job_id, chunk_results, chunk_index)

    jobs[job_id]["status"] = "complete"


def _save_chunk(job_id, results, chunk_index):
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads", "batch_output")
    os.makedirs(downloads_dir, exist_ok=True)
    filename = os.path.join(downloads_dir, f"results_{job_id}_part{chunk_index}.xlsx")
    pd.DataFrame(results).to_excel(filename, index=False)
    print(f"[Batch {job_id}] Saved chunk {chunk_index} ({len(results)} parts) -> {filename}")
    jobs[job_id].setdefault("saved_files", []).append(filename)


@app.route("/batch_start", methods=["POST"])
def batch_start():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        batch_df = pd.read_excel(io.BytesIO(request.files["file"].read()))
    except Exception as exc:
        return jsonify({"error": f"Could not read Excel file: {exc}"}), 400

    missing = [c for c in (BATCH_PART_COL, BATCH_NAME_COL) if c not in batch_df.columns]
    if missing:
        return jsonify({
            "error": f"File must have '{BATCH_PART_COL}' and '{BATCH_NAME_COL}' columns "
                     f"(missing: {', '.join(missing)})"
        }), 400

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"progress": 0, "total": len(batch_df), "status": "processing", "results": []}
    threading.Thread(target=_process_batch_job, args=(job_id, batch_df), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/batch_files/<job_id>", methods=["GET"])
def batch_files(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({"saved_files": job.get("saved_files", [])})


@app.route("/batch_status/<job_id>", methods=["GET"])
def batch_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


if __name__ == "__main__":
    app.run(debug=False, port=5000)