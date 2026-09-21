import pandas as pd
import re
from parsing import normalize_package

def _expand_comp_packages(valid_comp_pkgs_for_match):
    """Competitor package aliases -> the set of TI package names they can match.
    Module-level so both stage 1 and stage 2 use the identical set."""
    expanded_comp_pkgs = set(valid_comp_pkgs_for_match)
    for cp in list(expanded_comp_pkgs):
        # SOT23 without pin count defaults to 3-pin
        if cp == "SOT23":
            expanded_comp_pkgs.add("SOT233")

        # SOD882 variants -> DFN1006 (same footprint)
        if re.match(r"SOD.?882", cp):
            expanded_comp_pkgs.add("DFN1006")

        # DFN0402 / DFN-0402 -> DFN1006
        if re.match(r"DFN.?0402", cp):
            expanded_comp_pkgs.add("DFN1006")

        # SOT666 / SOT-666 -> SOT5X3
        if re.match(r"SOT.?666", cp):
            expanded_comp_pkgs.add("SOT5X3")

        # SOT363 / SOT-363 -> SC706
        if re.match(r"SOT.?363", cp):
            expanded_comp_pkgs.add("SC706")

        # SOT323 / SOT-323 -> SC703
        if re.match(r"SOT.?323", cp):
            expanded_comp_pkgs.add("SC703")

        # SOT143 / SOT-143 -> SOT234
        if re.match(r"SOT.?143", cp):
            expanded_comp_pkgs.add("SOT234")

        # SC88 / SC-88 -> SC706
        if re.match(r"SC.?88", cp):
            expanded_comp_pkgs.add("SC706")

        # 0603 / 603 bare token -> DFN0603
        if re.match(r"^0?603$", cp):
            expanded_comp_pkgs.add("DFN0603")

        # 0201 bare token -> DFN0603
        if re.match(r"^0?201$", cp):
            expanded_comp_pkgs.add("DFN0603")

        # 1006 bare token -> DFN1006
        if re.match(r"^1006$", cp):
            expanded_comp_pkgs.add("DFN1006")

        # X2SON -> DFN0603 (4-pin, but nearest TI family)
        if "X2SON" in cp:
            expanded_comp_pkgs.add("DFN0603")

        # MSOP -> VSSOP
        if cp.startswith("MSOP"):
            expanded_comp_pkgs.add("VSSOP")

        # WLCSP / DSBGA variants. TI's DSBGA is a 4-bump package, so a
        # competitor CSP that states a different bump count has no TI
        # equivalent - e.g. WLCSP0.43x0.23-2L, a 2-bump 0402-metric part,
        # must not match TI's DSBGA. A string with no stated bump count is
        # still expanded, as before.
        if "WLCSP" in cp or "DSBGA" in cp:
            _csp_leads = re.search(r'(\d)L$', cp.upper())
            if _csp_leads is None or _csp_leads.group(1) == '4':
                expanded_comp_pkgs.add("DSBGA")

        # SOT886 / SOT-886 -> SOT886
        if re.match(r"SOT.?886", cp):
            expanded_comp_pkgs.add("SOT886")

        # Amazing DFN{dims}P{pins}[letter] e.g. DFN2020P3E -> DFN2020
        _amazing_m = re.match(r'(?:U|W|X\d?)?[DQ]FN(\d+)P(\d+)[A-Z]*$', cp, re.IGNORECASE)
        if _amazing_m:
            _dfn_fam_map = {'0603': 'DFN0603', '1006': 'DFN1006', '1110': 'DFN1110',
                            '1610': 'DFN1610', '1616': 'DFN1616', '2020': 'DFN2020', '2510': 'DFN2510', '3030': 'DFN3030'}
            _fam = _dfn_fam_map.get(_amazing_m.group(1))
            if _fam:
                expanded_comp_pkgs.add(_fam)

        # DFN{family}{pins} normalized style e.g. "DFN20206" -> DFN2020
        _canon_dfn = re.match(r'(DFN\d{4})(\d+)$', cp)
        if _canon_dfn:
            expanded_comp_pkgs.add(_canon_dfn.group(1))

        # DFN{dims}-{pins} hyphen style e.g. "DFN2020-3" -> DFN2020
        _hyphen_dfn_exp = re.match(r'(?:U|W|X\d?)?[DQ]FN(\d+)-(\d+)[A-Z]*$', cp, re.IGNORECASE)
        if _hyphen_dfn_exp:
            _dfn_fam_exp = {'0603': 'DFN0603', '1006': 'DFN1006', '1110': 'DFN1110',
                            '1610': 'DFN1610', '1616': 'DFN1616', '2020': 'DFN2020', '2510': 'DFN2510', '3030': 'DFN3030'}
            _fam_exp = _dfn_fam_exp.get(_hyphen_dfn_exp.group(1))
            if _fam_exp:
                expanded_comp_pkgs.add(_fam_exp)

        # TO-263 variants -> SOT233 (SOT-23-3)
        if re.match(r'TO.?263[A-Z0-9-]*$', cp, re.IGNORECASE):
            expanded_comp_pkgs.add("SOT233")

        # DFN-2L -> DFN1006 (2-pin)
        if re.match(r'DFN.?2L$', cp, re.IGNORECASE):
            expanded_comp_pkgs.add("DFN1006")

        # ---- TI package equivalence aliases (all competitors) ----
        flat = re.sub(r'[\s\-_]+', '', cp.upper())
        if re.match(r'^(SOT363|TSSOP6|SC706|SOT3236|SC886)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SC706")
        if re.match(r'^(SOTSC70|SOT3233)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SC703")
        _sc70 = re.match(r'^SC70(\d*)[A-Z]*$', flat)
        if _sc70:
            _p = _sc70.group(1)
            if _p in ('', '3'):
                expanded_comp_pkgs.add("SC703")
            elif _p == '6':
                expanded_comp_pkgs.add("SC706")
        if re.match(r'^SOT3[A-Z]*$', flat):
            expanded_comp_pkgs.add("SOT9X3")
        # SOT-9X3: X is the pin count, so a competitor part reaches TI rows
        # spelled generically (SOT-9X3) or with its own pin count (SOT-953).
        # TI's 3-pin member is also spelled SOT-923.
        _sot9x3_exp = re.match(r'^SOT9(?:X3(\d+)|(\d)3)', flat)
        if _sot9x3_exp:
            expanded_comp_pkgs.add("SOT9X3")
            _p9x3_exp = _sot9x3_exp.group(1) or _sot9x3_exp.group(2)
            if _p9x3_exp == '2':
                _p9x3_exp = '3'
            if _p9x3_exp:
                expanded_comp_pkgs.add(f"SOT9{_p9x3_exp}3")
                if _p9x3_exp == '3':
                    expanded_comp_pkgs.add("SOT923")
        if re.match(r'^(SOT665|SOT523)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT5X3")
        if re.match(r'^(SOT553|SOT563|SOD563|SOT89|SC89)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT5X3")
        # TI-style "SOT-5X3-{pins}" (the pin count carried in from the canonical
        # package) still matches TI's SOT-5X3 family in stage 1.
        if re.match(r'^SOT-?5X3-?\d+[A-Z]*$', re.sub(r'[\s_]+', '', cp.upper())):
            expanded_comp_pkgs.add(normalize_package("SOT-5X3"))
            expanded_comp_pkgs.add("SOT5X3")
        if re.match(r'^(SC79|SOD5232|SOD523)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOD523")
        if re.match(r'^(SC90|SOD323)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOD323")
        if re.match(r'^(TO236|SC59|SOT233)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT233")
        if re.match(r'^(SOT143|SOT1434|SOT1234|SOT234)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT234")
        if re.match(r'^(SOT25|TSOT25|SC74|TSOP5|SOT235)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT235")
        if re.match(r'^(SOT457|TSOP6|SOT26|TSOT26|SOT236)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("SOT236")
        if re.match(r'^(SOT8833|LLP10063)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("DFN1006")
        if re.match(r'^(X1SON|SOD882)[A-Z]*$', flat):
            expanded_comp_pkgs.add("DFN1006")
        if re.match(r'^(X2SON|DFN0606|X2DFN8084|X2DFN|DFN4L)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("DFN0603")
            expanded_comp_pkgs.add("X2SON")
        if re.match(r'^(USON6|USON1616\d*)[A-Z]*$', flat):
            expanded_comp_pkgs.add("USON")
        # SOT-886 and TI's 6-pin USON are the same package under two names, so a
        # competitor SOT-886 (canonical USON_6) has to reach rows spelled either way.
        if re.match(r'^USON-?6[A-Z]*$', cp.upper()) or re.match(r'^USON6[A-Z]*$', flat):
            expanded_comp_pkgs.add("SOT886")
        # 6-UFDFN is TI's 6-pin DFN1616.
        _ufdfn_exp = re.match(r'^(\d+)U?F?DFN', flat)
        if _ufdfn_exp and int(_ufdfn_exp.group(1)) == 6:
            expanded_comp_pkgs.add("DFN1616")
        # Bare family with a -{pins} suffix, e.g. USON-6, X2SON-6, WSON-15
        _fam_pin = re.match(r'^([A-Z]\d?[A-Z]+)-?(\d+)[A-Z]*$', cp.upper())
        if _fam_pin and _fam_pin.group(1) in {'USON','X2SON','X1SON','WSON','UQFN','WQFN','DSBGA'}:
            expanded_comp_pkgs.add(_fam_pin.group(1))
        if re.match(r'^(QFN10L|QFN10|UDFN10)[A-Z]*$', flat):
            expanded_comp_pkgs.add("DFN2510")
        if re.match(r'^(XSON6|QFN6L|QFN6)[A-Z0-9]*$', flat):
            expanded_comp_pkgs.add("DFN2510")

    return expanded_comp_pkgs


def find_package_matches(comp_pkg_alias, ti_df):
    part_num_col = next((c for c in ti_df.columns if 'product or part number' in c.lower()), None)

    pkg_col = next((c for c in ti_df.columns if 'package name' in c.lower()), None)
    if not pkg_col:
        print("! Warning: 'Package name' column not found in TI database for matching.")
        empty = pd.DataFrame()
        return empty, empty

    print(f"Stage 1 - Filtering by competitor packages on {len(ti_df)} TI parts...")

    is_structured_pkg = isinstance(comp_pkg_alias, dict)
    valid_comp_pkgs_for_match = []
    if not is_structured_pkg:
        for _raw_cp in str(comp_pkg_alias).split(','):
            _raw_cp = _raw_cp.strip()
            if not _raw_cp:
                continue
            valid_comp_pkgs_for_match.append(normalize_package(_raw_cp))
            # "SOT-5X3-6" carries the pin count for the stage 2 filter, but TI's
            # package name is plain "SOT-5X3" - stage 1 has to match that too.
            if re.match(r'^SOT-?5X3-?\d+[A-Z]*$', _raw_cp.upper()):
                valid_comp_pkgs_for_match.append(normalize_package("SOT-5X3"))

    def get_normalized_ti_pkg_list_from_cell(ti_pkg_cell_value):
        """
        For structured matching: checks only the first package in the cell (original behavior).
        For unstructured fallback: caller checks all packages via valid_comp_pkgs_for_match.
        """
        if pd.isna(ti_pkg_cell_value) or ti_pkg_cell_value == "-":
            return []
        first_pkg_str = str(ti_pkg_cell_value).split(',')[0].strip()
        pkgs_to_normalize = [p.strip() for p in first_pkg_str.split('/') if p.strip()]
        return [normalize_package(p) for p in pkgs_to_normalize if p]

    def get_all_normalized_ti_pkgs_from_cell(ti_pkg_cell_value):
        """Parses ALL comma-separated packages for unstructured fallback matching."""
        if pd.isna(ti_pkg_cell_value) or ti_pkg_cell_value == "-":
            return []
        result = []
        for pkg_token in str(ti_pkg_cell_value).split(','):
            for p in pkg_token.split('/'):
                p = p.strip()
                if p:
                    result.append(normalize_package(p))
        return result

    matched_indices = []
    comp_type = None
    if is_structured_pkg and comp_pkg_alias:
        comp_props = comp_pkg_alias
        comp_type = comp_props.get("type")
        

    for index, ti_row in ti_df.iterrows():
        ti_pkgs_cell = ti_row.get(pkg_col)
        ti_normalized_pkgs = get_normalized_ti_pkg_list_from_cell(ti_pkgs_cell)
        is_match = False
        
        # =================================================
            # --- DIODES INC. PACKAGE MATCHING RULES ---
        # =================================================
        if comp_type == "MSOP":
            expected_ti_pkg = f"VSSOP"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True
        
        elif comp_type == "SOD":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True; break
        
        elif comp_type == "SOT2X":
            if comp_props['pins'] == 3:
                if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                    is_match = True
            else:
                expected_ti_pkg = f"SOT23{comp_props['pins']}"
                if expected_ti_pkg in ti_normalized_pkgs:
                    is_match = True
        
        elif comp_type == "SOT3X3":
            ti_equivalent_pkg = f"SC70{comp_props['pins']}"
            if ti_equivalent_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOTYXY":
            ti_raw_pkg_string = ti_row.get(pkg_col, "")
            
            for ti_pkg_part in str(ti_raw_pkg_string).split(','):
                ti_pkg_part_cleaned = ti_pkg_part.strip().upper()
                
                ti_sot_match = re.search(r"SOT-?(\d)[\dX](\d)", ti_pkg_part_cleaned)
                
                if ti_sot_match:
                    ti_body = f"{ti_sot_match.group(1)}{ti_sot_match.group(2)}"
                    
                    if ti_body == comp_props["body"]:
                        is_match = True
                        break  
        elif comp_type == "DFN_DIODES":
            for ti_pkg in ti_normalized_pkgs:
                ti_dfn_match = re.search(r"DFN(\d{4})", ti_pkg)
                if ti_dfn_match and ti_dfn_match.group(1) == comp_props["dims"]:
                    is_match = True; break
                
        # =================================================
            # --- AOS PACKAGE MATCHING RULES ---
        # =================================================

        elif comp_type == "SOT23":
                    if comp_props['pins'] == 3:
                        if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                            is_match = True
                    else:
                        expected_ti_pkg = f"SOT23{comp_props['pins']}"
                        if expected_ti_pkg in ti_normalized_pkgs:
                            is_match = True
        
        elif comp_type == "DFN":
            if "length" in comp_props:
                for ti_pkg in ti_normalized_pkgs:
                    if ti_pkg.startswith("DFN"):
                        ti_dims_match = re.search(r"DFN(\d{2})(\d{2})", ti_pkg)
                        if not ti_dims_match: continue
                        ti_len = float(ti_dims_match.group(1)) / 10.0
                        ti_wid = float(ti_dims_match.group(2)) / 10.0
                        len_ok = (comp_props["length"] * 0.9 <= ti_len <= comp_props["length"] * 1.1)
                        wid_ok = (comp_props["width"] * 0.9 <= ti_wid <= comp_props["width"] * 1.1)
                        if len_ok and wid_ok:
                            is_match = True
                            break
            elif "dims" in comp_props:
                if f"DFN{comp_props['dims']}" in ti_normalized_pkgs:
                    is_match = True

        # =================================================
            # --- NEXPERIA PACKAGE MATCHING RULES ---
        # =================================================
        elif comp_type == "NEXPERIA_SOT3X3":
            expected_ti_pkg = f"SC70{comp_props['pins']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True
                
        elif comp_type == "NEXPERIA_SOT66Y":
            ti_raw_pkg_string = ti_row.get(pkg_col, "")
            
            for ti_pkg_part in str(ti_raw_pkg_string).split(','):
                ti_pkg_part_cleaned = ti_pkg_part.strip().upper()
                
                ti_sot_match = re.search(r"SOT-?(\d)[\dX](\d)", ti_pkg_part_cleaned)
                
                if ti_sot_match:
                    ti_body = f"{ti_sot_match.group(1)}{ti_sot_match.group(2)}"
                    
                    if ti_body == "53":
                        is_match = True
                        break  
        
        elif comp_type == "NEXPERIA_SOD":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True
                        break
                        
        elif comp_type == "NEXPERIA_SOT23":
            if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                is_match = True
            

        elif comp_type == "NEXPERIA_SOT143":
            if "SOT234" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "NEXPERIA_SOT457":
            if "SOT236" in ti_normalized_pkgs:
                is_match = True

        # =================================================
            # --- AMAZING PACKAGE MATCHING RULES ---
        # =================================================

        elif comp_type == "SOD_AMAZING":
                for ti_pkg in ti_normalized_pkgs:
                    ti_sod_match = re.search(r"^SOD(\d)\d(\d)", ti_pkg)
                    if ti_sod_match:
                        ti_outer_digits = f"{ti_sod_match.group(1)}{ti_sod_match.group(2)}"
                        
                        if ti_outer_digits == comp_props["outer_digits"]:
                            is_match = True
                            break 
                        
        elif comp_type == "SOT23_AMAZING":
                    if comp_props['pins'] == 3:
                        if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                            is_match = True
                    else:
                        expected_ti_pkg = f"SOT23{comp_props['pins']}"
                        if expected_ti_pkg in ti_normalized_pkgs:
                            is_match = True
        
        elif comp_type == "SOT":
            comp_body_style = comp_props.get("body")
            if pd.notna(ti_pkgs_cell) and isinstance(ti_pkgs_cell, str):
                for ti_pkg_str in ti_pkgs_cell.split(','):
                    ti_pkg_str = ti_pkg_str.strip().upper()
                    ti_sot_match = re.search(r"SOT-(\d)X(\d)", ti_pkg_str)
                    if ti_sot_match:
                        ti_body_style = f"{ti_sot_match.group(1)}{ti_sot_match.group(2)}"
                        if ti_body_style == comp_body_style:
                            is_match = True
                            break
        
        elif comp_type == "SC70":
            if f"SC70{comp_props['pins']}" in ti_normalized_pkgs:
                is_match = True

        # =================================================
            # --- JIANGSU PACKAGE MATCHING RULES ---
        # =================================================
        elif comp_type == "DFN_JIANGSU":
            expected_ti_pkg = f"DFN{comp_props['dims']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOD_JIANGSU":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True
                        break
        
        elif comp_type == "SOT143_JIANGSU":
            if "SOT234" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOT3X3_JIANGSU":
            expected_ti_pkg = f"SC70{comp_props['pins']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOT23_JIANGSU":
            if comp_props['pins'] == 3:
                if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                    is_match = True
            else:
                expected_ti_pkg = f"SOT23{comp_props['pins']}"
                if expected_ti_pkg in ti_normalized_pkgs:
                    is_match = True
        
        elif comp_type == "SOTYXY_JIANGSU":
            ti_raw_pkg_string = ti_row.get(pkg_col, "")
            
            for ti_pkg_part in str(ti_raw_pkg_string).split(','):
                ti_pkg_part_cleaned = ti_pkg_part.strip().upper()

                ti_sot_match = re.search(r"SOT-?(\d)[\dX](\d)", ti_pkg_part_cleaned)
                
                if ti_sot_match:
                    ti_body = f"{ti_sot_match.group(1)}{ti_sot_match.group(2)}"
                    
                    if ti_body == comp_props["body"]:
                        is_match = True
                        break 

        
        # =================================================
         # --- LITTELFUSE PACKAGE MATCHING RULES ---
        # =================================================
        elif comp_type == "DFN_LITTEL" or comp_type == "WLCSP_LITTEL":
            expected_ti_pkg = f"DFN{comp_props['dims']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "MSOP_LITTEL":
            if "VSSOP" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SC70_LITTEL":
            expected_ti_pkg = f"SC70{comp_props['pins']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOD88X_LITTEL":
            if "DFN1006" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOD_LITTEL":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True; break
        
        elif comp_type == "SOT143_LITTEL":
            if "SOT234" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOT23_LITTEL":
            if comp_props['pins'] == 3:
                if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                    is_match = True
            else:
                expected_ti_pkg = f"SOT23{comp_props['pins']}"
                if expected_ti_pkg in ti_normalized_pkgs:
                    is_match = True

        elif comp_type == "SOTZYZ_LITTEL":
            comp_body = comp_props.get("body")
            expected_ti_pkg = f"SOT{comp_body[0]}X{comp_body[1]}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        # =================================================
            # --- SEMTECH PACKAGE MATCHING RULES ---
        # =================================================
        elif comp_type == "DFN_SEMTECH_SPECIAL":
            if comp_props.get("name") == "DFN0402" and "DFN1006" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "DFN_SEMTECH":
            expected_ti_pkg = f"DFN{comp_props['dims']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "MSOP_SEMTECH":
            if "VSSOP" in ti_normalized_pkgs:
                is_match = True
        
        elif comp_type == "SC70_SEMTECH":
            expected_ti_pkg = f"SC70{comp_props['pins']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True
        
        elif comp_type == "SOD_SEMTECH":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True
                        break
        
        elif comp_type == "SOT666_SEMTECH":
            if "SOT5X3" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOT143_SEMTECH":
            if "SOT234" in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOT23_SEMTECH":
            if comp_props['pins'] == 3:
                if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                    is_match = True
            else:
                expected_ti_pkg = f"SOT23{comp_props['pins']}"
                if expected_ti_pkg in ti_normalized_pkgs:
                    is_match = True

        # =================================================
        # --- GENERIC/REUSABLE PACKAGE MATCHING RULES ---
        # =================================================
        elif comp_type == "DFN_GENERIC":
            # Generic DFNXXXX maps to TI's DFNXXXX
            expected_ti_pkg = f"DFN{comp_props['dims']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SC70_GENERIC":
            expected_ti_pkg = f"SC70{comp_props['pins']}"
            if expected_ti_pkg in ti_normalized_pkgs:
                is_match = True

        elif comp_type == "SOD_GENERIC":
            for ti_pkg in ti_normalized_pkgs:
                if ti_pkg.startswith("SOD"):
                    ti_sod_num = re.search(r"SOD(\d+)", ti_pkg)
                    if ti_sod_num and ti_sod_num.group(1) == comp_props["body"]:
                        is_match = True
                        break 

        elif comp_type == "SOT23_GENERIC":
            if comp_props['pins'] == 3:
                if "SOT23" in ti_normalized_pkgs or "SOT233" in ti_normalized_pkgs:
                    is_match = True
            else:
                expected_ti_pkg = f"SOT23{comp_props['pins']}"
                if expected_ti_pkg in ti_normalized_pkgs:
                    is_match = True
        
        elif comp_type == "SOT5X3_GENERIC":
            if "SOT5X3" in ti_normalized_pkgs:
                is_match = True
        
        if not is_structured_pkg:
            all_ti_pkgs = get_all_normalized_ti_pkgs_from_cell(ti_pkgs_cell)

            # Expand competitor packages with known equivalences so rules apply
            # regardless of which competitor database or DigiKey the specs came from.
            expanded_comp_pkgs = _expand_comp_packages(valid_comp_pkgs_for_match)

            sot23_pkgs = {"SOT233", "SOT234", "SOT235", "SOT236"}
            comp_has_sot23 = any(cp in sot23_pkgs for cp in expanded_comp_pkgs)
            if comp_has_sot23:
                # A SOT-23 competitor used to be matched against the FIRST
                # package in the TI cell only, which dropped every part whose
                # cell lists another body first ("SC70-3, SOT-23-3"). TI listing
                # two packages means the part ships in both, so the position in
                # the cell is not a reason to reject it. All packages are checked
                # now, as they already were for every other competitor family.
                if any(comp_pkg in all_ti_pkgs for comp_pkg in expanded_comp_pkgs):
                    is_match = True
            elif any(comp_pkg in all_ti_pkgs for comp_pkg in expanded_comp_pkgs):
                is_match = True
        if is_match:
            matched_indices.append(index)

        
    package_matches_df = ti_df.loc[list(set(matched_indices))].copy()
    print(f"Found {len(package_matches_df)} TI parts with matching packages (Stage 1).")
    if package_matches_df.empty and not is_structured_pkg:
        print(f"  Stage 1 debug - competitor pkgs tried: "
              f"{sorted(_expand_comp_packages(valid_comp_pkgs_for_match))}")

    if package_matches_df.empty:
        return package_matches_df, package_matches_df.copy(), package_matches_df.copy()

    if is_structured_pkg and comp_type in ["DFN", "DFN_JIANGSU", "DFN_LITTEL", "SOT", "SOTYXY", "DFN_DIODES", "NEXPERIA_SOT66Y", "SOTYXY_JIANGSU"]:
        if 'pins' not in comp_props:
            return package_matches_df, package_matches_df.copy(), package_matches_df.copy()

        pin_col = next((c for c in ti_df.columns if 'pin count' in c.lower()), None)
        if not pin_col or part_num_col is None:
            print("! Warning: 'Pin count' column not found in TI database. Skipping pin count filter.")
            return package_matches_df, package_matches_df.copy(), package_matches_df.copy()

        comp_pins = comp_props['pins']
        print(f"Stage 2 - Filtering by pin count ({comp_pins} pins)...")

        stage2_indices = []
        for index, ti_row in package_matches_df.iterrows():
            pin_str = str(ti_row.get(pin_col, ""))
            try:
                pin_vals = [int(p.strip()) for p in pin_str.split(',') if p.strip()]
            except ValueError:
                pin_vals = []
            if comp_pins in pin_vals:
                stage2_indices.append(index)

        final_matches_df = package_matches_df.loc[list(set(stage2_indices))].copy()
        print(f"Found {len(final_matches_df)} TI parts after pin count filter (Stage 2).")

        return package_matches_df, final_matches_df, final_matches_df
    
    elif not is_structured_pkg:
        comp_pin_count = None
        for raw_cp in str(comp_pkg_alias).split(','):
            raw_cp = raw_cp.strip()
            m = re.search(r'-(\d+)(?:[A-Za-z]*)$', raw_cp)
            if m:
                comp_pin_count = int(m.group(1))
                break
            m2 = re.match(r'(DFN\d{4})(\d+)$', raw_cp.upper())
            if m2:
                comp_pin_count = int(m2.group(2))
                break
            m3 = re.search(r'P(\d+)[A-Z]?$', raw_cp.upper())
            if m3:
                comp_pin_count = int(m3.group(1))
                break

        if comp_pin_count is None:
            print("Stage 2 - Pin count not determinable for unstructured package. Skipping.")
            return package_matches_df, package_matches_df.copy(), package_matches_df.copy()

        pin_col = next((c for c in ti_df.columns if 'pin count' in c.lower()), None)
        if not pin_col:
            return package_matches_df, package_matches_df.copy(), package_matches_df.copy()

        print(f"Stage 2 - Filtering unstructured package by pin count ({comp_pin_count} pins)...")

        # A part matches only if one of ITS packages, resolved to that package's
        # own pin count, equals the competitor's package+pin. Checking "does any
        # pin in the row equal N" wrongly passes parts whose matching package has
        # a different pin count (e.g. a 10-pin USON row that also contains a
        # 6-pin SC70-6).
        from applesauce import (_resolve_pin_count, _classify_ti_package,
                                _classify_digikey_package)

        pkg_col_s2 = next((c for c in ti_df.columns if 'package name' in c.lower()), None)

        comp_canon = None
        for _cp in str(comp_pkg_alias).split(','):
            comp_canon = _classify_digikey_package(_cp.strip(), "")
            if comp_canon:
                break

        stage2_indices = []
        for index, ti_row in package_matches_df.iterrows():
            pin_str = str(ti_row.get(pin_col, ""))
            pkg_str_s2 = str(ti_row.get(pkg_col_s2, "")) if pkg_col_s2 else ""
            gpn_s2 = str(ti_row.get(part_num_col, "")) if part_num_col else ""

            matched = False
            for raw_pkg in [p.strip() for p in pkg_str_s2.split(',') if p.strip()]:
                resolved = _resolve_pin_count(normalize_package(raw_pkg),
                                              pkg_str_s2, pin_str, gpn_s2)
                if resolved is None:
                    continue          # ambiguous -> not a confident match
                if comp_canon:
                    if _classify_ti_package(raw_pkg, resolved, gpn_s2) == comp_canon:
                        matched = True
                        break
                elif resolved == comp_pin_count:
                    matched = True
                    break

            if matched:
                stage2_indices.append(index)

        final_matches_df = package_matches_df.loc[list(set(stage2_indices))].copy()
        print(f"Found {len(final_matches_df)} TI parts after pin count filter (Stage 2).")
        return package_matches_df, final_matches_df, final_matches_df

    else:
        print("Stage 2 - Pin count filter not applicable for this package type. Skipping.")
        return package_matches_df, package_matches_df.copy(), package_matches_df.copy()