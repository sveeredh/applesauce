# --- libraries ---
import re
import pandas as pd

# Package families that are automatically no-cross on the competitor side.
_SMX_PACKAGE_RE = re.compile(r'\bSM(?:C|D|F|8)\b', re.IGNORECASE)
_DO_PACKAGE_RE  = re.compile(r'\bDO-?\d', re.IGNORECASE)   # DO-214AC, DO-41, DO15, ...
# SOD-123 and its body variants: SOD123, SOD-123F, SOD-123FL, SOD-123W, ...
# The trailing [A-Z]* takes the suffix letters, and the \b after it keeps
# SOD-128 / SOD-1234 out. SOD-323 / SOD-523 / SOD-882 are untouched.
_SOD123_PACKAGE_RE = re.compile(r'\bSOD[\s\-]?123[A-Z]*\b', re.IGNORECASE)

def is_blocked_competitor_package(package_str):
    """True if the competitor package is an SMx, DO-, or SOD-123 type -> force no cross."""
    if not package_str:
        return False
    s = str(package_str).strip()
    if not s or s.lower() in ('nan', 'none', '-'):
        return False
    return bool(
        _SMX_PACKAGE_RE.search(s)
        or _DO_PACKAGE_RE.search(s)
        or _SOD123_PACKAGE_RE.search(s)
    )

# SMA / SMB competitors may cross, but only to a TI part in the same body.
_SMX_MATCH_ONLY_RE = re.compile(r'\bSM(?:A|B)\b', re.IGNORECASE)


def requires_exact_package_match(package_str):
    """
    True for the SMx bodies that are allowed to cross only into their own
    package (SMA -> TI SMAF, SMB -> TI SMBF). Unlike the blocked families
    these are not refused outright, but every alternative without a package
    match has to be dropped, so a parametric win cannot put an SMA part into
    a DFN.
    """
    if not package_str:
        return False
    s = str(package_str).strip()
    if not s or s.lower() in ('nan', 'none', '-'):
        return False
    return bool(_SMX_MATCH_ONLY_RE.search(s))

def safe_strip(text):
    if pd.isna(text):
        return "-"
    if isinstance(text, (int, float)): # if it's already a number, convert to string
        return str(text).strip()
    return str(text).strip() if text else "-" # ensure text is string before stripping

def parse_capacitance(value_str):
    if not value_str or value_str == "-": return "-"
    match = re.search(r"([\d.]+)\s*(pF|nF|µF|uF)", value_str, re.IGNORECASE)
    if match:
        val = float(match.group(1)); unit = match.group(2).lower()
        if unit == "nf": val *= 1000
        elif unit in ["uf", "µf"]: val *= 1000000
        return f"{val:.2f} pF"
    return value_str

def parse_channels(value_str):
    if not value_str or value_str == "-": return "-"
    num_match = re.search(r"(\d+)", value_str)
    if num_match: return num_match.group(1)
    if "unidirectional" in value_str.lower() or "bidirectional" in value_str.lower(): return "1"
    return value_str

# --- In parsing.py, ADD this new function ---

def parse_polarity_and_channels_from_row(row):
    """
    Determines the number of channels and polarity based on specific column names
    in a DataFrame row. It gives priority to bidirectional channels.

    Args:
        row (pd.Series): A row from a competitor's DataFrame.

    Returns:
        dict: A dictionary with 'Channels' and 'Direction' keys.
    """
    # Find column names, case-insensitive
    bi_chan_col = next((c for c in row.index if "bidirectional channels" in str(c).lower()), None)
    uni_chan_col = next((c for c in row.index if "unidirectional channels" in str(c).lower()), None)

    # Check Bidirectional first, as it's the more specific case
    if bi_chan_col and pd.notna(row[bi_chan_col]) and row[bi_chan_col] > 0:
        try:
            return {
                "Channels": str(int(row[bi_chan_col])),
                "Direction": "Bidirectional"
            }
        except (ValueError, TypeError):
            pass  # Fall through if the value isn't a valid number

    # If no valid bidirectional data, check Unidirectional
    if uni_chan_col and pd.notna(row[uni_chan_col]) and row[uni_chan_col] > 0:
        try:
            return {
                "Channels": str(int(row[uni_chan_col])),
                "Direction": "Unidirectional"
            }
        except (ValueError, TypeError):
            pass  # Fall through if the value isn't a valid number

    # Default fallback if neither column has valid data
    return {"Channels": "-", "Direction": "-"}

def parse_price(price_str):
    if not price_str or price_str == "-": return "-"
    numeric_match = re.search(r"([\d.,]+)", price_str) # numeric part first
    if numeric_match:
        try:
            value = float(numeric_match.group(1).replace(',', '')) # convert to float            
            unit_price = value / 1000.0
            
            return f"${unit_price:.5f}" # format to 5 decimal places
        except ValueError:
            return price_str if '$' in price_str or any(char.isdigit() for char in price_str) else "-"
    return price_str

def normalize_aos_package(pkg_str_input):
    """
    Parses an AOS package string and returns its key properties in a dictionary.
    Returns None if the package string is not a recognized AOS format.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    
    norm_pkg = str(pkg_str_input).strip().upper()

    # Rule: SOT23(-XL?) -> {'type': 'SOT23', 'pins': X}
    if norm_pkg.startswith("SOT23"):
        pins_match = re.search(r"SOT23-(\d+)", norm_pkg)
        pins = int(pins_match.group(1)) if pins_match else 3
        return {"type": "SOT23", "pins": pins}

    # Rule: DFN/WLCSP(A.A x B.B)-(Y)L -> {'type': 'DFN', 'length': A.A, 'width': B.B, 'pins': Y}
    
    if norm_pkg.startswith("DFN") or norm_pkg.startswith("WLCSP"):
        # This new, more robust regex handles formats with or without spaces/parentheses.
        # It looks for a prefix, then immediately captures numbers.
        match = re.search(
            r"^(?:DFN|WLCSP)\s*\(?(\d+\.?\d*)\s*[xX\*]\s*(\d+\.?\d*)\)?.*?[-_ ](\d+)", 
            norm_pkg
        )
        if match:
            length = float(match.group(1))
            width = float(match.group(2))
            pins = int(match.group(3))
            # Critically, we still classify the type as 'DFN' for matching purposes
            return {"type": "DFN", "length": length, "width": width, "pins": pins}

    # This will catch older, exact-match aliases from your original function
    if norm_pkg == "SOT-23":
        return {"type": "SOT23", "pins": 5, "legacy": True}
    if "SOT23-6" in norm_pkg:
        return {"type": "SOT23", "pins": 6, "legacy": True}

    return None # Return None if no rule matches

def normalize_amazing_package(pkg_str_input):
    """
    Parses an Amazing IC package string and returns its key properties in a dictionary.
    This version correctly handles hyphenated formats before cleaning the string.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    
    # Use the original, un-cleaned string for parsing to preserve hyphens.
    norm_pkg = str(pkg_str_input).strip().upper()
    # Create a cleaned version (no hyphens/spaces) for rules that need it.
    norm_pkg_cleaned = re.sub(r"[\s-]+", "", norm_pkg)

    # --- Rule: DFNxxxxP# -> Works best on the cleaned string ---
    dfn_match = re.search(r"^(?:DFN|CSP|LGA|MCSP)(\d{4})P(\d+)", norm_pkg_cleaned)
    if dfn_match:
        return {"type": "DFN", "dims": dfn_match.group(1), "pins": int(dfn_match.group(2))}

    # --- Rule: SOD-XYX -> Needs hyphen, so uses original string ---
    sod_match = re.search(r"^SOD-(\d)\d(\d)", norm_pkg)
    if sod_match:
        return {"type": "SOD_AMAZING", "outer_digits": f"{sod_match.group(1)}{sod_match.group(2)}"}

    # --- Rule: SC70-X -> Needs hyphen, so uses original string ---
    sc70_match = re.search(r"^SC70-(\d+)", norm_pkg)
    if sc70_match:
        return {"type": "SC70", "pins": int(sc70_match.group(1))}
    
    sot23_match = re.search(r"^(?:SOT23|SOT143)-(\d+)", norm_pkg)
    if sot23_match:
        pins = int(sot23_match.group(1))
        return {"type": "SOT23_AMAZING", "pins": pins}
    
    # --- Rule: SOT-YXY -> Needs hyphen, so uses original string ---
    sot_match = re.search(r"^SOT(\d)(\d)(\d)", norm_pkg_cleaned)
    if sot_match:
        # body is the first and third digits, pins is the middle digit.
        return {"type": "SOT", "body": f"{sot_match.group(1)}{sot_match.group(3)}", "pins": int(sot_match.group(2))}
    

    

    return None # Return None if no rule matches

def normalize_diodes_package(pkg_str_input):
    """
    Parses a Diodes Inc. package string into a structured dictionary for matching.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-": return None
    # Normalize by removing hyphens and making uppercase for easier regex
    norm_pkg = str(pkg_str_input).strip().upper().replace("-", "")

    # Rule: MSOP-X -> {'type': 'MSOP', 'pins': X}
    msop_match = re.search(r"^MSOP(\d+)", norm_pkg)
    if msop_match:
        return {"type": "MSOP", "pins": int(msop_match.group(1))}

    # Rule: SODXXX -> {'type': 'SOD', 'body': XXX}
    sod_match = re.search(r"^SOD(\d+)", norm_pkg)
    if sod_match:
        return {"type": "SOD", "body": sod_match.group(1)}

    # Rule: SOT2X or TSOT2X -> {'type': 'SOT2X', 'pins_str': '2X'}
    if norm_pkg == "SOT23":
        return {"type": "SOT2X", "pins": 3}
        
    sot2x_match = re.search(r"^(?:T)?SOT(2\d)", norm_pkg)
    if sot2x_match:
        pins_str = sot2x_match.group(1)  # Captures '23', '25', '26'
        pins = int(pins_str[-1])         # Extracts the last digit (3, 5, or 6)
        return {"type": "SOT2X", "pins": pins}
    

    # Rule: SOT3X3 -> {'type': 'SOT3X3', 'pins': X}
    if norm_pkg == "SOT323":
        # Force it into the SOT3X3 structure with pins=2
        norm_pkg = "SOT323"
        
    sot3x3_match = re.search(r"^SOT3(\d)3", norm_pkg)
    if sot3x3_match:
        pins = int(sot3x3_match.group(1))
        if pins == 2: pins = 3 # Treat 2-pin as 3-pin for matching SC70-3
        return {"type": "SOT3X3", "pins": pins}

    # Rule: SOTYXY -> {'type': 'SOTYXY', 'body': 'YY', 'pins': X}
    sotyxy_match = re.search(r"^SOT(\d)(\d)(\d)", norm_pkg)
    if sotyxy_match:
        body = f"{sotyxy_match.group(1)}{sotyxy_match.group(3)}"
        pins = int(sotyxy_match.group(2))
        return {"type": "SOTYXY", "body": body, "pins": pins}

    # Rule: (U/X1/X2...)-DFNAAAA-Y -> {'type': 'DFN_DIODES', 'dims': AAAA, 'pins': Y}
    dfn_match = re.search(r"(?:DFN|QFN|DSN|WLB)(\d{4})(\d+)", norm_pkg)
    if dfn_match:
        return {"type": "DFN_DIODES", "dims": dfn_match.group(1), "pins": int(dfn_match.group(2))}

    return norm_pkg # Fallback to the normalized string if no rule matches


def normalize_nexperia_package(pkg_info):
    """
    Parses Nexperia package info and returns a structured dictionary or string 
    representing the TI equivalent package for matching.
    --- VERSION 5: Corrects SOT323 pin count ---
    """
    if not isinstance(pkg_info, dict):
        return None 

    pkg_name = str(pkg_info.get("name", "")).strip().upper()
    pkg_version = str(pkg_info.get("version", "")).strip().upper()

    if "TO-236AB" in pkg_name or "TO236AB" in pkg_name:
        return {"type": "NEXPERIA_SOT23", "pins": 3}
    if "SOT-323" in pkg_name or "SOT323" in pkg_name:
        return {"type": "NEXPERIA_SOT3X3", "pins": 3}
    
    dfn_match2 = re.search(r"^DSN0402", pkg_name)
    if dfn_match2:
        return {"type": "DFN", "dims": "0603"}

    # --- Rule Set 1: Primary match based on 'Package name' ---
    dfn_match = re.search(r"^(?:DFN|DSN)(\d{4}).*?(?:-(\d+))?", pkg_name)
    if dfn_match:
        dims = dfn_match.group(1)
        # Default to 2 pins if no pin count is specified after a hyphen.
        pins = int(dfn_match.group(2)) if dfn_match.group(2) else 2
        return {"type": "DFN", "dims": dims, "pins": pins}
  
    
    # --- Rule Set 2: New specific aliases based on 'Package version' ---
    sot3x3_match = re.search(r"^SOT3(\d)3", pkg_version)
    if sot3x3_match:
        pins = int(sot3x3_match.group(1))
        # FIX: The '2' in SOT323 refers to the body style, not the pin count.
        # This package is the 3-pin JEDEC standard SC-70.
        if pins == 2: pins = 3
        # Use a unique type to ensure this rule is handled by the Nexperia-specific logic
        return {"type": "NEXPERIA_SOT3X3", "pins": pins}

    sot66_match = re.search(r"^SOT66(\d)", pkg_version)
    if sot66_match:
        pins = int(sot66_match.group(1))
        # Use a unique type that the matcher will recognize
        return {"type": "NEXPERIA_SOT66Y", "pins": pins}
    
    sod_match = re.search(r"^SOD(\d+)", pkg_version)
    if sod_match:
        # Use a unique type for specific handling
        return {"type": "NEXPERIA_SOD", "body": sod_match.group(1)}

    if re.search(r"^SOT23", pkg_version):
        # This specifically targets the 3-pin "SOT23" package
        return {"type": "NEXPERIA_SOT23", "pins": 3}

    if re.search(r"^SOT143", pkg_version):
        return {"type": "NEXPERIA_SOT143"} 

    if re.search(r"^SOT457", pkg_version):
        return {"type": "NEXPERIA_SOT457"}

    # --- Rule Set 3: Older DFN aliases from 'Package version' ---
    if re.search(r"^SO[TD](?:882|883|993|995|8079)", pkg_version):
        return {"type": "DFN", "dims": "1006"}
    if re.search(r"^SOT1215", pkg_version):
        return {"type": "DFN", "dims": "1010"}
    if re.search(r"^SOT8015", pkg_version):
        return {"type": "DFN", "dims": "1110"}
    # ... (all other DFN alias rules remain here) ...
    if re.search(r"^SOT1268", pkg_version) or re.search(r"^SOT8009", pkg_version):
        return {"type": "DFN", "dims": "1412"}
    if re.search(r"^SOD1608", pkg_version):
        return {"type": "DFN", "dims": "1608"}
    if re.search(r"^SOT1061", pkg_version):
        return {"type": "DFN", "dims": "2020"}
    if re.search(r"^SOD9(?:62|72|92)", pkg_version):
        return {"type": "DFN", "dims": "0603"}
    if re.search(r"^SOT8013", pkg_version):
        return {"type": "DFN", "dims": "0603"}
    if re.search(r"^SOT8006", pkg_version):
        return {"type": "DFN", "dims": "1308"}
    if re.search(r"^SOT1165", pkg_version) or re.search(r"^SOT1176", pkg_version):
        return {"type": "DFN", "dims": "2510"}


    # If no rules match, return None to indicate no valid TI equivalent was found
    return None

def normalize_littelfuse_package(pkg_str_input):
    """
    Parses a Littelfuse package string and returns its key properties in a dictionary.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    norm_pkg = str(pkg_str_input).strip().upper()

    dfn_with_pins_match = re.search(r"^(?:U)?DFN(\d{4}).*?-(\d+)", norm_pkg)
    if dfn_with_pins_match:
        dims = dfn_with_pins_match.group(1)
        pins = int(dfn_with_pins_match.group(2))
        return {"type": "DFN_LITTEL", "dims": dims, "pins": pins}

    # If that fails, fall back to matching just the dimensions.
    dfn_without_pins_match = re.search(r"^(?:U)?DFN(\d{4})", norm_pkg)
    if dfn_without_pins_match:
        # No pin info is parsed, so Stage 2 filter will be skipped for this part.
        return {"type": "DFN_LITTEL", "dims": dfn_without_pins_match.group(1)}


    # Rule: xxxx WLCSP
    wlcsp_match = re.search(r"^(\d{4})\s*WLCSP", norm_pkg)
    if wlcsp_match:
        return {"type": "WLCSP_LITTEL", "dims": wlcsp_match.group(1)}

    # Rule: MSOP-10L
    if "MSOP-10" in norm_pkg:
        return {"type": "MSOP_LITTEL", "pins": 10}

    # Rule: SC70-X(?)
    sc70_match = re.search(r"^SC70-(\d)", norm_pkg)
    if sc70_match:
        return {"type": "SC70_LITTEL", "pins": int(sc70_match.group(1))}

    # Rule: SOD88X
    if norm_pkg.startswith("SOD88"):
        return {"type": "SOD88X_LITTEL"}

    # Rule: SODXXX(?)
    sod_match = re.search(r"^SOD(\d{3})", norm_pkg)
    if sod_match:
        return {"type": "SOD_LITTEL", "body": sod_match.group(1)}

    # Rule: SOT143
    if "SOT143" in norm_pkg:
        return {"type": "SOT143_LITTEL"}

    # Rule: SOT23-X(?)
    sot23_match = re.search(r"^SOT23-(\d)", norm_pkg)
    if sot23_match:
        return {"type": "SOT23_LITTEL", "pins": int(sot23_match.group(1))}
    # Rule: Handle SOT23 (no pin count) as 3-pin
    if norm_pkg == "SOT23":
        return {"type": "SOT23_LITTEL", "pins": 3}

    # Rule: SOTZYZ
    sotzyz_match = re.search(r"^SOT(\d)\d(\d)", norm_pkg)
    if sotzyz_match:
        return {"type": "SOTZYZ_LITTEL", "body": f"{sotzyz_match.group(1)}{sotzyz_match.group(2)}"}

    return None # Return None if no structured rule matches


def normalize_jiangsu_package(pkg_str_input):
    """
    Parses a Jiangsu package string into a structured dictionary for matching.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    norm_pkg = str(pkg_str_input).strip().upper()

    # Rule: DFNWBX.XxX.X -> DFNXXXX (e.g., DFNWB1.0x0.6-2L -> DFN1006)
    dfnwb_match = re.search(r"^DFNWB(\d\.?\d*)[xX](\d\.?\d*).*?-(\d+)", norm_pkg)
    if dfnwb_match:
        len_str = dfnwb_match.group(1).replace('.', '')
        wid_str = dfnwb_match.group(2).replace('.', '')
        len_formatted = (len_str + '0') if len(len_str) == 1 else len_str
        wid_formatted = (wid_str + '0') if len(wid_str) == 1 else wid_str
        dims = (len_formatted + wid_formatted).ljust(4, '0')[:4]
        pins = int(dfnwb_match.group(3))
        return {"type": "DFN_JIANGSU", "dims": dims, "pins": pins}

    dfn_with_pins_match = re.search(r"^DFN(\d{4}).*?-(\d+)", norm_pkg)
    if dfn_with_pins_match:
        dims = dfn_with_pins_match.group(1)
        pins = int(dfn_with_pins_match.group(2))
        return {"type": "DFN_JIANGSU", "dims": dims, "pins": pins}

    # Fallback for DFNxxxx without explicit pins
    dfn_without_pins_match = re.search(r"^DFN(\d{4})", norm_pkg)
    if dfn_without_pins_match:
        return {"type": "DFN_JIANGSU", "dims": dfn_without_pins_match.group(1)}

    sot3x3_match = re.search(r"^SOT-3(\d)3", norm_pkg)
    if sot3x3_match:
        pins = int(sot3x3_match.group(1))
        if pins == 2:
            pins = 3
        return {"type": "SOT3X3_JIANGSU", "pins": pins}
    
    # Rule: SOD-XXX -> SODXXX
    sod_match = re.search(r"^SOD-(\d{3})", norm_pkg)
    if sod_match:
        return {"type": "SOD_JIANGSU", "body": sod_match.group(1)}

    # Rule: SOT-143 -> SOT-23-4
    if norm_pkg == "SOT-143":
        return {"type": "SOT143_JIANGSU"}

    # Rule: SOT-23(-X?) -> SOT-23-X
    sot23_match = re.search(r"^SOT-23(?:-(\d))?", norm_pkg)
    if sot23_match:
        pins = int(sot23_match.group(1)) if sot23_match.group(1) else 3
        return {"type": "SOT23_JIANGSU", "pins": pins}
        
    # Rule: SOT-YXY -> SOT-YXY
    sotyxy_match = re.search(r"^SOT-(\d)(\d)(\d)", norm_pkg)
    if sotyxy_match:
        body = f"{sotyxy_match.group(1)}{sotyxy_match.group(3)}"
        pins = int(sotyxy_match.group(2))
        return {"type": "SOTYXY_JIANGSU", "body": body, "pins": pins}

    return None # Return None if no structured rule matches


def normalize_vishay_package():
    return

def normalize_semtech_package(pkg_str_input):
    """
    Parses a Semtech package string and returns its key properties in a dictionary.
    Keywords are searched, ignoring preceding/succeeding text.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    norm_pkg = str(pkg_str_input).strip().upper()

    # Rule: DFN0402 -> DFN1006 (special case)
    if "DFN0402" in norm_pkg:
        return {"type": "DFN_SEMTECH_SPECIAL", "name": "DFN0402"}
        
    # Rule: DFNXXXX -> DFNXXXX
    dfn_match = re.search(r"DFN(\d{4})", norm_pkg)
    if dfn_match:
        return {"type": "DFN_SEMTECH", "dims": dfn_match.group(1)}

    # Rule: MSOP 10L -> VSSOP
    if "MSOP" in norm_pkg and "10" in norm_pkg:
        return {"type": "MSOP_SEMTECH", "pins": 10}

    # Rule: SC-70 XL -> SC70-X
    sc70_match = re.search(r"SC-70\s*(\d)", norm_pkg)
    if sc70_match:
        return {"type": "SC70_SEMTECH", "pins": int(sc70_match.group(1))}

    # Rule: SODXXX -> SODXXX
    sod_match = re.search(r"SOD(\d{3})", norm_pkg)
    if sod_match:
        return {"type": "SOD_SEMTECH", "body": sod_match.group(1)}

    # Rule: SOT 666 -> SOT-5X3
    if "SOT" in norm_pkg and "666" in norm_pkg:
        return {"type": "SOT666_SEMTECH"}

    # Rule: SOT143 -> SOT-23-4
    if "SOT143" in norm_pkg:
        return {"type": "SOT143_SEMTECH"}

    # Rule: SOT23-X -> SOT-23-X
    sot23_match = re.search(r"SOT23-(\d)", norm_pkg)
    if sot23_match:
        return {"type": "SOT23_SEMTECH", "pins": int(sot23_match.group(1))}
    # Handle base SOT23 as 3-pin
    if "SOT23" in norm_pkg:
        return {"type": "SOT23_SEMTECH", "pins": 3}

    return None # Return None if no structured rule matches


def normalize_onsemi_package(pkg_str_input):
    return

def normalize_comchip_package(pkg_str_input):
    return


# --- package handling ---
def normalize_package(pkg_str_input):
    """
    Performs simple, string-based normalization of a package name for display
    and basic matching. ALWAYS returns a string.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return ""
    norm_pkg = str(pkg_str_input).strip().upper()
    
    norm_pkg = re.sub(r'\s+THIN\b', '', norm_pkg).strip()
    original_input_upper = norm_pkg 
    norm_pkg = re.sub(r'\s*\([A-Z0-9\-\s/]+\)', '', norm_pkg).strip()

    # Aliases for SOT-23-6
    if "SOT-23-6" == norm_pkg or "SOT26" == norm_pkg.replace("-","") or "TSOT26" == norm_pkg.replace("-","") or "TSOT-23-6" == norm_pkg  or "6TSOP" == norm_pkg.replace("-",""):
        return "SOT236"
    # Aliases for SOT-23-5
    if "SOT-23-5" == norm_pkg or "SOT25" == norm_pkg.replace("-","") or "DBV" in original_input_upper:
        return "SOT235"
    # Aliases for SOT-23-4. Competitors write the 4-pin SOT-23 body as
    # SOT24 / TSOT24; TI writes it "SOT-23-4", and the canonical code the
    # package matcher keys on is SOT234 (see _expand_comp_packages).
    _sot24_flat = norm_pkg.replace("-", "")
    if ("SOT-23-4" == norm_pkg or "SOT24" == _sot24_flat or "TSOT24" == _sot24_flat
            or "TSOT-23-4" == norm_pkg or "SOT234" == _sot24_flat):
        return "SOT234"
    # Aliases for SOT-23-3
    # --- AFTER ---
    # Aliases for SOT-23-3
    if "SOT-23-3" == norm_pkg or norm_pkg == "SOT-23" or "SOT233" == norm_pkg.replace("-","") or "TO2363" == norm_pkg.replace("-","") or "SC59" == norm_pkg.replace("-","") or "TO236AB" == norm_pkg.replace("-",""):
        return "SOT233"

    # Aliases for SC70-6 / SOT-363
    if ("SC70-6" == norm_pkg or "6TSSOP" == norm_pkg.replace("-","") or "SC88" == norm_pkg.replace("-","")
            or "SOT363" == norm_pkg.replace("-","") or "SOTSC706" == norm_pkg.replace("-","")):
        return "SC706"
    
    # TI writes the 3-pin SC-70 as "SOT-SC70"; Nexperia and others write "SC-70".
    # These are the same package, so every spelling collapses to SC703 -- which
    # also lets _classify_ti_package reach SC703_3 and pick up the DCKR suffix.
    _sc70_flat = norm_pkg.replace("-", "")
    if ("SC70-3" == norm_pkg or "SC703" == _sc70_flat or "SOT323" == _sc70_flat
            or "SC70" == _sc70_flat or "SOTSC70" == _sc70_flat or "SOTSC703" == _sc70_flat):
        return "SC703"
    # SOT-523 is a 3-pin SOT-5x3 package
    if norm_pkg in ("SOT-523", "SOT523"):
        return "SOT-5X3-3"
    # Aliases for SOT-563 / SOT-666
    if "SOT-563" == norm_pkg or "SOT563" == norm_pkg.replace("-","") or "SOT-5X3" == norm_pkg or "SOT5X3" == norm_pkg.replace("-","") or "SOT-553" == norm_pkg or "SOT553" == norm_pkg.replace("-",""):
        return "SOT563"
    if "SOT-666" == norm_pkg or "SOT666" == norm_pkg.replace("-",""):
        return "SOT666"
    # Comchip writes this size as "01005/CSP0402", where 0402 is the METRIC
    # name for a 0.4 x 0.2 mm body - not the imperial 0402 (1.0 x 0.5 mm) that
    # maps to DFN1006 below. TI has nothing that small, so it must be caught
    # before the substring alias that follows.
    if "01005" in norm_pkg.replace("-", "").replace(" ", ""):
        return "01005"

    # Aliases for 0402 / 1006 metric packages
    if "0402" in norm_pkg or "1006" in norm_pkg or "X2SON" in norm_pkg or "2-XDFN" in norm_pkg or "DFN1006" in norm_pkg.replace("-",""):
        return "DFN1006"

    # JEDEC DO- bodies, checked BEFORE the metric size aliases below: "DO-201AD"
    # flattens to "DO201AD", which contains "201" and was being read as a 0201
    # metric part (-> DFN0603). The three DO-214 bodies keep their SMx names.
    _do_flat = re.sub(r'[\s\-]+', '', norm_pkg)
    if re.match(r'^DO\d', _do_flat):
        if _do_flat.startswith("DO214AB"): return "SMC"
        if _do_flat.startswith("DO214AA"): return "SMB"
        if _do_flat.startswith("DO214AC"): return "SMA"
        return _do_flat

    # "0603/DFN1608": the 0603 is the IMPERIAL code for a 1.6 x 0.8 mm body, not
    # the metric 0603 (0.6 x 0.3 mm) handled below. Checked first so the
    # explicit DFN1608 wins.
    if "DFN1608" in norm_pkg.replace("-", "").replace(" ", ""):
        return "DFN1608"
    
    # Aliases for 0201 / 0603 metric packages.
    # The size code has to stand on its own rather than appear anywhere in the
    # string: "DFN3020-10L" normalizes to "DFN302010L", which contains "0201"
    # (DFN3-0201-0L) and was being read as a 0603 part.
    if (re.search(r'(?<!\d)0?201(?!\d)', norm_pkg) or re.search(r'(?<!\d)0603(?!\d)', norm_pkg)
            or "DFN0603" in norm_pkg.replace("-","")):
        return "DFN0603"
    
    # Aliases for DFN2510 / 10-UFDFN packages
    if "DFN2510" in norm_pkg.replace("-","") or "10-UFDFN" in norm_pkg or "10UFDFN" in norm_pkg.replace("-","") or "UFDFN-10" in norm_pkg:
        return "DFN2510"

    # Nexperia appends a body-variant letter to DFN sizes (DFN1110D-3), which TI
    # does not (DFN1110-3). Drop the letter so the two normalize the same.
    dfn_match = re.match(r'^DFN(1110|1010)[A-Z]?(\d+)$', norm_pkg.replace("-", ""))
    if dfn_match:
        return f"DFN{dfn_match.group(1)}{dfn_match.group(2)}"
    
    # Standardize common JEDEC body names
    if "DO214AB" in norm_pkg.replace("-",""): return "SMC"
    if "DO214AA" in norm_pkg.replace("-",""): return "SMB"
    if "DO214AC" in norm_pkg.replace("-",""): return "SMA"

    # The SMx family picks up prefixes and suffixes across manufacturers -
    # CSMA, SMB-FL, SMAJ, SMCJ, SMAF - that all name the same JEDEC body. TI
    # lists them plain, so everything reduces to SMA / SMB / SMC rather than
    # being carried into the comparison as a package TI has never heard of.
    smx_match = re.match(r'^C?(SM[ABC])[A-Z0-9]*$', re.sub(r'[\s\-]+', '', norm_pkg))
    if smx_match:
        return smx_match.group(1)
    
    # Fallback to a cleaned-up version of the input string
    norm_pkg_cleaned = re.sub(r"[\s-]+", "", norm_pkg)
    if norm_pkg_cleaned:
        return norm_pkg_cleaned
    return re.sub(r"[\s-]+", "", original_input_upper) if re.sub(r"[\s-]+", "", original_input_upper) else original_input_upper


# In parsing.py

def normalize_structured_package(pkg_str_input):
    """
    A reusable function to parse common package strings from various competitors
    into a structured dictionary for advanced matching. ALWAYS returns a dictionary
    or None.
    """
    if pd.isna(pkg_str_input) or pkg_str_input == "-":
        return None
    # THE FIX: Clean the string by removing hyphens before matching.
    norm_pkg = str(pkg_str_input).strip().upper().replace("-", "")


    # --- START OF ADDED BLOCK ---
    # Aliases for 0201 / 0603 metric packages (DFN0603, X1SON)
    if "0201" in norm_pkg or "0603" in norm_pkg or "DFN0603" in norm_pkg.replace("-","") :
        return {"type": "DFN_GENERIC", "dims": "0603"}

    # Aliases for 0402 / 1006 metric packages (DFN1006, X2SON)
    if "0402" in norm_pkg or "1006" in norm_pkg or "X2SON" in norm_pkg or "DFN1006" in norm_pkg.replace("-",""):
        return {"type": "DFN_GENERIC", "dims": "1006"}
    # --- END OF ADDED BLOCK ---

    # Rule: SC70-3 / SOT-323 -> TI SC70-3
    if norm_pkg.startswith("SC70") or "SOT323" in norm_pkg:
        # Try to find a specific pin count (e.g., from "SC70-5").
        pins_match = re.search(r'SC70(\d)', norm_pkg)
        # If no specific pin count is found, default to 3, as "SC-70"
        # and "SOT-323" almost always refer to the 3-pin variant.
        pins = int(pins_match.group(1)) if pins_match else 3
        return {"type": "SC70_GENERIC", "pins": pins}

    # Rule: SODXXX -> TI SOD-XXX
    sod_match = re.search(r"^SOD(\d{3})", norm_pkg)
    if sod_match:
        return {"type": "SOD_GENERIC", "body": sod_match.group(1)}

    # Rule: SOT-23-X (or SOT-23 for 3-pin) -> TI SOT-23-X
    sot23_match = re.search(r"^SOT-?23(?:-(\d))?", norm_pkg)
    if sot23_match:
        pins = int(sot23_match.group(1)) if sot23_match.group(1) else 3
        return {"type": "SOT23_GENERIC", "pins": pins}

    # Rule: SOT-5X3 or SOT-666 -> TI SOT-5X3
    if re.search(r"^SOT-?5\d3", norm_pkg) or "SOT-666" in norm_pkg:
        return {"type": "SOT5X3_GENERIC"}

    return None # Return None if no structured rule matches
# --- unit conversion ---
def to_numeric_val(value_str, default_if_error=float('inf')):
    if value_str is None or value_str == "-":
        return default_if_error
    value_str = str(value_str).lower()
    numeric_part = re.search(r"([\d.]+)", value_str)
    if not numeric_part:
        return default_if_error
    val = float(numeric_part.group(1))
    if "kv" in value_str: val *= 1000
    elif "mv" in value_str: val /= 1000
    # for capacitance (pF is base)
    elif "nf" in value_str: val *= 1000
    elif "uf" in value_str or "μf" in value_str: val *= 1000000
    # for current (nA is base)
    elif "µa" in value_str or "ua" in value_str or "μa" in value_str: val *= 1000
    elif "ma" in value_str: val *= 1000000
    return val