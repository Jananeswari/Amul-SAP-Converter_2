"""
Core Mapping Engine
Loads the master Article Master file, parses case-pack sizes from SAP
descriptions, and provides lookup utilities used by the ETL pipeline.
"""

import re
import pandas as pd
import os
import shutil

MASTER_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "Amul_Article_Master.xlsx")
PREVIOUS_MASTER_FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "Amul_Article_Master_PREVIOUS.xlsx")

# Platform columns present in the master file (order matters for display)
PLATFORM_COLUMNS = [
    "Amazon", "Bigbasket", "D mart", "Flipkart", "Metro",
    "Reliance", "Swiggy", "Zepto", "Star", "Lots", "Spar", "Blinkit",
]

# Maps the platform name shown in the app/upload UI to the actual column
# name used in the master file. Most are identical; a few app-facing names
# are more descriptive (e.g. carry the operating company) than the short
# column header in the spreadsheet.
PLATFORM_DISPLAY_TO_COLUMN = {
    "Zepto": "Zepto",
    "Swiggy Instamart (Jupiter Kart)": "Swiggy",
    "Blinkit (Zomato Hyperpure)": "Blinkit",
    "BigBasket": "Bigbasket",
    "Reliance Retail": "Reliance",
}


def clean_sku_value(val) -> str:
    """
    Normalizes a raw master-file SKU cell into a comparable string.
    Master file SKUs are sometimes stored as Excel numbers (e.g. 2049.0)
    even though the platform's own file shows them as plain integers
    (e.g. '2049') or GUIDs/strings. This avoids the '2049' vs '2049.0'
    mismatch that would otherwise silently break the mapping.
    """
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if re.fullmatch(r"-?\d+\.0+", s):  # e.g. "2049.0" -> "2049"
        s = s.split(".")[0]
    return s.upper()


def parse_case_pack(description: str):
    """
    Extracts the number of consumer units packed into one SAP case/box
    from a free-text SAP product description.
    Returns (case_pack_size: int, confidence: str)
    confidence is one of: "high", "low" (ambiguous combo pack), "default" (assumed 1).
    """
    if not isinstance(description, str) or not description.strip():
        return 1, "default"
    d = description.strip()

    # Ambiguous combo packs like (1+1f)X19) -- flag as low confidence, don't guess silently
    if re.search(r'\(\s*\d+\s*\+\s*\d+\s*f?\s*\)', d, re.IGNORECASE):
        m = re.search(r'\(\s*\d+\s*\+\s*\d+\s*f?\s*\)\s*[xX]\s*(\d+)', d, re.IGNORECASE)
        if m:
            return int(m.group(1)), "low"
        return 1, "low"

    # N x ( N x ... )  e.g. 8x(6x200g)
    m = re.search(r'(\d+)\s*[xX]\s*\(\s*(\d+)\s*[xX*]', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # N x N x num unit   e.g. 8x8x200 Gm, 9x20x50 Gm
    m = re.search(r'(\d+)\s*[xX]\s*(\d+)\s*[xX]\s*[\d.]+\s*[a-zA-Z]', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # Trailing (1x10), (3x6), (16x9) style — case-count x units-per-case,
    # the two numbers MULTIPLY to give total units per case (e.g. "(3x6)"
    # on an ice cream tub means 3 trays x 6 tubs = 18 tubs/case). This must
    # run BEFORE the general "N x number+unit" rule below, since that rule
    # would otherwise match the first number inside the parenthesis (e.g.
    # "3x6" -> 3) and silently undercount the real case-pack size.
    m = re.search(r'\((\d+)\s*[xX]\s*(\d+)\)', d)
    if m:
        return int(m.group(1)) * int(m.group(2)), "high"

    # N x num(.num) unit-word required  e.g. 30x200ml, 24 x 500 ml, 12x1 Litre, 60x200gm
    # The unit word is now REQUIRED (not optional) so this can't accidentally
    # match a bare "NxM" that's actually a parenthesis case-pack pattern
    # handled above (which has no unit word attached to the second number).
    # Unit word list below was verified against every NxNum pattern actually
    # present in the master file's "Product Description as per SAP" column.
    m = re.search(r'(\d+)\s*[xX*]\s*[\d.]+\s*(?:gm|g|ml|mlpet|kg|kgs|l|lt|ltr|lit|litre|liters?)\b', d, re.IGNORECASE)
    if m:
        return int(m.group(1)), "high"

    # Bare N x num with NO unit word at all and NOT inside parentheses
    # e.g. "60x200" with nothing else -- still treat first number as case count
    m = re.search(r'(?<!\()(\d+)\s*[xX*]\s*[\d.]+(?!\s*[a-zA-Z])(?!\))', d)
    if m:
        return int(m.group(1)), "high"

    # No pack pattern found -> assume sold as single unit (case pack = 1)
    return 1, "default"


def load_master_mapping(file_path: str = None) -> pd.DataFrame:
    """
    Loads and normalizes the Article Master file.
    Returns a long-format DataFrame: one row per (SAP Code, Platform, Platform SKU).
    `platform` here uses the APP-FACING display name (e.g. "Swiggy Instamart
    (Jupiter Kart)"), not the raw master-file column name, so it can be
    joined directly against ETL output.
    """
    path = file_path or MASTER_FILE_PATH
    df = pd.read_excel(path, sheet_name="Sheet1")
    df.columns = [c.strip() for c in df.columns]

    df = df.dropna(subset=["SAP Code"]).copy()
    df["case_pack_size"], df["pack_confidence"] = zip(
        *df["Product Description as per SAP"].apply(parse_case_pack)
    )

    long_rows = []
    for _, row in df.iterrows():
        for display_name, master_col in PLATFORM_DISPLAY_TO_COLUMN.items():
            sku_val = row.get(master_col)
            sku_clean = clean_sku_value(sku_val)
            if sku_clean:
                long_rows.append({
                    "platform": display_name,
                    "platform_sku": sku_clean,
                    "sap_code": row["SAP Code"],
                    "sap_description": row["Product Description as per SAP"],
                    "fg_group": row.get("FG Group Description", ""),
                    "case_pack_size": row["case_pack_size"],
                    "pack_confidence": row["pack_confidence"],
                })
    return pd.DataFrame(long_rows)


def save_master_mapping(df_wide: pd.DataFrame, file_path: str = None):
    """
    Saves the wide-format master mapping back to disk. Before overwriting,
    snapshots whatever is currently on disk to PREVIOUS_MASTER_FILE_PATH,
    so there is always a one-step-back copy available for comparison
    (point 6: previous version download). This snapshot happens on EVERY
    save — add, bulk update, or correction — automatically.
    """
    path = file_path or MASTER_FILE_PATH
    if os.path.exists(path):
        shutil.copy(path, PREVIOUS_MASTER_FILE_PATH)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_wide.to_excel(writer, sheet_name="Sheet1", index=False)


def load_master_wide(file_path: str = None) -> pd.DataFrame:
    """Loads the original wide-format master file (one row per product, platform columns side by side)."""
    path = file_path or MASTER_FILE_PATH
    df = pd.read_excel(path, sheet_name="Sheet1")
    df.columns = [c.strip() for c in df.columns]
    return df


def add_new_mapping_row(new_row: dict, file_path: str = None):
    """Appends a brand-new product mapping row to the master file.
    Use this only when the SAP Code does not exist yet — for adding a
    platform SKU to an SAP Code that already exists, use
    update_or_add_mapping instead."""
    df_wide = load_master_wide(file_path)
    df_wide = pd.concat([df_wide, pd.DataFrame([new_row])], ignore_index=True)
    save_master_mapping(df_wide, file_path)
    return df_wide


def update_or_add_mapping(sap_code: str, sap_desc: str, fg_group: str,
                           platform_inputs: dict, file_path: str = None) -> dict:
    """
    Adds a new SAP product, OR — if the SAP Code already exists — attaches
    new platform SKU codes to that existing row.

    A platform SKU is "new" for that SAP code if the platform's column on
    that row is currently empty. If the platform column already holds a
    DIFFERENT SKU, that platform is reported as a conflict and is NOT
    overwritten, so existing mappings can never be silently clobbered.

    Returns a dict:
        {
          "action": "added_new_product" | "updated_existing" | "no_change",
          "updated_platforms": [...],   # platforms newly filled in
          "conflicts": {platform: existing_sku, ...},  # platforms that already had a different SKU
        }
    """
    df_wide = load_master_wide(file_path)
    platform_inputs = {p: v.strip() for p, v in platform_inputs.items() if v and v.strip()}

    existing_mask = df_wide["SAP Code"].astype(str).str.strip().str.upper() == str(sap_code).strip().upper()

    if not existing_mask.any():
        new_row = {
            "FG Group": "", "FG Group Description": fg_group,
            "Article Description ": sap_desc, "GTIN": None,
            "SAP Code": sap_code, "Product Description as per SAP": sap_desc,
        }
        new_row.update(platform_inputs)
        df_wide = pd.concat([df_wide, pd.DataFrame([new_row])], ignore_index=True)
        save_master_mapping(df_wide, file_path)
        return {"action": "added_new_product", "updated_platforms": list(platform_inputs.keys()), "conflicts": {}}

    row_idx = df_wide[existing_mask].index[0]
    updated_platforms, conflicts = [], {}

    for platform, new_sku in platform_inputs.items():
        if platform not in df_wide.columns:
            df_wide[platform] = pd.NA
        current_val = df_wide.at[row_idx, platform]
        current_val_clean = clean_sku_value(current_val)
        new_sku_clean = clean_sku_value(new_sku)

        if not current_val_clean:
            df_wide.at[row_idx, platform] = new_sku
            updated_platforms.append(platform)
        elif current_val_clean != new_sku_clean:
            conflicts[platform] = str(current_val).strip()
        # if it's already the same SKU, no change needed

    if updated_platforms:
        save_master_mapping(df_wide, file_path)
        return {"action": "updated_existing", "updated_platforms": updated_platforms, "conflicts": conflicts}

    return {"action": "no_change", "updated_platforms": [], "conflicts": conflicts}


def bulk_update_from_unmapped_list(filled_df: pd.DataFrame, file_path: str = None) -> dict:
    """
    Takes the manager's filled-in "Unmapped SKUs" sheet — same shape as the
    download (Platform | SKU | Product Name (from platform) | Qty Ordered),
    plus a SAP Code column the manager has typed in — and applies the same
    safe rule as the single-entry form to every row:

      - if the SAP Code doesn't exist at all -> that row is skipped and
        reported (bulk update only attaches SKUs to EXISTING SAP codes,
        since there's no description/FG group supplied here)
      - if the platform column on that SAP Code's row is empty -> filled in
      - if it already holds a DIFFERENT SKU -> left untouched, reported as
        a conflict
      - if it already holds the SAME SKU -> skipped silently

    Expected columns in filled_df (case-insensitive, whitespace-tolerant):
        Platform, SKU, SAP Code
    (Product Name / Qty Ordered columns, if present, are ignored.)

    Returns a summary dict:
        {
          "updated_count": int,
          "skipped_no_sap_code": [...],      # rows where SAP Code cell was left blank
          "skipped_unknown_sap_code": [...], # rows whose typed SAP Code doesn't exist in master
          "conflicts": [...],                # rows where the platform already had a different SKU
          "updated_rows": [...],             # rows successfully applied
        }
    """
    cols_lower = {c.strip().lower(): c for c in filled_df.columns}
    required = ["platform", "sku", "sap code"]
    missing = [r for r in required if r not in cols_lower]
    if missing:
        raise ValueError(
            f"Uploaded file is missing required column(s): {', '.join(missing)}. "
            f"Found columns: {list(filled_df.columns)}"
        )

    platform_col = cols_lower["platform"]
    sku_col = cols_lower["sku"]
    sap_col = cols_lower["sap code"]

    df_wide = load_master_wide(file_path)
    sap_lookup = {
        str(code).strip().upper(): idx
        for idx, code in df_wide["SAP Code"].items() if pd.notna(code)
    }

    updated_count = 0
    skipped_no_sap_code, skipped_unknown_sap_code, conflicts, updated_rows = [], [], [], []

    for _, row in filled_df.iterrows():
        platform = str(row[platform_col]).strip() if pd.notna(row[platform_col]) else ""
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        sap_code_in = str(row[sap_col]).strip() if pd.notna(row[sap_col]) else ""

        if not sku or not platform:
            continue  # nothing to do for a blank platform/sku row

        if not sap_code_in:
            skipped_no_sap_code.append({"platform": platform, "sku": sku})
            continue

        row_idx = sap_lookup.get(sap_code_in.upper())
        if row_idx is None:
            skipped_unknown_sap_code.append({"platform": platform, "sku": sku, "sap_code_entered": sap_code_in})
            continue

        if platform not in df_wide.columns:
            df_wide[platform] = pd.NA
        current_val = df_wide.at[row_idx, platform]
        current_val_clean = clean_sku_value(current_val)
        sku_clean = clean_sku_value(sku)

        if not current_val_clean:
            df_wide.at[row_idx, platform] = sku
            updated_count += 1
            updated_rows.append({"platform": platform, "sku": sku, "sap_code": sap_code_in})
        elif current_val_clean != sku_clean:
            conflicts.append({"platform": platform, "sku": sku, "sap_code": sap_code_in, "existing_sku": str(current_val).strip()})
        # else: identical SKU already mapped -> nothing to do

    if updated_count:
        save_master_mapping(df_wide, file_path)

    return {
        "updated_count": updated_count,
        "skipped_no_sap_code": skipped_no_sap_code,
        "skipped_unknown_sap_code": skipped_unknown_sap_code,
        "conflicts": conflicts,
        "updated_rows": updated_rows,
    }


# ─────────────────────────────────────────────────────────────────────────
# CORRECTION TOOL (point 3) — deliberately overwrites an existing mapping
# that was found to be wrong. Unlike Add New Mapping / Bulk Update, which
# refuse to touch a platform column that already holds a value, this tool
# exists specifically to fix that value on purpose. It always requires
# the caller to have looked up and shown the current value first, so the
# UI can get explicit confirmation before anything is overwritten.
# ─────────────────────────────────────────────────────────────────────────

def find_current_mapping(sap_code: str = None, platform: str = None, sku: str = None,
                          file_path: str = None) -> pd.DataFrame:
    """
    Looks up current mapping rows, for use BEFORE a correction, so the
    old value can be shown to the user. Supports three lookup directions:
      - by sap_code: returns that product's row with all platform columns
      - by platform + sku: returns whichever SAP Code row currently has
        that SKU in that platform's column (if any)
    Returns a DataFrame (may be empty if nothing matches).
    """
    df_wide = load_master_wide(file_path)

    if sap_code:
        mask = df_wide["SAP Code"].astype(str).str.strip().str.upper() == str(sap_code).strip().upper()
        return df_wide[mask]

    if platform and sku:
        if platform not in df_wide.columns:
            return df_wide.iloc[0:0]
        sku_clean = clean_sku_value(sku)
        mask = df_wide[platform].apply(lambda v: clean_sku_value(v) == sku_clean)
        return df_wide[mask]

    return df_wide.iloc[0:0]


def correct_mapping(sap_code: str, platform: str, new_sku: str, file_path: str = None) -> dict:
    """
    Deliberately overwrites the SKU mapped to `platform` for an EXISTING
    `sap_code`, regardless of what's currently there. This is the one
    function in the whole mapping engine that overwrites without the
    "only fill empty slots" safety rule — callers (the UI) MUST call
    find_current_mapping first and get explicit user confirmation, since
    this cannot be undone except by restoring the previous-version
    snapshot that save_master_mapping always takes.

    Returns a dict: {"success": bool, "old_value": str, "message": str}
    """
    df_wide = load_master_wide(file_path)
    mask = df_wide["SAP Code"].astype(str).str.strip().str.upper() == str(sap_code).strip().upper()

    if not mask.any():
        return {"success": False, "old_value": None,
                "message": f"SAP Code '{sap_code}' was not found in the master file."}

    row_idx = df_wide[mask].index[0]
    if platform not in df_wide.columns:
        df_wide[platform] = pd.NA

    old_value = df_wide.at[row_idx, platform]
    old_value_str = str(old_value).strip() if pd.notna(old_value) else ""

    df_wide.at[row_idx, platform] = new_sku.strip()
    save_master_mapping(df_wide, file_path)

    return {
        "success": True,
        "old_value": old_value_str,
        "message": f"{platform} for SAP Code {sap_code} updated from "
                    f"'{old_value_str or '(empty)'}' to '{new_sku.strip()}'.",
    }


# ─────────────────────────────────────────────────────────────────────────
# FUZZY MATCHING for unmapped SKUs (point 4) — suggests likely SAP Code
# matches by comparing the unmapped platform's product NAME text against
# every SAP description in the master file, using Python's built-in
# difflib (no extra dependency required).
# ─────────────────────────────────────────────────────────────────────────

def suggest_fuzzy_matches(product_name: str, top_n: int = 5, file_path: str = None) -> pd.DataFrame:
    """
    Given an unmapped platform product name, returns the top_n closest
    SAP descriptions by text similarity, each with a 0-100 similarity
    score, for the user to review and pick from (never auto-applied).
    """
    import difflib

    if not product_name or not str(product_name).strip():
        return pd.DataFrame(columns=["SAP Code", "SAP Product Description", "FG Group Description", "similarity"])

    df_wide = load_master_wide(file_path)
    df_wide = df_wide.dropna(subset=["SAP Code", "Product Description as per SAP"]).copy()

    query = str(product_name).strip().lower()
    descriptions = df_wide["Product Description as per SAP"].astype(str).tolist()

    scores = [
        difflib.SequenceMatcher(None, query, desc.lower()).ratio() * 100
        for desc in descriptions
    ]
    df_wide["similarity"] = scores

    top = df_wide.sort_values("similarity", ascending=False).head(top_n)
    return top[["SAP Code", "Product Description as per SAP", "FG Group Description", "similarity"]].reset_index(drop=True)
