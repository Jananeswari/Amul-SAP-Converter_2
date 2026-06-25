"""
Platform File Readers
Each e-commerce / quick-commerce platform exports POs in its own column
layout and sometimes in PDF instead of Excel. These functions normalize
each into a common internal schema:

    platform_sku | order_qty_units | order_date | raw_product_name

`order_date` is the date the business wants to plan delivery against.
Where a platform's PDF gives no explicit per-line date, order_date is
left blank and the PO-level date (if present) is used instead.
"""

import re
import pandas as pd
import pdfplumber
import io


# ─────────────────────────────────────────────────────────────────────────
# Shared helpers
# ─────────────────────────────────────────────────────────────────────────

def _read_all_sheets(file_bytes) -> pd.DataFrame:
    """Reads every sheet in the uploaded workbook and concatenates them
    (Zepto-style exports often split rows across Sheet1/Sheet2 by warehouse)."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    return pd.concat(sheets.values(), ignore_index=True)


def _excel_serial_to_date(series: pd.Series) -> pd.Series:
    """Converts Excel serial date numbers to real dates; passes through
    already-parsed dates untouched."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def _clean_cell(s) -> str:
    """Strips all whitespace (including the embedded newlines pdfplumber
    leaves when a PDF cell's text wraps across lines)."""
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s))


def _clean_text_cell(s) -> str:
    """Like _clean_cell but keeps single spaces between words (for
    product descriptions, where 'Amul Salted Butter' shouldn't become
    'AmulSaltedButter')."""
    if s is None:
        return ""
    return " ".join(str(s).split())


def _extract_all_table_rows(file_bytes) -> list:
    """Runs pdfplumber table extraction across every page and returns a
    flat list of raw rows (each a list of cell strings/None)."""
    rows = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                rows.extend(table)
    return rows


def _extract_pdf_field(file_bytes, pattern: str) -> str:
    """Pulls a single labeled value (like 'PO No :- JCCPO00810') from the
    PDF's free text, for use as a PO reference on every output row."""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    m = re.search(pattern, full_text)
    return m.group(1) if m else ""


def _extract_pdf_date(file_bytes, pattern: str):
    val = _extract_pdf_field(file_bytes, pattern)
    if not val:
        return pd.NaT
    return pd.to_datetime(val, dayfirst=True, errors="coerce")


# ─────────────────────────────────────────────────────────────────────────
# ZEPTO — Excel format (sku GUID, po_qty, recommended_date)
# ─────────────────────────────────────────────────────────────────────────

def _ensure_standard_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Guarantees every reader's output has the full standard column set,
    even if that platform doesn't supply a particular field. Readers that
    give pieces only (Zepto, Swiggy, Blinkit) won't have order_qty_cases;
    Reliance, which states cases directly, will. Downstream code treats
    a missing/NaN order_qty_cases as "this platform doesn't supply cases
    directly — derive cases from pieces via the case-pack size instead."
    """
    if "order_qty_cases" not in df.columns:
        df["order_qty_cases"] = pd.NA
    return df


def read_zepto_po(file_bytes) -> pd.DataFrame:
    """
    Zepto Open PO export format (.xlsx).
    Key columns: sku (GUID), product_name, pack_size, uom, po_qty,
    recommended_date, wh_name, city_name, postatus.
    """
    df = _read_all_sheets(file_bytes)
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame({
        "platform_sku":      df["sku"].astype(str).str.strip().str.upper(),
        "raw_product_name":  df.get("product_name", ""),
        "order_qty_units":   pd.to_numeric(df.get("po_qty", df.get("open_qty")), errors="coerce"),
        "order_date":        _excel_serial_to_date(df["recommended_date"]),
        "warehouse":         df.get("wh_name", ""),
        "city":              df.get("city_name", ""),
        "po_status":         df.get("postatus", ""),
        "po_reference":      df.get("externpocode", ""),
    })
    out["order_qty_units"] = out["order_qty_units"].fillna(0)
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))


# ─────────────────────────────────────────────────────────────────────────
# SWIGGY (Jupiter Kart) — PDF GRN format
# Columns confirmed: SKU Code, SKU Desc, Exp Qty
# ─────────────────────────────────────────────────────────────────────────

def read_swiggy_jupiterkart_pdf(file_bytes) -> pd.DataFrame:
    """
    Jupiter Kart (Swiggy Instamart) GRN PDF.
    Table columns (by position, repeated on every page):
      0 Sr.No | 1 SKU Code | 2 SKU Desc | 3 Vendor SKU | 4 SKU Bin |
      5 Lot No | 6 Lot MRP | 7 Exp Qty | 8 Recv Qty | 9 Unit Price | ...

    Per your instruction: use SKU Code, SKU Desc, and Exp Qty (column
    index 7) — not Recv Qty. The same SKU can appear on multiple lines
    (different lot numbers) and is summed.
    """
    raw_rows = _extract_all_table_rows(file_bytes)

    data_rows = [r for r in raw_rows if r and r[0] and str(r[0]).strip().isdigit() and len(r) >= 8]

    if not data_rows:
        raise ValueError(
            "Couldn't find any GRN line items in this PDF. Expected a table with "
            "columns: Sr.No, SKU Code, SKU Desc, ..., Exp Qty. The file format may "
            "have changed — share it again so the reader can be adjusted."
        )

    out = pd.DataFrame({
        "platform_sku":     [_clean_cell(r[1]) for r in data_rows],
        "raw_product_name": [_clean_text_cell(r[2]) for r in data_rows],
        "order_qty_units":  [pd.to_numeric(_clean_cell(r[7]), errors="coerce") for r in data_rows],
    })
    out["order_qty_units"] = out["order_qty_units"].fillna(0)

    # Same SKU can appear on multiple lot lines — sum to one row per SKU
    out = (
        out.groupby("platform_sku", as_index=False)
        .agg(raw_product_name=("raw_product_name", "first"),
             order_qty_units=("order_qty_units", "sum"))
    )
    out["warehouse"] = ""
    out["city"] = ""
    out["po_status"] = ""
    out["po_reference"] = _extract_pdf_field(file_bytes, r"PO No\s*:-\s*(\S+)")
    out["order_date"] = _extract_pdf_date(file_bytes, r"GRN Date\s*:-\s*([\d-]+)")
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))


# ─────────────────────────────────────────────────────────────────────────
# BLINKIT (Zomato Hyperpure) — PDF Purchase Order format
# Columns confirmed: Item Code, Product Description, Qty.
# ─────────────────────────────────────────────────────────────────────────

def read_blinkit_zomato_hyperpure_pdf(file_bytes) -> pd.DataFrame:
    """
    Zomato Hyperpure (Blinkit) Purchase Order PDF.
    Table columns (16 total, repeated across pages):
      0 # | 1 Item Code | 2 HSN Code | 3 Product UPC | 4 Product Description |
      5 Basic Cost Price | 6 CGST% | 7 SGST% | 8 CESS% | 9 ADDT.CESS |
      10 Tax Amt | 11 Landing Rate | 12 Qty. | 13 MRP | 14 Margin% | 15 Total Amt

    Cell text wraps across lines in the PDF, so numeric cells like
    '100004\\n29' are rejoined into '10000429' before use.
    """
    raw_rows = _extract_all_table_rows(file_bytes)

    # Data rows are identified by col 0 being a plain row number AND the
    # row having the full 16-column width (filters out header/section rows)
    data_rows = [
        r for r in raw_rows
        if r and len(r) == 16 and r[0] and str(r[0]).strip().isdigit()
    ]

    if not data_rows:
        raise ValueError(
            "Couldn't find any PO line items in this PDF. Expected a 16-column table "
            "with '#', 'Item Code', 'Product Description', ..., 'Qty.' columns. "
            "The file format may have changed — share it again so the reader can be adjusted."
        )

    out = pd.DataFrame({
        "platform_sku":     [_clean_cell(r[1]) for r in data_rows],
        "raw_product_name": [_clean_text_cell(r[4]) for r in data_rows],
        "order_qty_units":  [pd.to_numeric(_clean_cell(r[12]), errors="coerce") for r in data_rows],
    })
    out["order_qty_units"] = out["order_qty_units"].fillna(0)
    out["warehouse"] = ""
    out["city"] = ""
    out["po_status"] = ""
    out["po_reference"] = _extract_pdf_field(file_bytes, r"P\.O\. Number\s*:\s*(\S+)")
    out["order_date"] = _extract_pdf_date(file_bytes, r"PO delivery\s*:\s*([A-Za-z]+ \d{1,2}, \d{4})")
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))


# ─────────────────────────────────────────────────────────────────────────
# RELIANCE — PDF (or Excel) Purchase Order format
# Columns confirmed: Article No. (first line of stacked cell), Material
# Description (first line of stacked cell), Quantity — where the upper
# stacked number is CASES and the lower is PIECES. Reliance states case
# quantity directly, so no case-pack division/rounding is needed for it
# the way it is for Zepto/Swiggy/Blinkit (which give pieces only).
# ─────────────────────────────────────────────────────────────────────────

def read_reliance_pdf(file_bytes) -> pd.DataFrame:
    """
    Reliance Retail Purchase Order PDF.
    Table columns (11 total, on the line-item page(s)):
      0 Sr.No | 1 Article No.\\nHSN Code | 2 EAN No... | 3 Material Description\\n
      Delivery Date\\nSite | 4 Quantity (cases\\npieces) | 5 UOM | 6 MRP | 7 Base Cost |
      8 CGST/SGST/CESS % | 9 CGST/SGST/CESS Amt | 10 Total Base Value

    Article No. is the FIRST line of column 1 (HSN code is the second
    line and is NOT part of the article number — they are two separate
    values stacked in one cell, not one number to split).
    Quantity column 4 has two stacked numbers: upper = cases, lower =
    pieces. Verified against Base Cost x cases = Total Base Value across
    multiple real rows.
    """
    raw_rows = _extract_all_table_rows(file_bytes)

    data_rows = [
        r for r in raw_rows
        if r and len(r) == 11 and r[0] and str(r[0]).strip().isdigit()
    ]

    if not data_rows:
        raise ValueError(
            "Couldn't find any PO line items in this PDF. Expected an 11-column table "
            "with 'Sr.No', 'Article No./HSN Code', 'Material Description', 'Quantity', "
            "etc. The file format may have changed — share it again so the reader can be adjusted."
        )

    def first_line(cell):
        return _clean_text_cell(str(cell).split("\n")[0]) if cell else ""

    def qty_lines(cell):
        """Returns (cases, pieces) from the stacked Quantity cell."""
        if not cell:
            return (0, 0)
        lines = [l.strip() for l in str(cell).split("\n") if l.strip()]
        cases = pd.to_numeric(lines[0], errors="coerce") if len(lines) > 0 else 0
        pieces = pd.to_numeric(lines[1], errors="coerce") if len(lines) > 1 else cases
        return (cases or 0, pieces or 0)

    qtys = [qty_lines(r[4]) for r in data_rows]

    out = pd.DataFrame({
        "platform_sku":      [_clean_cell(str(r[1]).split("\n")[0]) for r in data_rows],
        "raw_product_name":  [first_line(r[3]) for r in data_rows],
        "order_qty_cases":   [q[0] for q in qtys],
        "order_qty_units":   [q[1] for q in qtys],
    })
    out["order_qty_cases"] = pd.to_numeric(out["order_qty_cases"], errors="coerce").fillna(0)
    out["order_qty_units"] = pd.to_numeric(out["order_qty_units"], errors="coerce").fillna(0)
    out["warehouse"] = ""
    out["city"] = ""
    out["po_status"] = ""
    out["po_reference"] = _extract_pdf_field(file_bytes, r"PO NO\.\s*:\s*(\S+)")
    out["order_date"] = _extract_pdf_date(file_bytes, r"DELIVERY DATE\s*:\s*([\d.]+)")
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))


def read_reliance_excel(file_bytes, sku_col="Article No.", qty_cases_col="Quantity (Cases)",
                         qty_pieces_col=None, date_col="Delivery Date", name_col="Material Description",
                         po_ref_col="PO No.") -> pd.DataFrame:
    """
    Generic Excel reader for Reliance, in case they send an Excel export
    instead of PDF. Column names are configurable since we haven't seen
    a real Reliance Excel sample yet — adjust the defaults above once one
    is shared.
    """
    df = _read_all_sheets(file_bytes)
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame({
        "platform_sku":      df[sku_col].astype(str).str.strip().str.upper(),
        "raw_product_name":  df[name_col] if name_col in df.columns else "",
        "order_qty_cases":   pd.to_numeric(df[qty_cases_col], errors="coerce").fillna(0) if qty_cases_col in df.columns else 0,
        "order_qty_units":   pd.to_numeric(df[qty_pieces_col], errors="coerce").fillna(0) if qty_pieces_col and qty_pieces_col in df.columns else 0,
        "order_date":        _excel_serial_to_date(df[date_col]) if date_col in df.columns else pd.NaT,
        "warehouse":         "",
        "city":              "",
        "po_status":         "",
        "po_reference":      df[po_ref_col] if po_ref_col in df.columns else "",
    })
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))

def read_generic_po(file_bytes, sku_col, qty_col, date_col, name_col=None,
                     po_ref_col=None) -> pd.DataFrame:
    """Generic Excel reader for platforms with manually-configured column names."""
    df = _read_all_sheets(file_bytes)
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame({
        "platform_sku":      df[sku_col].astype(str).str.strip().str.upper(),
        "raw_product_name":  df[name_col] if name_col and name_col in df.columns else "",
        "order_qty_units":   pd.to_numeric(df[qty_col], errors="coerce").fillna(0),
        "order_date":        _excel_serial_to_date(df[date_col]) if date_col in df.columns else pd.NaT,
        "warehouse":         "",
        "city":              "",
        "po_status":         "",
        "po_reference":      df[po_ref_col] if po_ref_col and po_ref_col in df.columns else "",
    })
    return _ensure_standard_columns(out.dropna(subset=["platform_sku"]))


# ─────────────────────────────────────────────────────────────────────────
# Registry: maps platform name -> file type + reader function(s)
# ─────────────────────────────────────────────────────────────────────────
# "file_types" lists which upload formats are accepted for that platform,
# each with its own verified reader. A platform can accept more than one
# format (e.g. Reliance can send Excel OR PDF) once both are confirmed.

PLATFORM_READERS = {
    "Zepto": {
        "file_types": {
            "xlsx": {"reader": read_zepto_po, "verified": True},
        },
    },
    "Swiggy Instamart (Jupiter Kart)": {
        "file_types": {
            "pdf": {"reader": read_swiggy_jupiterkart_pdf, "verified": True},
        },
    },
    "Blinkit (Zomato Hyperpure)": {
        "file_types": {
            "pdf": {"reader": read_blinkit_zomato_hyperpure_pdf, "verified": True},
        },
    },
    "Reliance Retail": {
        "file_types": {
            "pdf": {"reader": read_reliance_pdf, "verified": True},
            "xlsx": {
                "reader": read_reliance_excel, "verified": False,
                "default_cols": {"sku_col": "Article No.", "qty_cases_col": "Quantity (Cases)",
                                  "qty_pieces_col": None, "date_col": "Delivery Date",
                                  "name_col": "Material Description", "po_ref_col": "PO No."},
            },
        },
    },
    "BigBasket": {
        "file_types": {
            "xlsx": {
                "reader": read_generic_po, "verified": False,
                "default_cols": {"sku_col": "Vendor SKU", "qty_col": "Units Ordered",
                                  "date_col": "Purchase Date", "name_col": "Product Description",
                                  "po_ref_col": "BB Order Ref"},
            },
        },
    },
}
