# --- START OF FILE config.py ---

# --- File Configuration ---
INPUT_EXCEL_FILE = "test.xlsx" # Or specific name for this format
# Specify sheet name, or None to use the active sheet
SHEET_NAME = None
# OUTPUT_PICKLE_FILE = "invoice_data.pkl" # Example for future use

# --- Sheet Parsing Configuration ---
# Row/Column range to search for the header
HEADER_SEARCH_ROW_RANGE = 25
HEADER_SEARCH_COL_RANGE = 25
# A pattern (string or regex) to identify a cell within the header row
# Updated pattern based on the SECOND header image
HEADER_IDENTIFICATION_PATTERN = r"批次号|订单号|物料代码|总张数|净重|毛重" # Adjusted pattern slightly

# Column mapping: Canonical Name -> List of possible header texts found in Excel (lowercase)
# Incorporating headers from BOTH images where possible, prioritizing user's core needs.
TARGET_HEADERS_MAP = {
    # --- Canonical Names Required by User (Core Logic) ---
    "po": ["po", "po no", "purchase order", "订单号"], # "订单号" present in both
    "item": ["item", "item no", "料号", "产品编号", "物料代码"], # "物料代码" from 2nd image is key
    "pcs": ["pcs", "张数", "数量", "qty", "件数", "总张数"], # Includes "总张数" (Total Sheets) from 2nd image. This is now likely the BASIS.
                                                            # Removed "出货数量 (sf)" from this mapping.
    "net": ["net", "net wt", "net weight", "净重"], # "净重" present in both
    "gross": ["gross", "gross wt", "gross weight", "毛重"], # "毛重" present in both
    "unit": ["unit", "unit price", "单价", "价格", "usd"], # Assuming 2nd 'USD' col is unit price
    "sqft": ["sqft", "出货数量 (sf)"], # Added new canonical name 'sqft' mapped to "出货数量 (sf)"

    # --- Canonical names required by user, but less certain based on headers ---
    "cbm": ["cbm", "meas", "measurement", "材积", "量码版"], # "量码版" is a possible but uncertain match
    "desc": ["desc", "description", "品名规格"], # No clear header in either image
    "inv_no": ["inv no", "invoice no", "发票号码"], # No clear header in either image
    "inv_date": ["inv date", "invoice date", "发票日期"], # No clear header in either image

    # --- Other Headers Found (Mapped to Consistent or New Canonical Names) ---
    "batch_no": ["批次号", "batch number"], # Added from 2nd image
    "line_no": ["行号", "line number", "line no"], # Added from 2nd image
    "direction": ["内向", "direction", "inward"], # Added from 2nd image (meaning unclear)
    "production_date": ["生产日期", "production date"], # Added from 2nd image
    "production_order_no": ["生产单号", "production order number"], # Present in both, potentially different types
    "reference_code": ["jlf/ttx编号", "ttx编号", "reference code"], # Combined codes from both images
    "level": ["级别", "等级", "level", "grade"], # Combined level/grade from both images
    "pallet_count": ["拖数", "pallet count"], # Renamed from towel_count based on likely meaning
    "manual_no": ["手册号", "manual number"], # From 1st image
    "remarks": ["备注", "remarks", "notes"], # Present in both
    "amount": ["金额", "amount"], # From 1st image

    # Removed "shipping_qty_sf" as it's now "sqft"
}

# --- Data Extraction Configuration ---
# 'item' (mapped from 物料代码 in 2nd image) still seems like a good choice.
STOP_EXTRACTION_ON_EMPTY_COLUMN = 'item'
# Safety limit for the number of data rows to read below the header
MAX_DATA_ROWS_TO_SCAN = 1000

# --- Data Processing Configuration ---
# List of canonical header names for columns where values should be distributed
COLUMNS_TO_DISTRIBUTE = ["net", "gross"] # Still likely just these two. 'cbm' is uncertain.

# The canonical header name of the column used for proportional distribution (e.g., quantity)
# Reverted to 'pcs' which maps to "总张数" (Total Sheets), assuming this is the quantity basis.
DISTRIBUTION_BASIS_COLUMN = "pcs"

# --- END OF FILE config.py ---