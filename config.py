# --- START OF FILE config.py ---

# --- START OF FULL FILE: config.py ---

# --- File Configuration ---
INPUT_EXCEL_FILE = "JF.xlsx" # Or specific name for this format, e.g., "JF_Data_2024.xlsx"
# Specify sheet name, or None to use the active sheet
SHEET_NAME = None
# OUTPUT_PICKLE_FILE = "invoice_data.pkl" # Example for future use

# --- Sheet Parsing Configuration ---
# Row/Column range to search for the header
HEADER_SEARCH_ROW_RANGE = 50 # Increased range slightly, adjust if headers can be lower
HEADER_SEARCH_COL_RANGE = 30 # Increased range slightly, adjust if many columns
# A pattern (string or regex) to identify a cell within the header row
# This pattern helps find *any* header row, the mapping below specifies exact matches
HEADER_IDENTIFICATION_PATTERN = r"批次号|订单号|物料代码|总张数|净重|毛重|po|item|pcs|net|gross" # Broadened slightly

# --- Column Mapping Configuration ---
# Canonical Name -> List containing header variations (case-insensitive match)
TARGET_HEADERS_MAP = {
    # --- Core Logic Canonical Names ---
    "po": ["po", "订单号"],                 # Primary English: 'po', Primary Chinese: '订单号'
    "item": ["item no", "物料代码", 'item', "产品编号"],        # Primary English: 'item no', Primary Chinese: '物料代码'
    "pcs": ["pcs", "总张数", "张数"],                # Primary English: 'pcs', Primary Chinese: '总张数'
    "net": ["net weight", "净重", "net"],          # Primary English: 'net weight', Primary Chinese: '净重'
    "gross": ["gross weight", "毛重", "gross", ],       # Primary English: 'gross weight', Primary Chinese: '毛重'
    "unit": ["unit price", "单价", "price", "unit"],          # Primary English: 'unit price', Primary Chinese: '单价'
    "sqft": ["sqft", "出货数量 (sf)", "尺数"],      # Primary English: 'sqft', Primary Chinese: '出货数量 (sf)' (Assuming this specific text)
    "amount": ["amount", "金额", "总价"],            # Primary English: 'amount', Primary Chinese: '金额' # Ensure this is present and mapped

    # --- Less Certain Canonical Names ---
    "cbm": ["cbm", "材积"],                # Primary English: 'cbm', Primary Chinese: '材积' (Verify '材积' is correct/common)
    "description": ["description", "品名规格", "产品编号"],      # Primary English: 'description', Primary Chinese: '品名规格'
    "inv_no": ["invoice no", "发票号码"],    # Primary English: 'invoice no', Primary Chinese: '发票号码'
    "inv_date": ["invoice date", "发票日期"], # Primary English: 'invoice date', Primary Chinese: '发票日期'
    "inv_ref": ["ref", "invoice ref", "ref no"],

    # --- Other Found Headers ---
    "dc": ["批次号", "DC", "dc"],
    "batch_no": ["batch number", "批次号"],  # Primary English: 'batch number', Primary Chinese: '批次号'
    "line_no": ["line no", "行号"],           # Primary English: 'line no', Primary Chinese: '行号'
    "direction": ["direction", "内向"],      # Primary English: 'direction', Primary Chinese: '内向' (Meaning still unclear)
    "production_date": ["production date", "生产日期"], # Primary English: 'production date', Primary Chinese: '生产日期'
    "production_order_no": ["production order number", "生产单号"], # Primary English: 'production order number', Primary Chinese: '生产单号'
    "reference_code": ["reference code", "ttx编号"], # Primary English: 'reference code', Primary Chinese: 'ttx编号' (Verify 'ttx编号')
    "level": ["grade", "等级"],              # Primary English: 'grade', Primary Chinese: '等级'
    "pallet_count": ["pallet count", "拖数", "PALLET", "件数"],# Primary English: 'pallet count', Primary Chinese: '拖数'
    "manual_no": ["manual number", "手册号"], # Primary English: 'manual number', Primary Chinese: '手册号'
    "remarks": ["remarks", "备注", "Remark"],          # Primary English: 'remarks', Primary Chinese: '备注'
    # 'amount' is already defined above

    # Add any other essential headers here following the variations list format
}

# --- Data Extraction Configuration ---
# Choose a column likely to be empty *only* when the data rows truly end.
# 'item' is often a good candidate if item codes are always present for data rows.
STOP_EXTRACTION_ON_EMPTY_COLUMN = 'item'
# Safety limit for the number of data rows to read below the header within a table
MAX_DATA_ROWS_TO_SCAN = 1000

# --- Data Processing Configuration ---
# List of canonical header names for columns where values should be distributed
# CBM processing/distribution depends on the 'cbm' mapping above and if the column contains L*W*H strings
COLUMNS_TO_DISTRIBUTE = ["net", "gross", "cbm"] # Include 'cbm' if you want to distribute calculated CBM values

# The canonical header name of the column used for proportional distribution
DISTRIBUTION_BASIS_COLUMN = "pcs"

# --- Aggregation Strategy Configuration ---
# List or Tuple of *workbook filename* prefixes (case-sensitive) that trigger CUSTOM aggregation.
# Custom aggregation sums 'sqft' and 'amount' based ONLY on 'po' and 'item'.
# Standard aggregation sums 'sqft' based on 'po', 'item', and 'unit'.
# Example: If INPUT_EXCEL_FILE is "JF_Report_Q1.xlsx", it will match "JF".
CUSTOM_AGGREGATION_WORKBOOK_PREFIXES = ("JF", "MOTO") # Renamed Variable


# --- END OF FULL FILE: config.py ---