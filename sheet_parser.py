# --- START OF FILE sheet_parser.py ---

import re
import logging
from typing import Dict, List, Optional, Tuple, Any # For type hinting

# Import config values (consider passing them as arguments for more flexibility)
from config import (
    TARGET_HEADERS_MAP,
    HEADER_SEARCH_ROW_RANGE,
    HEADER_SEARCH_COL_RANGE,
    HEADER_IDENTIFICATION_PATTERN,
    STOP_EXTRACTION_ON_EMPTY_COLUMN,
    MAX_DATA_ROWS_TO_SCAN,
    # --- ADDED MISSING IMPORTS ---
    DISTRIBUTION_BASIS_COLUMN,
    COLUMNS_TO_DISTRIBUTE
    # -----------------------------
)

def find_header_row(sheet, search_pattern, row_range, col_range) -> Optional[int]:
    """
    Finds the 1-indexed row number containing a header based on a pattern.
    """
    try:
        regex = re.compile(search_pattern, re.IGNORECASE)
        max_row = min(row_range, sheet.max_row)
        max_col = min(col_range, sheet.max_column)

        for r_idx in range(1, max_row + 1):
            for c_idx in range(1, max_col + 1):
                cell = sheet.cell(row=r_idx, column=c_idx)
                if cell.value is not None:
                    cell_value_str = str(cell.value).strip()
                    if regex.search(cell_value_str):
                        logging.info(f"Header pattern '{search_pattern}' found in cell {cell.coordinate} (Row: {r_idx}).")
                        return r_idx
        logging.warning(f"Header pattern '{search_pattern}' not found within search range (Rows: 1-{max_row}, Cols: 1-{max_col}).")
        return None
    except Exception as e:
        logging.error(f"Error finding header row: {e}", exc_info=True)
        return None

def map_columns_to_headers(sheet, header_row: int, col_range: int) -> Dict[str, int]:
    """
    Maps canonical header names to their 1-indexed column numbers based on the header row content.
    """
    if header_row is None or header_row < 1:
        return {}

    column_mapping: Dict[str, int] = {}
    found_canonical_headers = set()
    max_col_to_check = min(col_range, sheet.max_column)

    logging.info(f"Mapping columns in header row {header_row} up to column {max_col_to_check}.")

    for col_idx in range(1, max_col_to_check + 1):
        cell = sheet.cell(row=header_row, column=col_idx)
        actual_header_text = str(cell.value).strip().lower() if cell.value is not None else ""

        if not actual_header_text:
            continue

        # Find the best match in TARGET_HEADERS_MAP
        best_match_canonical = None
        for canonical_name, variations in TARGET_HEADERS_MAP.items():
            # Optimization: Check if we already found this canonical name
            if canonical_name in column_mapping: # Check column_mapping instead of found_canonical_headers for stricter single mapping
                continue

            if actual_header_text in variations:
                 best_match_canonical = canonical_name
                 break # Found a match for this cell

        if best_match_canonical:
             # Check if this column index is already mapped to something else (less likely but possible)
             # This check is less crucial than checking if the canonical name is already mapped.
             # if col_idx in column_mapping.values():
             #    logging.warning(f"Column index {col_idx} is already mapped. Check header mapping logic.")

             column_mapping[best_match_canonical] = col_idx
             # found_canonical_headers.add(best_match_canonical) # Not strictly needed if checking column_mapping above
             logging.info(f"Mapped column {col_idx} ('{cell.value}') -> '{best_match_canonical}'")

    if not column_mapping:
        logging.warning(f"No target headers found or mapped in row {header_row}.")
    else:
        # Verify essential columns are mapped (optional but recommended)
        # This check now uses the imported variables
        required = {DISTRIBUTION_BASIS_COLUMN} | set(COLUMNS_TO_DISTRIBUTE)
        missing = required - set(column_mapping.keys())
        if missing:
             logging.warning(f"Missing required header mappings needed for distribution: {missing}")

    return column_mapping

def extract_raw_data(sheet, header_row: int, column_mapping: Dict[str, int]) -> Dict[str, List[Any]]:
    """
    Extracts data from rows below the header into a dictionary of lists.
    """
    if not column_mapping or header_row is None:
        logging.error("Cannot extract data without valid header row and column mapping.")
        # Return empty lists for all *expected* headers, not just mapped ones
        return {key: [] for key in TARGET_HEADERS_MAP.keys()}

    # Initialize raw_data only for the columns that were actually mapped
    raw_data: Dict[str, List[Any]] = {key: [] for key in column_mapping.keys()}
    start_data_row = header_row + 1
    stop_col_idx = column_mapping.get(STOP_EXTRACTION_ON_EMPTY_COLUMN) if STOP_EXTRACTION_ON_EMPTY_COLUMN else None

    if STOP_EXTRACTION_ON_EMPTY_COLUMN and not stop_col_idx:
        logging.warning(f"Stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' not found in mapping. Relying on MAX_DATA_ROWS_TO_SCAN.")

    logging.info(f"Starting data extraction from row {start_data_row}.")
    rows_processed = 0
    max_extraction_row = min(start_data_row + MAX_DATA_ROWS_TO_SCAN, sheet.max_row + 1)

    for current_row in range(start_data_row, max_extraction_row):
        rows_processed += 1

        # Check stopping condition based on designated column
        if stop_col_idx:
            stop_cell = sheet.cell(row=current_row, column=stop_col_idx)
            # More robust check for empty: None or empty string after stripping
            stop_cell_value_str = str(stop_cell.value).strip() if stop_cell.value is not None else ""
            if not stop_cell_value_str:
                logging.info(f"Stopping extraction at row {current_row}: Empty cell in stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' (Col {stop_col_idx}).")
                break

        # Extract data for mapped columns in this row
        for header, col_idx in column_mapping.items():
            cell = sheet.cell(row=current_row, column=col_idx)
            cell_value = cell.value
            # Basic cleaning: strip whitespace if string
            if isinstance(cell_value, str):
                cell_value = cell_value.strip()
            # Ensure the list exists before appending (should always exist here)
            raw_data[header].append(cell_value)

    # Check if loop finished due to reaching max rows
    if current_row == max_extraction_row - 1 and rows_processed >= MAX_DATA_ROWS_TO_SCAN :
         logging.warning(f"Reached MAX_DATA_ROWS_TO_SCAN limit ({MAX_DATA_ROWS_TO_SCAN}). Extraction might be incomplete.")

    # Get length based on a reliably mapped column if possible, otherwise handle empty dict
    data_rows_extracted = 0
    if raw_data:
        # Try getting length from the basis column if mapped, otherwise first key
        basis_key = DISTRIBUTION_BASIS_COLUMN if DISTRIBUTION_BASIS_COLUMN in raw_data else next(iter(raw_data.keys()), None)
        if basis_key:
            data_rows_extracted = len(raw_data.get(basis_key, []))

    logging.info(f"Extracted raw data for {data_rows_extracted} rows.")
    return raw_data

# --- END OF FILE sheet_parser.py ---