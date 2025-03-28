# --- START OF FULL FILE: sheet_parser.py ---

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
    DISTRIBUTION_BASIS_COLUMN, # Ensure these are available
    COLUMNS_TO_DISTRIBUTE     # Ensure these are available
)

def find_all_header_rows(sheet, search_pattern, row_range, col_range) -> List[int]:
    """
    Finds all 1-indexed row numbers containing a header based on a pattern.
    Returns a list of row numbers, sorted in ascending order.
    """
    header_rows: List[int] = []
    try:
        # Compile the regex pattern once
        regex = re.compile(search_pattern, re.IGNORECASE)
        # Determine search boundaries, ensuring they don't exceed sheet dimensions
        max_row_to_search = min(row_range, sheet.max_row)
        max_col_to_search = min(col_range, sheet.max_column)

        logging.info(f"Searching for headers using pattern '{search_pattern}' in rows 1-{max_row_to_search}, cols 1-{max_col_to_search}")

        # Iterate through the specified range to find header cells
        for r_idx in range(1, max_row_to_search + 1):
            # Optimization: Check only necessary columns if pattern is specific
            for c_idx in range(1, max_col_to_search + 1):
                cell = sheet.cell(row=r_idx, column=c_idx)
                if cell.value is not None:
                    cell_value_str = str(cell.value).strip()
                    # If the cell content matches the pattern, consider this a header row
                    if regex.search(cell_value_str):
                        logging.debug(f"Header pattern found in cell {cell.coordinate} (Row: {r_idx}). Adding row to list.")
                        header_rows.append(r_idx)
                        # Once a header is found in a row, move to the next row
                        break

        # Sort the found header rows
        header_rows.sort()

        if not header_rows:
            logging.warning(f"Header pattern '{search_pattern}' not found within the search range.")
        else:
            logging.info(f"Found {len(header_rows)} potential header rows at: {header_rows}")

        return header_rows
    except Exception as e:
        logging.error(f"Error finding header rows: {e}", exc_info=True)
        return []

def map_columns_to_headers(sheet, header_row: int, col_range: int) -> Dict[str, int]:
    """
    Maps canonical header names to their 1-indexed column numbers based on the
    header row content, prioritizing the first match found based on TARGET_HEADERS_MAP order.
    (Uses the variation -> canonical lookup for clarity)

    Args:
        sheet: The openpyxl worksheet object.
        header_row: The 1-indexed row number containing the headers.
        col_range: The maximum number of columns to search for headers.

    Returns:
        A dictionary mapping canonical names (str) to column indices (int).
    """
    if header_row is None or header_row < 1:
        logging.error("Invalid header_row provided for column mapping.")
        return {}

    column_mapping: Dict[str, int] = {}
    processed_canonicals = set() # Track canonical names already assigned to a column
    max_col_to_check = min(col_range, sheet.max_column)

    logging.info(f"Mapping columns based on header row {header_row} up to column {max_col_to_check}.")

    # Build a reverse lookup: lowercase variation -> canonical name
    variation_to_canonical_lookup: Dict[str, str] = {}
    ambiguous_variations = set()
    for canonical_name, variations in TARGET_HEADERS_MAP.items():
        for variation in variations:
            variation_lower = str(variation).lower().strip()
            if not variation_lower: continue

            if variation_lower in variation_to_canonical_lookup and variation_to_canonical_lookup[variation_lower] != canonical_name:
                 if variation_lower not in ambiguous_variations:
                      logging.warning(f"Config Issue: Header variation '{variation_lower}' mapped to multiple canonical names ('{variation_to_canonical_lookup[variation_lower]}', '{canonical_name}', etc.). Check TARGET_HEADERS_MAP.")
                      ambiguous_variations.add(variation_lower)
            variation_to_canonical_lookup[variation_lower] = canonical_name

    # Iterate through Excel columns and map using the lookup
    for col_idx in range(1, max_col_to_check + 1):
        cell = sheet.cell(row=header_row, column=col_idx)
        actual_header_text = str(cell.value).lower().strip() if cell.value is not None else ""

        if not actual_header_text:
            continue

        matched_canonical = variation_to_canonical_lookup.get(actual_header_text)

        if matched_canonical:
            if matched_canonical not in processed_canonicals:
                column_mapping[matched_canonical] = col_idx
                processed_canonicals.add(matched_canonical)
                logging.info(f"Mapped column {col_idx} ('{cell.value}') -> '{matched_canonical}'")
            else:
                logging.warning(f"Duplicate Header/Mapping: Canonical name '{matched_canonical}' (from Excel header '{cell.value}' in Col {col_idx}) was already mapped to Col {column_mapping.get(matched_canonical)}. Ignoring this duplicate column.")

    if not column_mapping:
        logging.warning(f"No target headers were successfully mapped in row {header_row}. Check Excel headers and TARGET_HEADERS_MAP.")
    else:
        # Verify essential columns needed for later processing are mapped
        required = set()
        if DISTRIBUTION_BASIS_COLUMN:
            required.add(DISTRIBUTION_BASIS_COLUMN)
        if COLUMNS_TO_DISTRIBUTE:
            required.update(COLUMNS_TO_DISTRIBUTE)

        missing = required - set(column_mapping.keys())
        if missing:
             logging.warning(f"Missing required header mappings needed for processing: {missing}. Processing might fail or be incomplete.")

    return column_mapping


def extract_multiple_tables(sheet, header_rows: List[int], column_mapping: Dict[str, int]) -> Dict[int, Dict[str, List[Any]]]:
    """
    Extracts data for multiple tables defined by header_rows.

    Args:
        sheet: The openpyxl worksheet object.
        header_rows: A sorted list of 1-indexed header row numbers.
        column_mapping: A dictionary mapping canonical header names to 1-indexed column numbers.

    Returns:
        A dictionary where keys are table indices (1, 2, 3...) and values are
        dictionaries representing each table's data ({'header': [values...]}).
    """
    if not header_rows:
        logging.warning("No header rows provided, cannot extract tables.")
        return {}
    if not column_mapping:
        logging.error("Column mapping is empty, cannot extract data meaningfully.")
        return {}

    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}
    stop_col_idx = column_mapping.get(STOP_EXTRACTION_ON_EMPTY_COLUMN) if STOP_EXTRACTION_ON_EMPTY_COLUMN else None

    if STOP_EXTRACTION_ON_EMPTY_COLUMN and not stop_col_idx:
        logging.warning(f"Stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' is configured but not found in column mapping. Extraction will rely on MAX_DATA_ROWS_TO_SCAN or next header.")

    # Iterate through each identified header row to define table boundaries
    for i, header_row in enumerate(header_rows):
        table_index = i + 1
        start_data_row = header_row + 1

        # Determine the end row for the current table's data
        if i + 1 < len(header_rows):
            max_possible_end_row = header_rows[i + 1] # End before the next header
        else:
            max_possible_end_row = sheet.max_row + 1 # Last table, go to sheet end

        # Apply MAX_DATA_ROWS_TO_SCAN limit relative to start_data_row
        scan_limit_row = start_data_row + MAX_DATA_ROWS_TO_SCAN
        # Actual end row is the minimum of the next header, scan limit, and sheet max row + 1
        end_data_row = min(max_possible_end_row, scan_limit_row)

        logging.info(f"Extracting data for Table {table_index} (Header Row: {header_row}, Data Rows: {start_data_row} to {end_data_row - 1})")

        current_table_data: Dict[str, List[Any]] = {key: [] for key in column_mapping.keys()}
        rows_extracted_for_table = 0

        # Extract data row by row for the current table
        for current_row in range(start_data_row, end_data_row):

            # Check stopping condition based on designated empty column
            if stop_col_idx:
                stop_cell = sheet.cell(row=current_row, column=stop_col_idx)
                stop_cell_value = stop_cell.value
                # Consider empty if None or an empty string after stripping
                if stop_cell_value is None or (isinstance(stop_cell_value, str) and not stop_cell_value.strip()):
                    logging.info(f"Stopping extraction for Table {table_index} at row {current_row}: Empty cell found in stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' (Col {stop_col_idx}).")
                    break # Stop processing rows for *this* table

            # Extract data for all mapped columns in this row
            for header, col_idx in column_mapping.items():
                cell = sheet.cell(row=current_row, column=col_idx)
                cell_value = cell.value
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                current_table_data[header].append(cell_value)

            rows_extracted_for_table += 1

        # Log if MAX_DATA_ROWS_TO_SCAN limit was hit for this table
        # Check if the loop actually finished *because* it hit the limit row exactly
        if current_row == scan_limit_row - 1 and rows_extracted_for_table >= MAX_DATA_ROWS_TO_SCAN:
             logging.warning(f"Reached MAX_DATA_ROWS_TO_SCAN limit ({MAX_DATA_ROWS_TO_SCAN}) for Table {table_index}. Extraction might be incomplete for this table.")

        # Store the extracted data for the current table using its index
        if rows_extracted_for_table > 0:
            all_tables_data[table_index] = current_table_data
            logging.info(f"Finished extracting {rows_extracted_for_table} rows for Table {table_index}.")
        else:
            logging.info(f"No data rows extracted for Table {table_index} (between row {start_data_row} and {end_data_row}).")


    logging.info(f"Completed extraction. Found data for {len(all_tables_data)} tables.")
    return all_tables_data

# --- END OF FULL FILE: sheet_parser.py ---