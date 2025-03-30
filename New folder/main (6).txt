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

        logging.info(f"[find_all_header_rows] Searching for headers using pattern '{search_pattern}' in rows 1-{max_row_to_search}, cols 1-{max_col_to_search}")

        # Iterate through the specified range to find header cells
        for r_idx in range(1, max_row_to_search + 1):
            # Optimization: Check only necessary columns if pattern is specific
            for c_idx in range(1, max_col_to_search + 1):
                cell = sheet.cell(row=r_idx, column=c_idx)
                if cell.value is not None:
                    cell_value_str = str(cell.value).strip()
                    # If the cell content matches the pattern, consider this a header row
                    if regex.search(cell_value_str):
                        logging.debug(f"[find_all_header_rows] Header pattern found in cell {cell.coordinate} (Row: {r_idx}). Adding row to list.")
                        # Check if this row is already added to prevent duplicates if multiple cells match in the same row
                        if r_idx not in header_rows:
                            header_rows.append(r_idx)
                        # Once a header is found in a row, move to the next row
                        break # Break inner column loop

        # Sort the found header rows
        header_rows.sort()

        if not header_rows:
            logging.warning(f"[find_all_header_rows] Header pattern '{search_pattern}' not found within the search range.")
        else:
            # Log the final list found AT INFO level for clarity
            logging.info(f"[find_all_header_rows] Found {len(header_rows)} potential header rows at: {header_rows}")

        return header_rows
    except Exception as e:
        logging.error(f"[find_all_header_rows] Error finding header rows: {e}", exc_info=True)
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
        logging.error("[map_columns_to_headers] Invalid header_row provided for column mapping.")
        return {}

    column_mapping: Dict[str, int] = {}
    processed_canonicals = set() # Track canonical names already assigned to a column
    max_col_to_check = min(col_range, sheet.max_column)

    logging.info(f"[map_columns_to_headers] Mapping columns based on header row {header_row} up to column {max_col_to_check}.")

    # Build a reverse lookup: lowercase variation -> canonical name
    variation_to_canonical_lookup: Dict[str, str] = {}
    ambiguous_variations = set()
    for canonical_name, variations in TARGET_HEADERS_MAP.items():
        # Ensure variations is iterable, even if it's a single string
        if isinstance(variations, str):
            variations = [variations] # Treat single string as a list with one item
        elif not hasattr(variations, '__iter__'): # Check if it's iterable but not string
             logging.error(f"[map_columns_to_headers] Config Error: Value for canonical name '{canonical_name}' in TARGET_HEADERS_MAP is not a list or string: {variations}. Skipping this canonical.")
             continue

        for variation in variations:
            variation_lower = str(variation).lower().strip()
            if not variation_lower: continue

            if variation_lower in variation_to_canonical_lookup and variation_to_canonical_lookup[variation_lower] != canonical_name:
                 if variation_lower not in ambiguous_variations:
                      logging.warning(f"[map_columns_to_headers] Config Issue: Header variation '{variation_lower}' mapped to multiple canonical names ('{variation_to_canonical_lookup[variation_lower]}', '{canonical_name}', etc.). Check TARGET_HEADERS_MAP. Using first encountered mapping.")
                      ambiguous_variations.add(variation_lower)
                 # Decide on behavior: either overwrite or keep first. Keeping first based on warning.
                 # variation_to_canonical_lookup[variation_lower] = canonical_name # <-- would overwrite
            else:
                 variation_to_canonical_lookup[variation_lower] = canonical_name


    # Iterate through Excel columns and map using the lookup
    for col_idx in range(1, max_col_to_check + 1):
        cell = sheet.cell(row=header_row, column=col_idx)
        # Use .value directly; openpyxl handles data types
        cell_value = cell.value
        actual_header_text = str(cell_value).lower().strip() if cell_value is not None else ""

        if not actual_header_text:
            # Log empty header cells at DEBUG level
            logging.debug(f"[map_columns_to_headers] Cell {cell.coordinate} in header row {header_row} is empty or None.")
            continue

        matched_canonical = variation_to_canonical_lookup.get(actual_header_text)

        if matched_canonical:
            if matched_canonical not in processed_canonicals:
                column_mapping[matched_canonical] = col_idx
                processed_canonicals.add(matched_canonical)
                # Log successful mapping at INFO level
                logging.info(f"[map_columns_to_headers] Mapped column {col_idx} (Header Text: '{cell.value}') -> Canonical: '{matched_canonical}'")
            else:
                # Log duplicate canonical mapping as warning
                logging.warning(f"[map_columns_to_headers] Duplicate Canonical Mapping: Canonical name '{matched_canonical}' (from Excel header '{cell.value}' in Col {col_idx}) was already mapped to Col {column_mapping.get(matched_canonical)}. Ignoring this duplicate column for '{matched_canonical}'.")
        else:
             # Log headers found in Excel but not matching any variation at DEBUG level
             logging.debug(f"[map_columns_to_headers] Excel header '{cell.value}' (Col {col_idx}) in row {header_row} did not match any known variations in TARGET_HEADERS_MAP.")


    if not column_mapping:
        logging.warning(f"[map_columns_to_headers] No target headers were successfully mapped in row {header_row}. Check Excel headers and TARGET_HEADERS_MAP content.")
    else:
        # Check for essential columns needed later
        required = set()
        if DISTRIBUTION_BASIS_COLUMN:
            required.add(DISTRIBUTION_BASIS_COLUMN)
        if COLUMNS_TO_DISTRIBUTE:
            required.update(COLUMNS_TO_DISTRIBUTE)
        # Also check essentials for SQFT aggregation if known
        required.update(['po', 'item', 'unit', 'sqft'])

        missing = required - set(column_mapping.keys())
        if missing:
             # This is important, log as WARNING
             logging.warning(f"[map_columns_to_headers] Mapping complete for row {header_row}, but MISSING essential canonical mappings needed for processing: {missing}. Subsequent steps might fail or be incomplete.")
        else:
             logging.info(f"[map_columns_to_headers] All essential columns ({required}) appear to be mapped successfully for header row {header_row}.")


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
        logging.warning("[extract_multiple_tables] No header rows provided, cannot extract tables.")
        return {}
    if not column_mapping:
        logging.error("[extract_multiple_tables] Column mapping is empty, cannot extract data meaningfully.")
        return {}

    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}
    stop_col_idx = column_mapping.get(STOP_EXTRACTION_ON_EMPTY_COLUMN) if STOP_EXTRACTION_ON_EMPTY_COLUMN else None
    prefix = "[extract_multiple_tables]" # Log prefix

    logging.info(f"{prefix} Starting extraction for {len(header_rows)} identified header(s): {header_rows}")

    if STOP_EXTRACTION_ON_EMPTY_COLUMN:
        if stop_col_idx:
             logging.info(f"{prefix} Will stop reading rows within a table if column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' (Index: {stop_col_idx}) is empty.")
        else:
            logging.warning(f"{prefix} Stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' is configured but was not found in the column mapping. Extraction will rely solely on MAX_DATA_ROWS_TO_SCAN or the next header row.")

    # Iterate through each identified header row to define table boundaries
    for i, header_row in enumerate(header_rows):
        table_index = i + 1
        logging.info(f"{prefix} >>> Processing Header Row {header_row} as Table Index {table_index}")

        start_data_row = header_row + 1

        # Determine the end row for the current table's data
        if i + 1 < len(header_rows):
            # End before the next header row starts
            max_possible_end_row = header_rows[i + 1]
            logging.debug(f"{prefix} Table {table_index}: Next header found at row {max_possible_end_row}. Data extraction will stop before this row.")
        else:
            # Last table, potential end is sheet max row + 1
            max_possible_end_row = sheet.max_row + 1
            logging.debug(f"{prefix} Table {table_index}: This is the last header. Max possible end row: {max_possible_end_row} (Sheet max_row: {sheet.max_row})")

        # Apply MAX_DATA_ROWS_TO_SCAN limit relative to the start_data_row
        scan_limit_row = start_data_row + MAX_DATA_ROWS_TO_SCAN
        # Actual end row is the minimum of the next header, scan limit, and sheet end (+1)
        end_data_row = min(max_possible_end_row, scan_limit_row)

        # Adjust if start row is already beyond end row (e.g., two headers immediately adjacent)
        if start_data_row >= end_data_row:
            logging.warning(f"{prefix} Table {table_index}: Start data row ({start_data_row}) is not before calculated end data row ({end_data_row}). No data rows will be extracted for this table.")
            # Store empty structure and continue to next header
            all_tables_data[table_index] = {key: [] for key in column_mapping.keys()}
            continue


        logging.info(f"{prefix} Table {table_index}: Extracting Data Rows {start_data_row} to {end_data_row - 1} (Header: {header_row}, LimitNextHeader: {max_possible_end_row}, LimitScan: {scan_limit_row})")

        current_table_data: Dict[str, List[Any]] = {key: [] for key in column_mapping.keys()}
        rows_extracted_for_table = 0
        last_row_processed = start_data_row - 1 # Track the last row index processed
        stop_condition_met = False # Flag if stop column caused early exit

        # Extract data row by row for the current table
        for current_row in range(start_data_row, end_data_row):
            last_row_processed = current_row # Update last processed row

            # Check stopping condition based on designated empty column
            if stop_col_idx:
                stop_cell = sheet.cell(row=current_row, column=stop_col_idx)
                stop_cell_value = stop_cell.value
                # Consider empty if None or an empty string after stripping
                is_empty = stop_cell_value is None or (isinstance(stop_cell_value, str) and not stop_cell_value.strip())
                if is_empty:
                    logging.info(f"{prefix} Stopping extraction for Table {table_index} at row {current_row}: Empty cell found in stop column '{STOP_EXTRACTION_ON_EMPTY_COLUMN}' (Col {stop_col_idx}).")
                    stop_condition_met = True
                    break # Stop processing rows for *this* table

            # Extract data for all mapped columns in this row
            row_has_data = False # Check if the row has any data at all in mapped columns
            logging.debug(f"{prefix} Table {table_index}, Reading row {current_row}:") # Row-level debug
            for header, col_idx in column_mapping.items():
                cell = sheet.cell(row=current_row, column=col_idx)
                cell_value = cell.value # Get value using openpyxl's type handling
                # Strip leading/trailing whitespace from strings ONLY
                if isinstance(cell_value, str):
                    processed_value = cell_value.strip()
                else:
                    processed_value = cell_value # Keep numbers, dates, None, etc. as is

                current_table_data[header].append(processed_value)
                logging.debug(f"{prefix}   Col '{header}' ({col_idx}): Value='{processed_value}' (Type: {type(processed_value).__name__})") # Cell-level debug
                # Check if this specific cell has meaningful data
                if processed_value is not None and processed_value != "":
                    row_has_data = True

            # Log if a row seems entirely empty across mapped columns
            if not row_has_data:
                logging.debug(f"{prefix} Table {table_index}, Row {current_row}: No data found in any mapped columns for this row.")
                # Decide if you want to STOP on a fully empty row (could be risky if there are intentional gaps)
                # if STOP_ON_FULLY_EMPTY_ROW_CONFIG: break

            rows_extracted_for_table += 1

        # Log if MAX_DATA_ROWS_TO_SCAN limit was hit
        # This happens if the loop finished *and* the last row processed was the limit boundary
        # *and* the stop condition wasn't the reason for finishing early.
        if not stop_condition_met and last_row_processed == scan_limit_row - 1 and rows_extracted_for_table >= MAX_DATA_ROWS_TO_SCAN:
             logging.warning(f"{prefix} Reached MAX_DATA_ROWS_TO_SCAN limit ({MAX_DATA_ROWS_TO_SCAN}) for Table {table_index} at row {last_row_processed}. Extraction might be incomplete for this table.")

        # --- Store results ---
        logging.debug(f"{prefix} Finished row scanning loop for Table {table_index}. Rows processed in loop: {rows_extracted_for_table}.")
        if rows_extracted_for_table > 0:
            # Verify list lengths (should always match if extraction logic is correct)
            list_lengths = {hdr: len(lst) for hdr, lst in current_table_data.items()}
            if len(set(list_lengths.values())) > 1:
                 logging.error(f"{prefix} !!! Internal Error: List lengths inconsistent after extracting Table {table_index}. Lengths: {list_lengths}. Storing data as is, but review logic.")
            elif list(list_lengths.values())[0] != rows_extracted_for_table:
                 logging.error(f"{prefix} !!! Internal Error: List length ({list(list_lengths.values())[0]}) does not match rows extracted count ({rows_extracted_for_table}) for Table {table_index}. Storing data, but review logic.")

            all_tables_data[table_index] = current_table_data
            logging.info(f"{prefix} Successfully stored {rows_extracted_for_table} rows of data for Table Index {table_index} in the results dictionary.")
            logging.debug(f"{prefix} Current keys in all_tables_data after adding Table {table_index}: {list(all_tables_data.keys())}")
        else:
            # Store empty dict even if no rows found
            all_tables_data[table_index] = current_table_data # Contains empty lists
            logging.info(f"{prefix} No data rows were extracted or met criteria for Table {table_index} (Header row {header_row}). Storing empty structure for this table index.")
            logging.debug(f"{prefix} Current keys in all_tables_data after adding empty Table {table_index}: {list(all_tables_data.keys())}")

        logging.debug(f"{prefix} <<< Finished processing Header Row {header_row} (Table Index {table_index}).")


    logging.info(f"{prefix} Completed extraction process. Final dictionary contains data for {len(all_tables_data)} table index(es): {list(all_tables_data.keys())}")
    return all_tables_data

# --- END OF FULL FILE: sheet_parser.py ---