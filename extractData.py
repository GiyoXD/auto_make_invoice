# --- START OF FULL FILE: main.py ---
# --- Fixed datetime JSON serialization ---

import logging
import pprint
import re
import decimal
import os
import json # Added for JSON output
import datetime # <<< ADDED IMPORT for datetime handling
from typing import Dict, List, Any, Optional, Tuple, Union

# Import from our refactored modules
import config as cfg
from excel_handler import ExcelHandler
import sheet_parser
import data_processor # Includes all processing functions

# Configure logging (Set level as needed, DEBUG is useful)
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs (for DEBUG)

# --- Constants for FOB Compounding Formatting ---
FOB_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\PO2) [cite: 1]
FOB_INTRA_CHUNK_SEPARATOR = "\\"  # Separator within a group (e.g., backslash) [cite: 2]
FOB_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline) [cite: 2]


# Type alias for the two possible initial aggregation structures
InitialAggregationResults = Union[
    Dict[Tuple[Any, Any, Optional[decimal.Decimal]], Dict[str, decimal.Decimal]], # Standard Result [cite: 2]
    Dict[Tuple[Any, Any], Dict[str, decimal.Decimal]]                             # Custom Result [cite: 2]
]
# Type alias for the FOB compounding result structure
FobCompoundingResult = Dict[str, Union[str, decimal.Decimal]]


# *** FOB Compounding Function with Chunking ***
def perform_fob_compounding(
    initial_results: InitialAggregationResults, # [cite: 3]
    aggregation_mode: str # 'standard' or 'custom' -> Needed to parse input keys correctly [cite: 3]
) -> Optional[FobCompoundingResult]:
    """
    Performs FOB Compounding from standard or custom aggregation results. [cite: 4]
    This function *always* runs after the initial aggregation step. [cite: 4]
    It combines all unique POs and Items into formatted SINGLE strings
    (chunked based on constants). [cite: 4]
    It sums all SQFT and Amount values. Returns a single dictionary record. [cite: 5]
    Args:
        initial_results: The dictionary from either standard or custom aggregation. [cite: 6]
        aggregation_mode: String ('standard' or 'custom') indicating how to parse
                          the keys of initial_results. [cite: 7]
    Returns:
        A single dictionary with 'combined_po', 'combined_item', 'total_sqft',
        'total_amount', or a default structure if input is empty/invalid. [cite: 8]
        Returns None only on critical internal errors. [cite: 9]
    """
    prefix = "[perform_fob_compounding]" # [cite: 10]
    logging.info(f"{prefix} Starting FOB Compounding (runs always). Using input from '{aggregation_mode}' aggregation.") # [cite: 10]

    # Handle empty input consistently
    if not initial_results: # [cite: 10]
        logging.warning(f"{prefix} Input aggregation results map is empty. Returning default zero/empty FOB record.") # [cite: 10]
        return {
            'combined_po': '', # [cite: 10]
            'combined_item': '', # [cite: 10]
            'total_sqft': decimal.Decimal(0), # [cite: 11]
            'total_amount': decimal.Decimal(0) # [cite: 11]
        }

    unique_pos = set() # [cite: 11]
    unique_items = set() # [cite: 11]
    total_sqft = decimal.Decimal(0) # [cite: 11]
    total_amount = decimal.Decimal(0) # [cite: 11]

    logging.debug(f"{prefix} Processing {len(initial_results)} entries from initial aggregation.") # [cite: 11]

    # Iterate through the initial results
    for key, sums_dict in initial_results.items(): # [cite: 11]
        po_key_val = None # [cite: 11]
        item_key_val = None # [cite: 11]

        # Extract PO and Item from the key based on the initial aggregation mode [cite: 12]
        try:
            if aggregation_mode == 'standard': # [cite: 12]
                if len(key) == 3: po_key_val, item_key_val, _ = key # [cite: 12]
                else: raise ValueError("Invalid standard key length") # [cite: 12]
            elif aggregation_mode == 'custom': # [cite: 12]
                 if len(key) == 2: po_key_val, item_key_val = key # [cite: 13]
                 else: raise ValueError("Invalid custom key length") # [cite: 13]
            else:
                logging.error(f"{prefix} Unknown aggregation mode '{aggregation_mode}' passed. Halting compounding.") # [cite: 13]
                return None # Indicate critical error [cite: 14]
        except Exception as e:
             logging.warning(f"{prefix} Error unpacking key {key} in '{aggregation_mode}' mode: {e}. Skipping entry.") # [cite: 14]
             continue # [cite: 14]

        # --- Collect unique POs/Items ---
        # Ensure conversion to string, handle None gracefully
        po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>" # [cite: 15]
        item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>" # [cite: 15]
        unique_pos.add(po_str) # [cite: 15]
        unique_items.add(item_str) # [cite: 15]

        # --- Sum numeric values ---
        sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0)) # [cite: 15]
        amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0)) # [cite: 15]
        # Validate types before summing
        if not isinstance(sqft_sum, decimal.Decimal): # [cite: 16]
             logging.warning(f"{prefix} Invalid SQFT sum type for key {key}: {type(sqft_sum)}. Using 0.") # [cite: 16]
             sqft_sum = decimal.Decimal(0) # [cite: 16]
        if not isinstance(amount_sum, decimal.Decimal): # [cite: 16]
             logging.warning(f"{prefix} Invalid Amount sum type for key {key}: {type(amount_sum)}. Using 0.") # [cite: 16]
             amount_sum = decimal.Decimal(0) # [cite: 17]
        total_sqft += sqft_sum # [cite: 17]
        total_amount += amount_sum # [cite: 17]

    logging.debug(f"{prefix} Finished processing entries.") # [cite: 17]
    logging.debug(f"{prefix} Unique POs collected: {unique_pos}") # [cite: 17]
    logging.debug(f"{prefix} Unique Items collected: {unique_items}") # [cite: 17]

    # --- Final Combination with Chunking ---
    # Convert sets to lists and sort alphabetically/numerically
    sorted_pos = sorted(list(unique_pos)) # [cite: 17]
    sorted_items = sorted(list(unique_items)) # [cite: 17]
    logging.debug(f"{prefix} Sorted PO list before chunking: {sorted_pos}") # [cite: 17]
    logging.debug(f"{prefix} Sorted Item list before chunking: {sorted_items}") # [cite: 18]

    # Helper function for chunking and joining
    def format_chunks(items: List[str], chunk_size: int, intra_sep: str, inter_sep: str) -> str: # [cite: 18]
        if not items: # [cite: 18]
            return "" # [cite: 18]
        processed_chunks = [] # [cite: 18]
        for i in range(0, len(items), chunk_size): # [cite: 18]
            chunk = items[i:i + chunk_size] # [cite: 18]
            joined_chunk = intra_sep.join(chunk) # Join items within the chunk [cite: 19]
            processed_chunks.append(joined_chunk) # [cite: 19]
        return inter_sep.join(processed_chunks) # Join the chunks together [cite: 19]

    # Apply the chunking format using configured constants
    combined_po_string = format_chunks( # [cite: 19]
        sorted_pos, # [cite: 19]
        FOB_CHUNK_SIZE, # [cite: 19]
        FOB_INTRA_CHUNK_SEPARATOR, # [cite: 19]
        FOB_INTER_CHUNK_SEPARATOR # [cite: 19]
    )
    combined_item_string = format_chunks( # [cite: 20]
        sorted_items, # [cite: 20]
        FOB_CHUNK_SIZE, # [cite: 20]
        FOB_INTRA_CHUNK_SEPARATOR, # [cite: 20]
        FOB_INTER_CHUNK_SEPARATOR # [cite: 20]
    )

    # DEBUG LOGGING remains useful
    logging.debug(f"{prefix} Final combined_po_string (Type: {type(combined_po_string)}): '{combined_po_string}'") # [cite: 20]
    logging.debug(f"{prefix} Final combined_item_string (Type: {type(combined_item_string)}): '{combined_item_string}'") # [cite: 20]

    # Construct Result Dictionary
    fob_compounded_result: FobCompoundingResult = { # [cite: 20]
        'combined_po': combined_po_string,    # Now formatted with chunks [cite: 20]
        'combined_item': combined_item_string, # Now formatted with chunks [cite: 21]
        'total_sqft': total_sqft, # [cite: 21]
        'total_amount': total_amount # [cite: 21]
    }

    logging.info(f"{prefix} FOB Compounding complete.") # [cite: 21]
    return fob_compounded_result # [cite: 21]


# --- >>> ADDED: Default JSON Serializer Function <<< ---
def json_serializer_default(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat() # Convert date/datetime to ISO string format
    elif isinstance(obj, decimal.Decimal): # Keep Decimal handling here too
        return str(obj)
    elif isinstance(obj, set): # Optional: Handle sets if needed
        return list(obj)
    # Add other custom types if needed
    # elif isinstance(obj, YourCustomClass):
    #     return obj.__dict__
    raise TypeError (f"Object of type {obj.__class__.__name__} is not JSON serializable")
# --- >>> END OF ADDED FUNCTION <<< ---


# Helper function to make data JSON serializable
# Handles tuple keys in aggregation results (Decimal handled by default now)
def make_json_serializable(data):
    """Recursively converts tuple keys in dicts to strings."""
    # NOTE: Decimal and datetime conversion removed as it's now handled by the 'default' serializer
    if isinstance(data, dict): # [cite: 21]
        # Special handling for aggregation results with tuple keys
        if data and all(isinstance(k, tuple) for k in data.keys()): # [cite: 22]
             # Convert tuple keys to strings
             return {str(k): make_json_serializable(v) for k, v in data.items()} # [cite: 22]
        else:
             # Regular dictionary processing
             return {str(k): make_json_serializable(v) for k, v in data.items()} # [cite: 23]
    elif isinstance(data, list): # [cite: 23]
        return [make_json_serializable(item) for item in data] # [cite: 23]
    # elif isinstance(data, decimal.Decimal): # Removed - handled by default
    #     return str(data)
    elif data is None: # [cite: 23]
        return None # JSON null [cite: 23]
    # Datetime also removed - handled by default
    return data # [cite: 24]

def run_invoice_automation():
    """Main function to find tables, extract, and process data for each.""" # [cite: 24]
    logging.info("--- Starting Invoice Automation ---") # [cite: 24]
    handler = None # [cite: 24]
    actual_sheet_name = None # [cite: 24]
    input_filename = "Unknown" # Initialize [cite: 24]

    processed_tables: Dict[int, Dict[str, Any]] = {} # [cite: 24]
    all_tables_data: Dict[int, Dict[str, List[Any]]] = {} # [cite: 24]

    # Global dictionaries for initial aggregation results
    global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal]], Dict[str, decimal.Decimal]] = {} # [cite: 24]
    global_custom_aggregation_results: Dict[Tuple[Any, Any], Dict[str, decimal.Decimal]] = {} # [cite: 24]
    # Global variable for the final single FOB compounded result [cite: 25]
    global_fob_compounded_result: Optional[FobCompoundingResult] = None # [cite: 25]

    aggregation_mode_used = "standard" # Default, determines initial aggregation type [cite: 25]

    # --- Determine Initial Aggregation Strategy ---
    use_custom_aggregation = False # [cite: 25]
    # Ensure cfg.INPUT_EXCEL_FILE is accessible for filename check
    try:
        input_filename = os.path.basename(cfg.INPUT_EXCEL_FILE) # [cite: 25]
        logging.info(f"Checking workbook filename '{input_filename}' for custom aggregation trigger.") # [cite: 25]
        for prefix in cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES: # [cite: 25]
             if input_filename.startswith(prefix): # [cite: 26]
                use_custom_aggregation = True # [cite: 26]
                aggregation_mode_used = "custom" # [cite: 26]
                logging.info(f"Workbook filename matches prefix '{prefix}'. Using CUSTOM initial aggregation.") # [cite: 26]
                break # [cite: 27]
        if not use_custom_aggregation: # [cite: 27]
             logging.info(f"Workbook filename does not match custom prefixes. Using STANDARD initial aggregation.") # [cite: 27]
             aggregation_mode_used = "standard" # [cite: 27]
    except Exception as e:
        logging.error(f"Error accessing config or input filename for aggregation strategy: {e}") # [cite: 27]
        # Decide on a fallback or re-raise, here we default to standard [cite: 28]
        logging.warning("Defaulting to STANDARD aggregation due to error.") # [cite: 28]
        aggregation_mode_used = "standard" # [cite: 28]
        use_custom_aggregation = False # [cite: 28]
        input_filename = "ErrorDeterminingFilename" # [cite: 28]
    # ---------------------------------------------

    try:
        # --- Steps 1-4: Load, Find Headers, Map Columns, Extract Data ---
        logging.info(f"Loading workbook: {cfg.INPUT_EXCEL_FILE}") # [cite: 28]
        handler = ExcelHandler(cfg.INPUT_EXCEL_FILE) # [cite: 29]
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True) # [cite: 29]
        if sheet is None: raise RuntimeError(f"Failed to load sheet from '{cfg.INPUT_EXCEL_FILE}'.") # [cite: 29]
        actual_sheet_name = sheet.title # [cite: 29]
        logging.info(f"Successfully loaded worksheet: '{actual_sheet_name}' from '{input_filename}'") # [cite: 29]

        logging.info("Searching for all header rows...") # [cite: 29]
        header_rows = sheet_parser.find_all_header_rows(sheet, cfg.HEADER_IDENTIFICATION_PATTERN, cfg.HEADER_SEARCH_ROW_RANGE, cfg.HEADER_SEARCH_COL_RANGE) # [cite: 29]
        if not header_rows: raise RuntimeError("Could not find any header rows.") # [cite: 29]
        logging.info(f"Found {len(header_rows)} potential header row(s) at: {header_rows}") # [cite: 30]

        first_header_row = header_rows[0] # [cite: 30]
        logging.info(f"Mapping columns based on first header row ({first_header_row})...") # [cite: 30]
        column_mapping = sheet_parser.map_columns_to_headers(sheet, first_header_row, cfg.HEADER_SEARCH_COL_RANGE) # [cite: 30]
        if not column_mapping: raise RuntimeError("Failed to map columns.") # [cite: 30]
        logging.debug(f"Mapped columns:\n{pprint.pformat(column_mapping)}") # [cite: 30]
        if 'amount' not in column_mapping: raise RuntimeError("Essential 'amount' column mapping failed.") # [cite: 30]

        logging.info("Extracting data for all tables...") # [cite: 30]
        all_tables_data = sheet_parser.extract_multiple_tables(sheet, header_rows, column_mapping) # [cite: 31]
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG: # [cite: 31]
            log_str = pprint.pformat(all_tables_data) # [cite: 31]
            if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)" # [cite: 31]
            logging.debug(f"--- Raw Extracted Data ({len(all_tables_data)} Table(s)) ---\n{log_str}") # [cite: 31]
        if not all_tables_data: logging.warning("Extraction resulted in empty data structure.") # [cite: 31]
        # --- End Steps 1-4 --- [cite: 32]


        # --- 5. Process Each Table (CBM, Distribute, Initial Aggregate) ---
        logging.info(f"--- Starting Data Processing Loop for {len(all_tables_data)} Extracted Table(s) ---") # [cite: 32]
        for table_index, raw_data_dict in all_tables_data.items(): # [cite: 32]
            current_table_data = all_tables_data.get(table_index) # [cite: 32]
            if current_table_data is None: # [cite: 32]
                logging.error(f"Skipping processing for missing table_index {table_index}.") # [cite: 32]
                continue # [cite: 33]

            logging.info(f"--- Processing Table Index {table_index} ---") # [cite: 33]
            if not isinstance(current_table_data, dict) or not current_table_data or not any(isinstance(v, list) and v for v in current_table_data.values()): # [cite: 33]
                logging.warning(f"Table {table_index} empty or invalid. Skipping steps.") # [cite: 33]
                processed_tables[table_index] = current_table_data # Still store the (empty/invalid) raw data [cite: 34]
                continue # [cite: 34]

            # 5a. CBM Calculation
            logging.info(f"Table {table_index}: Calculating CBM values...") # [cite: 34]
            try:
                 data_after_cbm = data_processor.process_cbm_column(current_table_data) # [cite: 35]
            except Exception as e:
                logging.error(f"CBM calc error Table {table_index}: {e}", exc_info=True) # [cite: 35]
                data_after_cbm = current_table_data # Use original data if CBM fails [cite: 35]

            # 5b. Distribution
            logging.info(f"Table {table_index}: Distributing values...") # [cite: 35]
            try: # [cite: 36]
                data_after_distribution = data_processor.distribute_values(data_after_cbm, cfg.COLUMNS_TO_DISTRIBUTE, cfg.DISTRIBUTION_BASIS_COLUMN) # [cite: 36]
                processed_tables[table_index] = data_after_distribution # Store successfully processed data [cite: 36]
            except data_processor.ProcessingError as pe: # type: ignore # Assuming ProcessingError is defined in data_processor
                logging.error(f"Distribution failed Table {table_index}: {pe}. Storing pre-distribution data.") # [cite: 36]
                processed_tables[table_index] = data_after_cbm # Store data after CBM but before failed distribution [cite: 37]
                continue # Skip aggregation for this table if distribution failed [cite: 37]
            except Exception as e:
                logging.error(f"Unexpected distribution error Table {table_index}: {e}", exc_info=True) # [cite: 37]
                processed_tables[table_index] = data_after_cbm # Store data after CBM [cite: 38]
                continue # Skip aggregation [cite: 38]

            # 5c. Initial Aggregation (Standard or Custom) - Only run if distribution succeeded
            data_for_aggregation = processed_tables.get(table_index) # Get the successfully distributed data [cite: 38]
            if isinstance(data_for_aggregation, dict) and data_for_aggregation: # [cite: 38]
                 try: # [cite: 39]
                    if use_custom_aggregation: # [cite: 39]
                        logging.info(f"Table {table_index}: Updating global CUSTOM aggregation...") # [cite: 39]
                        data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results) # [cite: 39]
                        logging.info(f"Table {table_index}: CUSTOM aggregation map updated.") # [cite: 40]
                    else:
                        logging.info(f"Table {table_index}: Updating global STANDARD aggregation...") # [cite: 40]
                        data_processor.aggregate_standard_by_po_item_price(data_for_aggregation, global_standard_aggregation_results) # [cite: 40]
                        logging.info(f"Table {table_index}: STANDARD aggregation map updated.") # [cite: 41]
                 except Exception as agg_e:
                    logging.error(f"Global {aggregation_mode_used.upper()} aggregation update failed for Table {table_index}: {agg_e}", exc_info=True) # [cite: 41]
                    # Decide if you want to proceed or halt. Here we log and continue. [cite: 42]
            else:
                 # This case should ideally not happen if distribution succeeded and returned data
                 logging.warning(f"Table {table_index}: Skipping initial aggregation update (processed data invalid/empty after distribution step).") # [cite: 42]

            logging.info(f"--- Finished Processing All Steps for Table Index {table_index} ---") # [cite: 42]
        # --- End Processing Loop ---


        # --- 6. Post-Loop: Perform FOB Compounding (ALWAYS RUNS) --- [cite: 43]
        logging.info("--- All Table Processing Loops Completed ---") # [cite: 43]
        logging.info("--- Performing Final FOB Compounding (Always Runs) ---") # [cite: 43]
        try:
            # Determine the source data based on the mode used during the loop
            initial_agg_data_source = global_custom_aggregation_results if use_custom_aggregation else global_standard_aggregation_results # [cite: 43]
            global_fob_compounded_result = perform_fob_compounding( # [cite: 44]
                initial_agg_data_source, # [cite: 44]
                aggregation_mode_used # Pass mode to help parse input keys [cite: 44]
            )
            logging.info("--- FOB Compounding Finished ---") # [cite: 44]
        except Exception as fob_e:
             logging.error(f"An error occurred during the final FOB Compounding step: {fob_e}", exc_info=True) # [cite: 44]
             logging.error("FOB Compounding results may be incomplete or missing.") # [cite: 45]


        # --- 7. Output / Further Steps ---
        logging.info(f"Final processed data structure contains {len(processed_tables)} table(s).") # [cite: 45]
        logging.info(f"Initial aggregation mode used: {aggregation_mode_used.upper()}") # [cite: 45]

        # --- Log Initial Aggregation Results (DEBUG Level) ---
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG: # [cite: 45]
            if aggregation_mode_used == "custom": # [cite: 46]
                 log_str = pprint.pformat(global_custom_aggregation_results) # [cite: 46]
                 if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)" # [cite: 46]
                 logging.debug(f"--- Full Global CUSTOM Aggregation Results (Input to FOB) ---\n{log_str}") # [cite: 46]
            else:
                 log_str = pprint.pformat(global_standard_aggregation_results) # [cite: 47]
                 if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)" # [cite: 47]
                 logging.debug(f"--- Full Global STANDARD Aggregation Results (Input to FOB) ---\n{log_str}") # [cite: 47]

        # --- Log Final FOB Compounded Result (INFO Level) - REFINED --- #
        logging.info(f"--- Final FOB Compounded Result (Workbook: '{input_filename}') ---") # [cite: 47]
        if global_fob_compounded_result is not None: # [cite: 48]
            # Get the final values
            po_string_value = global_fob_compounded_result.get('combined_po', '<Not Found>') # [cite: 48]
            item_string_value = global_fob_compounded_result.get('combined_item', '<Not Found>') # [cite: 48]
            total_sqft_value = global_fob_compounded_result.get('total_sqft', 'N/A') # [cite: 48]
            total_amount_value = global_fob_compounded_result.get('total_amount', 'N/A') # [cite: 48]

            # Log each component explicitly for clarity [cite: 49]
            logging.info(f"Combined POs (Type: {type(po_string_value)}):") # [cite: 49]
            logging.info(f"  repr(): {repr(po_string_value)}") # Shows literal \n and \\ [cite: 49]
            logging.info(f"  Raw Value:\n{po_string_value}")   # Renders multi-line if \n present [cite: 49]
            logging.info("-" * 30) # [cite: 49]

            logging.info(f"Combined Items (Type: {type(item_string_value)}):") # [cite: 49]
            logging.info(f"  repr(): {repr(item_string_value)}") # [cite: 50]
            logging.info(f"  Raw Value:\n{item_string_value}") # [cite: 50]
            logging.info("-" * 30) # [cite: 50]

            logging.info(f"Total SQFT: {total_sqft_value} (Type: {type(total_sqft_value)})") # [cite: 50]
            logging.info(f"Total Amount: {total_amount_value} (Type: {type(total_amount_value)})") # [cite: 50]
            logging.info("-" * 30) # [cite: 50]
        else:
            logging.error("FOB Compounding result is None or was not set, potentially due to an error during compounding.") # [cite: 51]
        # --- End Final Logging ---


        # --- 8. Generate JSON Output ---
        logging.info("--- Preparing Data for JSON Output ---") # [cite: 51]
        try:
            # Determine which initial aggregation results were used as input for FOB
             # This dictionary holds the sums *before* they were compounded into the final string/total [cite: 52]
            initial_aggregation_data_input_to_fob = global_custom_aggregation_results if use_custom_aggregation else global_standard_aggregation_results # [cite: 52]

            # Create the structure to be converted to JSON
            # Use the helper function to ensure serializability
            final_json_structure = { # [cite: 52]
                 "metadata": { # [cite: 53]
                    "workbook_filename": input_filename, # [cite: 53]
                    "worksheet_name": actual_sheet_name, # [cite: 53]
                    "aggregation_mode_used": aggregation_mode_used, # [cite: 53]
                    "fob_chunk_size": FOB_CHUNK_SIZE, # [cite: 53]
                     "fob_intra_separator": FOB_INTRA_CHUNK_SEPARATOR, # [cite: 54]
                    "fob_inter_separator": FOB_INTER_CHUNK_SEPARATOR, # [cite: 54]
                },
                # IMPORTANT: processed_tables can be very large. [cite: 54]
                 # Including it fully might create huge JSON files. [cite: 55]
                 # Consider summarizing or excluding if size is an issue. [cite: 55]
                 "processed_tables_data": make_json_serializable(processed_tables), # [cite: 56]

                # This reflects the aggregated sums *before* the final string/total compounding
                "initial_aggregation_input_to_fob": make_json_serializable(initial_aggregation_data_input_to_fob), # [cite: 56]

                # Include the final compounded result (strings and totals)
                "final_fob_compounded_result": make_json_serializable(global_fob_compounded_result) # [cite: 56]
            }

             # Convert the structure to a JSON string (pretty-printed)
             # --- >>> MODIFIED LINE: Added default argument <<< ---
            json_output_string = json.dumps(final_json_structure,
                                            indent=4,
                                            default=json_serializer_default) # <<< FIX IMPLEMENTED [cite: 57]

            # Log the JSON output (or a preview if too large)
            logging.info("--- Generated JSON Output ---") # [cite: 57]
            max_log_json_len = 5000 # Limit length for console logging [cite: 57]
            if len(json_output_string) <= max_log_json_len: # [cite: 58]
                logging.info(json_output_string) # [cite: 58]
            else:
                logging.info(f"JSON output is large ({len(json_output_string)} chars). Logging preview:") # [cite: 58]
                logging.info(json_output_string[:max_log_json_len] + "\n... (JSON output truncated in log)") # [cite: 59]

            # Save the JSON to a file
            json_output_filename = f"output_{os.path.splitext(input_filename)[0]}.json" # [cite: 59]
            try:
                with open(json_output_filename, 'w', encoding='utf-8') as f_json: # [cite: 59]
                     f_json.write(json_output_string) # [cite: 60]
                logging.info(f"Successfully saved JSON output to '{json_output_filename}'") # [cite: 60]
            except IOError as io_err:
                logging.error(f"Failed to write JSON output to file '{json_output_filename}': {io_err}") # [cite: 60]
            except Exception as write_err:
                 logging.error(f"An unexpected error occurred while writing JSON file: {write_err}", exc_info=True) # [cite: 61]

        except TypeError as json_err: # Catches errors from json.dumps if default handler fails
            logging.error(f"Failed to serialize data to JSON: {json_err}. Check data types and default handler.", exc_info=True) # [cite: 61]
        except Exception as e:
            logging.error(f"An unexpected error occurred during JSON generation: {e}", exc_info=True) # [cite: 61]
        # --- End JSON Generation ---


        logging.info("--- Invoice Automation Finished Successfully ---") # [cite: 61]

    except FileNotFoundError as e: logging.error(f"Input file error: {e}") # [cite: 62]
    except RuntimeError as e: logging.error(f"Processing halted due to critical error: {e}") # [cite: 62]
    except Exception as e: logging.error(f"An unexpected error occurred in the main script execution: {e}", exc_info=True) # [cite: 62]
    finally:
        if handler: # [cite: 62]
            handler.close() # Ensure workbook is closed even on error [cite: 62]
        logging.info("--- Automation Run Complete ---") # [cite: 62]


if __name__ == "__main__": # [cite: 62]
    # Make sure config.py exists and defines necessary variables like: [cite: 62]
    # INPUT_EXCEL_FILE, SHEET_NAME, HEADER_IDENTIFICATION_PATTERN, [cite: 63]
    # HEADER_SEARCH_ROW_RANGE, HEADER_SEARCH_COL_RANGE, [cite: 63]
    # COLUMNS_TO_DISTRIBUTE, DISTRIBUTION_BASIS_COLUMN, [cite: 63]
    # CUSTOM_AGGREGATION_WORKBOOK_PREFIXES [cite: 63]
    run_invoice_automation() # [cite: 63]

# --- END OF FULL FILE: main.py --- [cite: 63]