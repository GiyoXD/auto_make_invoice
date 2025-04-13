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
# UPDATED Type Alias to reflect new key structures
InitialAggregationResults = Union[
    Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]], # Standard Result (PO, Item, Price, Desc)
    Dict[Tuple[Any, Any, Optional[str]], Dict[str, decimal.Decimal]]                             # Custom Result (PO, Item, Desc)
]
# Type alias for the FOB compounding result structure
FobCompoundingResult = Dict[str, Union[str, decimal.Decimal]]


# *** FOB Compounding Function with Chunking ***
def perform_fob_compounding(
    initial_results: InitialAggregationResults, # Type hint updated
    aggregation_mode: str # 'standard' or 'custom' -> Needed to parse input keys correctly
) -> Optional[FobCompoundingResult]:
    """
    Performs FOB Compounding from standard or custom aggregation results.
    This function *always* runs after the initial aggregation step.
    It combines all unique POs and Items into formatted SINGLE strings
    (chunked based on constants).
    It sums all SQFT and Amount values. Returns a single dictionary record.
    Args:
        initial_results: The dictionary from EITHER standard OR custom aggregation (determined by aggregation_mode).
                         The key structure depends on the mode.
        aggregation_mode: String ('standard' or 'custom') indicating how to parse
                          the keys of initial_results.
    Returns:
        A single dictionary with 'combined_po', 'combined_item', 'total_sqft',
        'total_amount', or a default structure if input is empty/invalid.
        Returns None only on critical internal errors.
    """
    prefix = "[perform_fob_compounding]"
    logging.info(f"{prefix} Starting FOB Compounding. Using input from '{aggregation_mode}' aggregation map.")

    # Handle empty input consistently
    if not initial_results:
        logging.warning(f"{prefix} Input aggregation results map is empty. Returning default zero/empty FOB record.")
        return {
            'combined_po': '',
            'combined_item': '',
            'total_sqft': decimal.Decimal(0),
            'total_amount': decimal.Decimal(0)
        }

    unique_pos = set()
    unique_items = set()
    total_sqft = decimal.Decimal(0)
    total_amount = decimal.Decimal(0)

    logging.debug(f"{prefix} Processing {len(initial_results)} entries from initial aggregation.")

    # Iterate through the initial results
    for key, sums_dict in initial_results.items():
        po_key_val = None
        item_key_val = None
        # Note: Description from key is ignored for compounding, but needs unpacking

        # Extract PO and Item from the key based on the initial aggregation mode
        try:
            if aggregation_mode == 'standard':
                 # UPDATED: Expecting 4 elements now (PO, Item, Price, Desc)
                if len(key) == 4: po_key_val, item_key_val, _, _ = key # Unpack first two
                else: raise ValueError(f"Invalid standard key length ({len(key)}), expected 4")
            elif aggregation_mode == 'custom':
                 # UPDATED: Expecting 3 elements now (PO, Item, Desc)
                 if len(key) == 3: po_key_val, item_key_val, _ = key # Unpack first two
                 else: raise ValueError(f"Invalid custom key length ({len(key)}), expected 3")
            else:
                logging.error(f"{prefix} Unknown aggregation mode '{aggregation_mode}' passed. Halting compounding.")
                return None # Indicate critical error
        except (ValueError, TypeError, IndexError) as e: # Catch more specific errors
             logging.warning(f"{prefix} Error unpacking key {key} (type: {type(key)}) in '{aggregation_mode}' mode: {e}. Skipping entry.")
             continue

        # --- Collect unique POs/Items ---
        # Ensure conversion to string, handle None gracefully
        po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
        item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
        unique_pos.add(po_str)
        unique_items.add(item_str)

        # --- Sum numeric values ---
        sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
        amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))
        # Validate types before summing
        if not isinstance(sqft_sum, decimal.Decimal):
             logging.warning(f"{prefix} Invalid SQFT sum type for key {key}: {type(sqft_sum)}. Using 0.")
             sqft_sum = decimal.Decimal(0)
        if not isinstance(amount_sum, decimal.Decimal):
             logging.warning(f"{prefix} Invalid Amount sum type for key {key}: {type(amount_sum)}. Using 0.")
             amount_sum = decimal.Decimal(0)
        total_sqft += sqft_sum
        total_amount += amount_sum

    logging.debug(f"{prefix} Finished processing entries.")
    logging.debug(f"{prefix} Unique POs collected ({len(unique_pos)}): {unique_pos if len(unique_pos) < 20 else str(list(unique_pos)[:20]) + '...'}") # Truncate long logs
    logging.debug(f"{prefix} Unique Items collected ({len(unique_items)}): {unique_items if len(unique_items) < 20 else str(list(unique_items)[:20]) + '...'}") # Truncate long logs

    # --- Final Combination with Chunking ---
    # Convert sets to lists and sort alphabetically/numerically
    sorted_pos = sorted(list(unique_pos))
    sorted_items = sorted(list(unique_items))
    # logging.debug(f"{prefix} Sorted PO list before chunking: {sorted_pos}") # Reduced verbosity
    # logging.debug(f"{prefix} Sorted Item list before chunking: {sorted_items}") # Reduced verbosity

    # Helper function for chunking and joining
    def format_chunks(items: List[str], chunk_size: int, intra_sep: str, inter_sep: str) -> str:
        if not items:
            return ""
        processed_chunks = []
        for i in range(0, len(items), chunk_size):
            chunk = items[i:i + chunk_size]
            joined_chunk = intra_sep.join(chunk) # Join items within the chunk
            processed_chunks.append(joined_chunk)
        return inter_sep.join(processed_chunks) # Join the chunks together

    # Apply the chunking format using configured constants
    combined_po_string = format_chunks(
        sorted_pos,
        FOB_CHUNK_SIZE,
        FOB_INTRA_CHUNK_SEPARATOR,
        FOB_INTER_CHUNK_SEPARATOR
    )
    combined_item_string = format_chunks(
        sorted_items,
        FOB_CHUNK_SIZE,
        FOB_INTRA_CHUNK_SEPARATOR,
        FOB_INTER_CHUNK_SEPARATOR
    )

    # DEBUG LOGGING remains useful
    # logging.debug(f"{prefix} Final combined_po_string (Type: {type(combined_po_string)}): '{combined_po_string}'") # Reduced verbosity
    # logging.debug(f"{prefix} Final combined_item_string (Type: {type(combined_item_string)}): '{combined_item_string}'") # Reduced verbosity

    # Construct Result Dictionary
    fob_compounded_result: FobCompoundingResult = {
        'combined_po': combined_po_string,    # Now formatted with chunks
        'combined_item': combined_item_string, # Now formatted with chunks
        'total_sqft': total_sqft,
        'total_amount': total_amount
    }

    logging.info(f"{prefix} FOB Compounding complete.")
    return fob_compounded_result


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
# Handles tuple keys in aggregation results
def make_json_serializable(data):
    """Recursively converts tuple keys in dicts to strings and handles non-serializable types."""
    # NOTE: Using the default serializer for json.dumps handles Decimal and datetime now.
    # This function primarily focuses on converting tuple keys.
    if isinstance(data, dict):
        # Convert all keys to string, including tuple keys
        return {str(k): make_json_serializable(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [make_json_serializable(item) for item in data]
    elif data is None:
        return None # JSON null
    # Let the default handler in json.dumps deal with Decimal, datetime, etc.
    return data

def run_invoice_automation():
    """Main function to find tables, extract, and process data for each."""
    logging.info("--- Starting Invoice Automation ---")
    handler = None
    actual_sheet_name = None
    input_filename = "Unknown"

    processed_tables: Dict[int, Dict[str, Any]] = {}
    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}

    # Global dictionaries for initial aggregation results
    # UPDATED Type Hints for new key structures
    global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]] = {}
    global_custom_aggregation_results: Dict[Tuple[Any, Any, Optional[str]], Dict[str, decimal.Decimal]] = {}
    # Global variable for the final single FOB compounded result
    global_fob_compounded_result: Optional[FobCompoundingResult] = None

    aggregation_mode_used = "standard" # Default, determines WHICH aggregation feeds FOB

    # --- Determine Initial Aggregation Strategy ---
    use_custom_aggregation_for_fob = False # Determines which map feeds FOB
    # Ensure cfg.INPUT_EXCEL_FILE is accessible for filename check
    try:
        input_filename = os.path.basename(cfg.INPUT_EXCEL_FILE)
        logging.info(f"Checking workbook filename '{input_filename}' to determine PRIMARY aggregation mode for FOB compounding.")
        for prefix in cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES:
             if input_filename.startswith(prefix):
                use_custom_aggregation_for_fob = True # This workbook primarily uses custom for FOB
                aggregation_mode_used = "custom"
                logging.info(f"Workbook filename matches prefix '{prefix}'. Will use CUSTOM aggregation results for FOB compounding.")
                break
        if not use_custom_aggregation_for_fob:
             logging.info(f"Workbook filename does not match custom prefixes. Will use STANDARD aggregation results for FOB compounding.")
             aggregation_mode_used = "standard"
    except Exception as e:
        logging.error(f"Error accessing config or input filename for aggregation strategy: {e}")
        logging.warning("Defaulting to STANDARD aggregation for FOB compounding due to error.")
        aggregation_mode_used = "standard"
        use_custom_aggregation_for_fob = False
        input_filename = "ErrorDeterminingFilename"
    # ---------------------------------------------

    try:
        # --- Steps 1-4: Load, Find Headers, Map Columns, Extract Data ---
        logging.info(f"Loading workbook: {cfg.INPUT_EXCEL_FILE}")
        handler = ExcelHandler(cfg.INPUT_EXCEL_FILE)
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True)
        if sheet is None: raise RuntimeError(f"Failed to load sheet from '{cfg.INPUT_EXCEL_FILE}'.")
        actual_sheet_name = sheet.title
        logging.info(f"Successfully loaded worksheet: '{actual_sheet_name}' from '{input_filename}'")

        logging.info("Searching for all header rows...")
        header_rows = sheet_parser.find_all_header_rows(sheet, cfg.HEADER_IDENTIFICATION_PATTERN, cfg.HEADER_SEARCH_ROW_RANGE, cfg.HEADER_SEARCH_COL_RANGE)
        if not header_rows: raise RuntimeError("Could not find any header rows.")
        logging.info(f"Found {len(header_rows)} potential header row(s) at: {header_rows}")

        first_header_row = header_rows[0]
        logging.info(f"Mapping columns based on first header row ({first_header_row})...")
        column_mapping = sheet_parser.map_columns_to_headers(sheet, first_header_row, cfg.HEADER_SEARCH_COL_RANGE)
        if not column_mapping: raise RuntimeError("Failed to map columns.")
        logging.debug(f"Mapped columns:\n{pprint.pformat(column_mapping)}")
        # Ensure core columns are present, but allow processing even if description isn't mapped initially
        if 'amount' not in column_mapping: raise RuntimeError("Essential 'amount' column mapping failed.")
        if 'description' not in column_mapping:
             logging.warning("Column 'description' not found during initial mapping. Aggregation keys will use None for description.")


        logging.info("Extracting data for all tables...")
        all_tables_data = sheet_parser.extract_multiple_tables(sheet, header_rows, column_mapping)
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
            log_str = pprint.pformat(all_tables_data)
            if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
            logging.debug(f"--- Raw Extracted Data ({len(all_tables_data)} Table(s)) ---\n{log_str}")
        if not all_tables_data: logging.warning("Extraction resulted in empty data structure.")
        # --- End Steps 1-4 ---


        # --- 5. Process Each Table (CBM, Distribute, Initial Aggregate) ---
        logging.info(f"--- Starting Data Processing Loop for {len(all_tables_data)} Extracted Table(s) ---")
        for table_index, raw_data_dict in all_tables_data.items():
            current_table_data = all_tables_data.get(table_index)
            if current_table_data is None:
                logging.error(f"Skipping processing for missing table_index {table_index}.")
                continue

            logging.info(f"--- Processing Table Index {table_index} ---")
            if not isinstance(current_table_data, dict) or not current_table_data or not any(isinstance(v, list) and v for v in current_table_data.values()):
                logging.warning(f"Table {table_index} empty or invalid. Skipping processing steps.")
                processed_tables[table_index] = current_table_data # Store the raw data
                continue

            # 5a. CBM Calculation
            logging.info(f"Table {table_index}: Calculating CBM values...")
            try:
                 data_after_cbm = data_processor.process_cbm_column(current_table_data)
            except Exception as e:
                logging.error(f"CBM calc error Table {table_index}: {e}", exc_info=True)
                data_after_cbm = current_table_data # Use original data if CBM fails

            # 5b. Distribution
            logging.info(f"Table {table_index}: Distributing values...")
            try:
                data_after_distribution = data_processor.distribute_values(data_after_cbm, cfg.COLUMNS_TO_DISTRIBUTE, cfg.DISTRIBUTION_BASIS_COLUMN)
                processed_tables[table_index] = data_after_distribution # Store successfully processed data
            except data_processor.ProcessingError as pe: # type: ignore
                logging.error(f"Distribution failed Table {table_index}: {pe}. Storing pre-distribution data.")
                processed_tables[table_index] = data_after_cbm
                # Continue to aggregation even if distribution failed, using pre-distribution data
                data_for_aggregation = data_after_cbm
                # continue # Original logic skipped aggregation on distribution failure
            except Exception as e:
                logging.error(f"Unexpected distribution error Table {table_index}: {e}", exc_info=True)
                processed_tables[table_index] = data_after_cbm
                # Continue to aggregation even if distribution failed, using pre-distribution data
                data_for_aggregation = data_after_cbm
                # continue # Original logic skipped aggregation on unexpected distribution failure
            else:
                 # If distribution succeeded, use the distributed data for aggregation
                 data_for_aggregation = processed_tables.get(table_index)


            # 5c. Initial Aggregation (ALWAYS RUN BOTH Standard and Custom)
            if isinstance(data_for_aggregation, dict) and data_for_aggregation:
                 # Run Standard Aggregation
                 try:
                    logging.info(f"Table {table_index}: Updating global STANDARD aggregation...")
                    data_processor.aggregate_standard_by_po_item_price(data_for_aggregation, global_standard_aggregation_results)
                    logging.debug(f"Table {table_index}: STANDARD aggregation map updated. Size: {len(global_standard_aggregation_results)}")
                 except Exception as agg_e_std:
                    logging.error(f"Global STANDARD aggregation update failed for Table {table_index}: {agg_e_std}", exc_info=True)

                 # Run Custom Aggregation
                 try:
                    logging.info(f"Table {table_index}: Updating global CUSTOM aggregation...")
                    data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results)
                    logging.debug(f"Table {table_index}: CUSTOM aggregation map updated. Size: {len(global_custom_aggregation_results)}")
                 except Exception as agg_e_cust:
                    logging.error(f"Global CUSTOM aggregation update failed for Table {table_index}: {agg_e_cust}", exc_info=True)
            else:
                 logging.warning(f"Table {table_index}: Skipping initial aggregation update (data for aggregation invalid/empty).")

            logging.info(f"--- Finished Processing All Steps for Table Index {table_index} ---")
        # --- End Processing Loop ---


        # --- 6. Post-Loop: Perform FOB Compounding (ALWAYS RUNS) ---
        logging.info("--- All Table Processing Loops Completed ---")
        logging.info(f"--- Performing Final FOB Compounding (Using '{aggregation_mode_used.upper()}' aggregation results as input) ---")
        try:
            # Determine the source data based on the mode determined earlier by filename
            initial_agg_data_source = global_custom_aggregation_results if use_custom_aggregation_for_fob else global_standard_aggregation_results
            global_fob_compounded_result = perform_fob_compounding(
                initial_agg_data_source, # Pass the selected map
                aggregation_mode_used # Pass mode to help parse input keys correctly
            )
            logging.info("--- FOB Compounding Finished ---")
        except Exception as fob_e:
             logging.error(f"An error occurred during the final FOB Compounding step: {fob_e}", exc_info=True)
             logging.error("FOB Compounding results may be incomplete or missing.")


        # --- 7. Output / Further Steps ---
        logging.info(f"Final processed data structure contains {len(processed_tables)} table(s).")
        logging.info(f"Primary aggregation mode used for FOB Compounding: {aggregation_mode_used.upper()}")

        # --- Log Initial Aggregation Results (DEBUG Level) ---
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
            # Log Standard Results
            log_str_std = pprint.pformat(global_standard_aggregation_results)
            if len(log_str_std) > MAX_LOG_DICT_LEN: log_str_std = log_str_std[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
            logging.debug(f"--- Full Global STANDARD Aggregation Results ---\n{log_str_std}")
            # Log Custom Results
            log_str_cust = pprint.pformat(global_custom_aggregation_results)
            if len(log_str_cust) > MAX_LOG_DICT_LEN: log_str_cust = log_str_cust[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
            logging.debug(f"--- Full Global CUSTOM Aggregation Results ---\n{log_str_cust}")


        # --- Log Final FOB Compounded Result (INFO Level) - REFINED --- #
        logging.info(f"--- Final FOB Compounded Result (Workbook: '{input_filename}', Based on '{aggregation_mode_used.upper()}' Input) ---")
        if global_fob_compounded_result is not None:
            po_string_value = global_fob_compounded_result.get('combined_po', '<Not Found>')
            item_string_value = global_fob_compounded_result.get('combined_item', '<Not Found>')
            total_sqft_value = global_fob_compounded_result.get('total_sqft', 'N/A')
            total_amount_value = global_fob_compounded_result.get('total_amount', 'N/A')

            logging.info(f"Combined POs (Type: {type(po_string_value)}):")
            logging.info(f"  repr(): {repr(po_string_value)}")
            logging.info(f"  Raw Value:\n{po_string_value}")
            logging.info("-" * 30)

            logging.info(f"Combined Items (Type: {type(item_string_value)}):")
            logging.info(f"  repr(): {repr(item_string_value)}")
            logging.info(f"  Raw Value:\n{item_string_value}")
            logging.info("-" * 30)

            logging.info(f"Total SQFT: {total_sqft_value} (Type: {type(total_sqft_value)})")
            logging.info(f"Total Amount: {total_amount_value} (Type: {type(total_amount_value)})")
            logging.info("-" * 30)
        else:
            logging.error("FOB Compounding result is None or was not set, potentially due to an error during compounding.")
        # --- End Final Logging ---


        # --- 8. Generate JSON Output ---
        logging.info("--- Preparing Data for JSON Output ---")
        try:
            # Create the structure to be converted to JSON
            # Use the helper function to ensure serializability
            final_json_structure = {
                 "metadata": {
                    "workbook_filename": input_filename,
                    "worksheet_name": actual_sheet_name,
                    "fob_compounding_input_mode": aggregation_mode_used, # Clarify which mode fed FOB
                    "fob_chunk_size": FOB_CHUNK_SIZE,
                     "fob_intra_separator": FOB_INTRA_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "fob_inter_separator": FOB_INTER_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "timestamp": datetime.datetime.now() # Add generation timestamp
                },
                 # Include processed table data (potentially large)
                 "processed_tables_data": make_json_serializable(processed_tables),

                # Include BOTH aggregation results explicitly
                "standard_aggregation_results": make_json_serializable(global_standard_aggregation_results),
                "custom_aggregation_results": make_json_serializable(global_custom_aggregation_results),

                # Include the final compounded result (derived from one of the above, based on mode)
                "final_fob_compounded_result": make_json_serializable(global_fob_compounded_result)
            }

             # Convert the structure to a JSON string (pretty-printed)
            json_output_string = json.dumps(final_json_structure,
                                            indent=4,
                                            default=json_serializer_default) # Use the default serializer

            # Log the JSON output (or a preview if too large)
            logging.info("--- Generated JSON Output ---")
            max_log_json_len = 5000
            if len(json_output_string) <= max_log_json_len:
                logging.info(json_output_string)
            else:
                logging.info(f"JSON output is large ({len(json_output_string)} chars). Logging preview:")
                logging.info(json_output_string[:max_log_json_len] + "\n... (JSON output truncated in log)")

            # Save the JSON to a file
            json_output_filename = f"output_{os.path.splitext(input_filename)[0]}.json"
            try:
                with open(json_output_filename, 'w', encoding='utf-8') as f_json:
                     f_json.write(json_output_string)
                logging.info(f"Successfully saved JSON output to '{json_output_filename}'")
            except IOError as io_err:
                logging.error(f"Failed to write JSON output to file '{json_output_filename}': {io_err}")
            except Exception as write_err:
                 logging.error(f"An unexpected error occurred while writing JSON file: {write_err}", exc_info=True)

        except TypeError as json_err:
            logging.error(f"Failed to serialize data to JSON: {json_err}. Check data types and default handler.", exc_info=True)
        except Exception as e:
            logging.error(f"An unexpected error occurred during JSON generation: {e}", exc_info=True)
        # --- End JSON Generation ---


        logging.info("--- Invoice Automation Finished Successfully ---")

    except FileNotFoundError as e: logging.error(f"Input file error: {e}")
    except RuntimeError as e: logging.error(f"Processing halted due to critical error: {e}")
    except Exception as e: logging.error(f"An unexpected error occurred in the main script execution: {e}", exc_info=True)
    finally:
        if handler:
            handler.close()
        logging.info("--- Automation Run Complete ---")


if __name__ == "__main__":
    # Make sure config.py exists and defines necessary variables like:
    # INPUT_EXCEL_FILE, SHEET_NAME, HEADER_IDENTIFICATION_PATTERN,
    # HEADER_SEARCH_ROW_RANGE, HEADER_SEARCH_COL_RANGE,
    # COLUMNS_TO_DISTRIBUTE, DISTRIBUTION_BASIS_COLUMN,
    # CUSTOM_AGGREGATION_WORKBOOK_PREFIXES
    run_invoice_automation()

# --- END OF FULL FILE: main.py ---