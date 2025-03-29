# --- START OF FULL FILE: main.py ---

import logging
import pprint
import re
import decimal
import os
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
FOB_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\PO2)
FOB_INTRA_CHUNK_SEPARATOR = "\\"  # Separator within a group (e.g., backslash)
FOB_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline)


# Type alias for the two possible initial aggregation structures
InitialAggregationResults = Union[
    Dict[Tuple[Any, Any, Optional[decimal.Decimal]], Dict[str, decimal.Decimal]], # Standard Result
    Dict[Tuple[Any, Any], Dict[str, decimal.Decimal]]                             # Custom Result
]
# Type alias for the FOB compounding result structure
FobCompoundingResult = Dict[str, Union[str, decimal.Decimal]]


# *** FOB Compounding Function with Chunking ***
def perform_fob_compounding(
    initial_results: InitialAggregationResults,
    aggregation_mode: str # 'standard' or 'custom' -> Needed to parse input keys correctly
) -> Optional[FobCompoundingResult]:
    """
    Performs FOB Compounding from standard or custom aggregation results.

    This function *always* runs after the initial aggregation step.
    It combines all unique POs and Items into formatted SINGLE strings
    (chunked based on constants).
    It sums all SQFT and Amount values. Returns a single dictionary record.

    Args:
        initial_results: The dictionary from either standard or custom aggregation.
        aggregation_mode: String ('standard' or 'custom') indicating how to parse
                          the keys of initial_results.

    Returns:
        A single dictionary with 'combined_po', 'combined_item', 'total_sqft',
        'total_amount', or a default structure if input is empty/invalid. Returns
        None only on critical internal errors.
    """
    prefix = "[perform_fob_compounding]"
    logging.info(f"{prefix} Starting FOB Compounding (runs always). Using input from '{aggregation_mode}' aggregation.")

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

        # Extract PO and Item from the key based on the initial aggregation mode
        try:
            if aggregation_mode == 'standard':
                if len(key) == 3: po_key_val, item_key_val, _ = key
                else: raise ValueError("Invalid standard key length")
            elif aggregation_mode == 'custom':
                if len(key) == 2: po_key_val, item_key_val = key
                else: raise ValueError("Invalid custom key length")
            else:
                logging.error(f"{prefix} Unknown aggregation mode '{aggregation_mode}' passed. Halting compounding.")
                return None # Indicate critical error
        except Exception as e:
             logging.warning(f"{prefix} Error unpacking key {key} in '{aggregation_mode}' mode: {e}. Skipping entry.")
             continue

        # --- Collect unique POs/Items ---
        po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
        item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
        unique_pos.add(po_str)
        unique_items.add(item_str)

        # --- Sum numeric values ---
        sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
        amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))
        if not isinstance(sqft_sum, decimal.Decimal):
             logging.warning(f"{prefix} Invalid SQFT sum type for key {key}: {type(sqft_sum)}. Using 0.")
             sqft_sum = decimal.Decimal(0)
        if not isinstance(amount_sum, decimal.Decimal):
             logging.warning(f"{prefix} Invalid Amount sum type for key {key}: {type(amount_sum)}. Using 0.")
             amount_sum = decimal.Decimal(0)
        total_sqft += sqft_sum
        total_amount += amount_sum

    logging.debug(f"{prefix} Finished processing entries.")
    logging.debug(f"{prefix} Unique POs collected: {unique_pos}")
    logging.debug(f"{prefix} Unique Items collected: {unique_items}")

    # --- Final Combination with Chunking ---
    sorted_pos = sorted(list(unique_pos))
    sorted_items = sorted(list(unique_items))
    logging.debug(f"{prefix} Sorted PO list before chunking: {sorted_pos}")
    logging.debug(f"{prefix} Sorted Item list before chunking: {sorted_items}")

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
    logging.debug(f"{prefix} Final combined_po_string (Type: {type(combined_po_string)}): '{combined_po_string}'")
    logging.debug(f"{prefix} Final combined_item_string (Type: {type(combined_item_string)}): '{combined_item_string}'")

    # Construct Result Dictionary
    fob_compounded_result: FobCompoundingResult = {
        'combined_po': combined_po_string,    # Now formatted with chunks
        'combined_item': combined_item_string, # Now formatted with chunks
        'total_sqft': total_sqft,
        'total_amount': total_amount
    }

    logging.info(f"{prefix} FOB Compounding complete.")
    return fob_compounded_result


def run_invoice_automation():
    """Main function to find tables, extract, and process data for each."""
    logging.info("--- Starting Invoice Automation ---")
    handler = None
    actual_sheet_name = None

    processed_tables: Dict[int, Dict[str, Any]] = {}
    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}

    # Global dictionaries for initial aggregation results
    global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal]], Dict[str, decimal.Decimal]] = {}
    global_custom_aggregation_results: Dict[Tuple[Any, Any], Dict[str, decimal.Decimal]] = {}
    # Global variable for the final single FOB compounded result
    global_fob_compounded_result: Optional[FobCompoundingResult] = None

    aggregation_mode_used = "standard" # Default, determines initial aggregation type

    # --- Determine Initial Aggregation Strategy ---
    use_custom_aggregation = False
    input_filename = os.path.basename(cfg.INPUT_EXCEL_FILE)
    logging.info(f"Checking workbook filename '{input_filename}' for custom aggregation trigger.")
    for prefix in cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES:
        if input_filename.startswith(prefix):
            use_custom_aggregation = True
            aggregation_mode_used = "custom"
            logging.info(f"Workbook filename matches prefix '{prefix}'. Using CUSTOM initial aggregation.")
            break
    if not use_custom_aggregation:
         logging.info(f"Workbook filename does not match custom prefixes. Using STANDARD initial aggregation.")
         aggregation_mode_used = "standard"
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
        if 'amount' not in column_mapping: raise RuntimeError("Essential 'amount' column mapping failed.")

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
                logging.warning(f"Table {table_index} empty or invalid. Skipping steps.")
                processed_tables[table_index] = current_table_data
                continue

            # 5a. CBM Calculation
            logging.info(f"Table {table_index}: Calculating CBM values...")
            try: data_after_cbm = data_processor.process_cbm_column(current_table_data)
            except Exception as e: logging.error(f"CBM calc error Table {table_index}: {e}", exc_info=True); data_after_cbm = current_table_data

            # 5b. Distribution
            logging.info(f"Table {table_index}: Distributing values...")
            try:
                data_after_distribution = data_processor.distribute_values(data_after_cbm, cfg.COLUMNS_TO_DISTRIBUTE, cfg.DISTRIBUTION_BASIS_COLUMN)
                processed_tables[table_index] = data_after_distribution
            except data_processor.ProcessingError as pe:
                logging.error(f"Distribution failed Table {table_index}: {pe}. Storing pre-distribution data.")
                processed_tables[table_index] = data_after_cbm; continue
            except Exception as e:
                logging.error(f"Unexpected distribution error Table {table_index}: {e}", exc_info=True)
                processed_tables[table_index] = data_after_cbm; continue

            # 5c. Initial Aggregation (Standard or Custom)
            data_for_aggregation = processed_tables.get(table_index)
            if isinstance(data_for_aggregation, dict) and data_for_aggregation:
                try:
                    if use_custom_aggregation:
                        logging.info(f"Table {table_index}: Updating global CUSTOM aggregation...")
                        data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results)
                        logging.info(f"Table {table_index}: CUSTOM aggregation map updated.")
                    else:
                        logging.info(f"Table {table_index}: Updating global STANDARD aggregation...")
                        data_processor.aggregate_standard_by_po_item_price(data_for_aggregation, global_standard_aggregation_results)
                        logging.info(f"Table {table_index}: STANDARD aggregation map updated.")
                except Exception as agg_e:
                    logging.error(f"Global {aggregation_mode_used.upper()} aggregation update failed for Table {table_index}: {agg_e}", exc_info=True)
            else:
                 logging.warning(f"Table {table_index}: Skipping initial aggregation update (processed data invalid/empty after distribution).")

            logging.info(f"--- Finished Processing All Steps for Table Index {table_index} ---")
        # --- End Processing Loop ---


        # --- 6. Post-Loop: Perform FOB Compounding (ALWAYS RUNS) ---
        logging.info("--- All Table Processing Loops Completed ---")
        logging.info("--- Performing Final FOB Compounding (Always Runs) ---")
        try:
            initial_agg_data_source = global_custom_aggregation_results if use_custom_aggregation else global_standard_aggregation_results
            global_fob_compounded_result = perform_fob_compounding(
                initial_agg_data_source,
                aggregation_mode_used # Pass mode to help parse input keys
            )
            logging.info("--- FOB Compounding Finished ---")
        except Exception as fob_e:
             logging.error(f"An error occurred during the final FOB Compounding step: {fob_e}", exc_info=True)
             logging.error("FOB Compounding results may be incomplete or missing.")


        # --- 7. Output / Further Steps ---
        logging.info(f"Final processed data structure contains {len(processed_tables)} table(s).")
        logging.info(f"Initial aggregation mode used: {aggregation_mode_used.upper()}")

        # --- Log Initial Aggregation Results (DEBUG Level) ---
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
            if aggregation_mode_used == "custom":
                 log_str = pprint.pformat(global_custom_aggregation_results)
                 if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
                 logging.debug(f"--- Full Global CUSTOM Aggregation Results (Input to FOB) ---\n{log_str}")
            else:
                 log_str = pprint.pformat(global_standard_aggregation_results)
                 if len(log_str) > MAX_LOG_DICT_LEN: log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
                 logging.debug(f"--- Full Global STANDARD Aggregation Results (Input to FOB) ---\n{log_str}")

        # --- Log Final FOB Compounded Result (INFO Level) - REFINED --- #
        logging.info(f"--- Final FOB Compounded Result (Workbook: '{input_filename}') ---")
        if global_fob_compounded_result is not None:
            # Get the final values
            po_string_value = global_fob_compounded_result.get('combined_po', '<Not Found>')
            item_string_value = global_fob_compounded_result.get('combined_item', '<Not Found>')
            total_sqft_value = global_fob_compounded_result.get('total_sqft', 'N/A')
            total_amount_value = global_fob_compounded_result.get('total_amount', 'N/A')

            # Log each component explicitly for clarity
            logging.info(f"Combined POs (Type: {type(po_string_value)}):")
            logging.info(f"  repr(): {repr(po_string_value)}") # Shows literal \n and \\
            logging.info(f"  Raw Value:\n{po_string_value}")   # Renders multi-line if \n present
            logging.info("-" * 30)

            logging.info(f"Combined Items (Type: {type(item_string_value)}):")
            logging.info(f"  repr(): {repr(item_string_value)}")
            logging.info(f"  Raw Value:\n{item_string_value}")
            logging.info("-" * 30)

            logging.info(f"Total SQFT: {total_sqft_value} (Type: {type(total_sqft_value)})")
            logging.info(f"Total Amount: {total_amount_value} (Type: {type(total_amount_value)})")
            logging.info("-" * 30)

            # Optional notes based on newline presence (less informative now with chunks)
            # if isinstance(po_string_value, str) and FOB_INTER_CHUNK_SEPARATOR not in po_string_value:
            #     logging.info(f"NOTE: 'combined_po' fits in one chunk group.")
            # if isinstance(item_string_value, str) and FOB_INTER_CHUNK_SEPARATOR not in item_string_value:
            #     logging.info(f"NOTE: 'combined_item' fits in one chunk group.")
        else:
            logging.error("FOB Compounding result is None or was not set, potentially due to an error during compounding.")
        # --- End Final Logging ---

        # --- Optional: Save results ---
        # import pickle
        # output_file = f"processed_{os.path.splitext(input_filename)[0]}.pkl"
        # try:
        #     combined_results = {
        #         # "processed_tables": processed_tables,
        #         "aggregation_mode": aggregation_mode_used,
        #         "workbook_filename": input_filename,
        #         "worksheet_name": actual_sheet_name,
        #         # "standard_aggregation_results": global_standard_aggregation_results,
        #         # "custom_aggregation_results": global_custom_aggregation_results,
        #         "fob_compounded_result": global_fob_compounded_result # The main result
        #     }
        #     with open(output_file, 'wb') as f_pickle:
        #         pickle.dump(combined_results, f_pickle)
        #     logging.info(f"Successfully saved compounded data to {output_file}")
        # except Exception as pickle_e:
        #     logging.error(f"Failed to save data to pickle file {output_file}: {pickle_e}")


        logging.info("--- Invoice Automation Finished Successfully ---")

    except FileNotFoundError as e: logging.error(f"Input file error: {e}")
    except RuntimeError as e: logging.error(f"Processing halted due to critical error: {e}")
    except Exception as e: logging.error(f"An unexpected error occurred in the main script execution: {e}", exc_info=True)
    finally:
        if handler:
            handler.close()
        logging.info("--- Automation Run Complete ---")


if __name__ == "__main__":
    run_invoice_automation()

# --- END OF FULL FILE: main.py ---