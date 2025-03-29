# --- START OF FILE main.py ---

# --- START OF FULL FILE: main.py ---

import logging
import pprint # For nicely printing dictionaries
import re     # Keep import re, potentially used internally or future use
import decimal # Import decimal for type hints and potential direct use
import os     # *** Import os module ***
from typing import Dict, List, Any, Optional, Tuple # Import necessary types

# Import from our refactored modules
import config as cfg # Use alias for clarity
from excel_handler import ExcelHandler
import sheet_parser
import data_processor # Includes all processing functions

# Configure logging (Set to DEBUG for detailed investigation)
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s') # USE DEBUG

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs

def run_invoice_automation():
    """Main function to find tables, extract, and process data for each."""
    logging.info("--- Starting Invoice Automation ---")
    handler = None  # Initialize handler outside try block
    actual_sheet_name = None # Still useful to know which sheet was processed

    # Dictionary to store the final processed data (post-distribution) per table
    processed_tables: Dict[int, Dict[str, Any]] = {}
    # Dictionary to store raw extracted data per table
    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}

    # *** Global dictionaries for aggregation results ***
    global_standard_aggregated_sqft_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal]], decimal.Decimal] = {}
    global_custom_aggregation_results: Dict[Tuple[Any, Any], Dict[str, decimal.Decimal]] = {}
    # Flag to track which aggregation mode is used
    aggregation_mode_used = "standard" # Default

    # --- Determine Aggregation Strategy based on WORKBOOK FILENAME ---
    # Do this BEFORE loading the sheet, as it only depends on config
    use_custom_aggregation = False
    input_filename = os.path.basename(cfg.INPUT_EXCEL_FILE) # Get filename part only
    logging.info(f"Checking workbook filename '{input_filename}' for custom aggregation trigger.")
    for prefix in cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES: # Use renamed config variable
        # Use case-sensitive comparison as defined in config comment
        if input_filename.startswith(prefix):
            use_custom_aggregation = True
            aggregation_mode_used = "custom"
            logging.info(f"Workbook filename '{input_filename}' starts with prefix '{prefix}'. Will use CUSTOM aggregation (by PO, Item for SQFT & Amount).")
            break # Found a match
    if not use_custom_aggregation:
         logging.info(f"Workbook filename '{input_filename}' does not match custom prefixes {cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES}. Will use STANDARD aggregation (by PO, Item, Price for SQFT).")
         aggregation_mode_used = "standard"
    # ----------------------------------------------------------------

    try:
        # --- 1. Load Excel Sheet ---
        logging.info(f"Loading workbook: {cfg.INPUT_EXCEL_FILE}")
        handler = ExcelHandler(cfg.INPUT_EXCEL_FILE)
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True) # data_only=True is crucial

        if sheet is None:
            raise RuntimeError(f"Failed to load sheet from '{cfg.INPUT_EXCEL_FILE}'. Exiting.")
        else:
            # Store the actual sheet name that was loaded (still useful info)
            actual_sheet_name = sheet.title
            logging.info(f"Successfully loaded worksheet: '{actual_sheet_name}' from workbook '{input_filename}'")

        # --- Aggregation Strategy was determined above ---

        # --- 2. Find All Header Rows ---
        logging.info("Searching for all header rows...")
        header_rows = sheet_parser.find_all_header_rows(
            sheet,
            cfg.HEADER_IDENTIFICATION_PATTERN,
            cfg.HEADER_SEARCH_ROW_RANGE,
            cfg.HEADER_SEARCH_COL_RANGE
        )

        if not header_rows:
             raise RuntimeError("Could not find any header rows matching the pattern. Cannot proceed.")

        logging.info(f"Found {len(header_rows)} potential header row(s) at: {header_rows}")

        # --- 3. Map Columns (Based on the FIRST header) ---
        first_header_row = header_rows[0]
        logging.info(f"Mapping columns based on the first identified header row ({first_header_row})...")
        column_mapping = sheet_parser.map_columns_to_headers(
            sheet,
            first_header_row,
            cfg.HEADER_SEARCH_COL_RANGE
        )

        if not column_mapping:
             raise RuntimeError("Failed to map required columns from the first header row. Check header names and config.")
        logging.debug(f"Successfully mapped {len(column_mapping)} columns:\n{pprint.pformat(column_mapping)}")

        # --- 4. Extract Data for All Tables ---
        logging.info("Extracting data for all identified tables using the derived column mapping...")
        all_tables_data = sheet_parser.extract_multiple_tables(
            sheet,
            header_rows,
            column_mapping
        )

        logging.debug(f"--- Raw Extracted Data ({len(all_tables_data)} Table(s)) ---\n{pprint.pformat(all_tables_data)}")
        if not all_tables_data:
             logging.warning("Extraction resulted in an empty data structure. Processing will stop.")
             # return

        # --- 5. Process Each Table Individually ---
        logging.info(f"--- Starting Data Processing Loop for {len(all_tables_data)} Extracted Table(s) ---")
        for table_index, raw_data_dict in all_tables_data.items():
            # [ ... CBM and Distribution steps remain unchanged ... ]
            current_table_data = all_tables_data.get(table_index)
            if current_table_data is None:
                 logging.error(f"!!! Internal Error: Skipping processing for missing table_index {table_index}.")
                 continue

            logging.info(f"--- Processing Table Index {table_index} ---")
            logging.debug(f"Initial data sample (first 5 rows):\n{pprint.pformat({k: v[:5] for k, v in current_table_data.items() if isinstance(v, list)})}")

            if not isinstance(current_table_data, dict) or not current_table_data or not any(isinstance(v, list) and v for v in current_table_data.values()):
                logging.warning(f"Table {table_index} empty or invalid. Skipping CBM, Distribution, and Aggregation.")
                processed_tables[table_index] = current_table_data
                continue

            # --- 5a. Pre-process: Calculate CBM ---
            logging.info(f"Table {table_index}: Calculating CBM values...")
            try:
                 data_after_cbm = data_processor.process_cbm_column(current_table_data) # Modifies current_table_data
            except Exception as cbm_e:
                 logging.error(f"Error during CBM calculation for Table {table_index}: {cbm_e}", exc_info=True)
                 data_after_cbm = current_table_data # Use potentially modified state
                 logging.warning(f"Proceeding for Table {table_index} using data state after CBM error.")

            # --- 5b. Process: Distribute Values ---
            logging.info(f"Table {table_index}: Distributing values...")
            try:
                data_after_distribution = data_processor.distribute_values(
                    data_after_cbm,
                    cfg.COLUMNS_TO_DISTRIBUTE,
                    cfg.DISTRIBUTION_BASIS_COLUMN
                )
                processed_tables[table_index] = data_after_distribution # Store successful distribution

            except data_processor.ProcessingError as pe:
                 logging.error(f"Value distribution failed for Table {table_index}: {pe}. Storing data state *before* distribution.")
                 processed_tables[table_index] = data_after_cbm # Fallback
                 logging.warning(f"Skipping Aggregation update for Table {table_index} due to distribution error.")
                 continue # Move to next table
            except Exception as dist_e:
                logging.error(f"Unexpected error during value distribution for Table {table_index}: {dist_e}", exc_info=True)
                processed_tables[table_index] = data_after_cbm # Fallback
                logging.warning(f"Skipping Aggregation update for Table {table_index} due to distribution error.")
                continue # Move to next table

            # --- 5c. Process: Aggregate (Conditional - decision already made) ---
            data_for_aggregation = processed_tables.get(table_index)

            if isinstance(data_for_aggregation, dict) and data_for_aggregation:
                if use_custom_aggregation: # Use the flag set earlier
                    logging.info(f"Table {table_index}: Updating global CUSTOM aggregation (SQFT & Amount)...")
                    try:
                        data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results)
                        logging.info(f"Table {table_index}: Successfully updated global CUSTOM aggregation map.")
                    except Exception as agg_e:
                        logging.error(f"Global CUSTOM aggregation update failed for Table {table_index}: {agg_e}", exc_info=True)
                else:
                    logging.info(f"Table {table_index}: Updating global STANDARD SQFT aggregation...")
                    try:
                        data_processor.aggregate_sqft_by_po_item_price(data_for_aggregation, global_standard_aggregated_sqft_results)
                        logging.info(f"Table {table_index}: Successfully updated global STANDARD SQFT aggregation map.")
                    except Exception as agg_e:
                        logging.error(f"Global STANDARD SQFT aggregation update failed for Table {table_index}: {agg_e}", exc_info=True)
            else:
                 logging.warning(f"Table {table_index}: Skipping aggregation update because processed data is invalid/empty.")

            logging.info(f"--- Finished Processing All Steps for Table Index {table_index} ---")


        # --- 6. Output / Further Steps ---
        logging.info("--- All Table Processing Loops Completed ---")
        logging.info(f"Final processed data structure contains results for {len(processed_tables)} table index(es): {list(processed_tables.keys())}")
        # Log based on workbook filename now
        logging.info(f"Aggregation mode used based on workbook '{input_filename}': {aggregation_mode_used.upper()}")

        # --- Log Final Processed Data --- #
        logging.debug("--- Final Processed Data (Post-Distribution, All Tables) ---\n{pprint.pformat(processed_tables)}")
        # ... (rest of processed data logging) ...

        # --- Log Final Aggregation Results (Conditional) --- #
        if aggregation_mode_used == "custom":
            logging.info(f"--- Final Global CUSTOM Aggregation Results (Workbook: '{input_filename}') ---")
            if global_custom_aggregation_results:
                agg_output = pprint.pformat(global_custom_aggregation_results)
                if len(agg_output) > MAX_LOG_DICT_LEN and logging.getLogger().getEffectiveLevel() > logging.DEBUG:
                     agg_output = agg_output[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
                logging.info(f"Custom aggregation resulted in {len(global_custom_aggregation_results)} unique (PO, Item) keys.")
                logging.info(f"\n{agg_output}")
                logging.debug(f"Full Final Global Custom Aggregation:\n{pprint.pformat(global_custom_aggregation_results)}")
            else:
                logging.info("No custom aggregated results were generated.")
        else: # Standard aggregation was used
            logging.info(f"--- Final Global STANDARD Aggregation Results (Workbook: '{input_filename}') ---")
            if global_standard_aggregated_sqft_results:
                agg_output = pprint.pformat(global_standard_aggregated_sqft_results)
                if len(agg_output) > MAX_LOG_DICT_LEN and logging.getLogger().getEffectiveLevel() > logging.DEBUG:
                     agg_output = agg_output[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
                logging.info(f"Standard aggregation resulted in {len(global_standard_aggregated_sqft_results)} unique (PO, Item, Price) keys.")
                logging.info(f"\n{agg_output}")
                logging.debug(f"Full Final Global Standard Aggregation:\n{pprint.pformat(global_standard_aggregated_sqft_results)}")
            else:
                logging.info("No standard aggregated SQFT results were generated.")
        ########################################################

        # --- Future steps example: Save results ---
        # Consider saving the relevant aggregation result based on the mode
        # import pickle
        # output_file = f"processed_{os.path.splitext(input_filename)[0]}.pkl" # Dynamic output name
        # try:
        #     combined_results = {
        #         "processed_tables": processed_tables,
        #         "aggregation_mode": aggregation_mode_used,
        #         "workbook_filename": input_filename,
        #         "worksheet_name": actual_sheet_name,
        #         "standard_aggregation_results": global_standard_aggregated_sqft_results if aggregation_mode_used == 'standard' else {},
        #         "custom_aggregation_results": global_custom_aggregation_results if aggregation_mode_used == 'custom' else {}
        #     }
        #     with open(output_file, 'wb') as f_pickle:
        #         pickle.dump(combined_results, f_pickle)
        #     logging.info(f"Successfully saved combined processed data to {output_file}")
        # except Exception as pickle_e:
        #     logging.error(f"Failed to save data to pickle file {output_file}: {pickle_e}")


        logging.info("--- Invoice Automation Finished Successfully ---")

    except FileNotFoundError as e:
        logging.error(f"Input file error: {e}")
    except RuntimeError as e:
        logging.error(f"Processing halted due to critical error: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred in the main script execution: {e}", exc_info=True)
    finally:
        if handler:
            handler.close()
        logging.info("--- Automation Run Complete ---")


if __name__ == "__main__":
    run_invoice_automation()

# --- END OF FULL FILE: main.py ---