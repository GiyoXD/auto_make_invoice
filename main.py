import logging
import pprint # For nicely printing dictionaries
import re # Keep import re, used internally by data_processor potentially

# Import from our refactored modules
import config as cfg # Use alias for clarity
from excel_handler import ExcelHandler
import sheet_parser
import data_processor # data_processor now includes CBM processing

# Configure logging (consider setting level via config or environment variable)
# Set to DEBUG to see detailed CBM calculations or distribution steps
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


def run_invoice_automation():
    """Main function to find tables, extract, and process data for each."""
    logging.info("--- Starting Invoice Automation ---")
    handler = None  # Initialize handler outside try block

    try:
        # --- 1. Load Excel Sheet ---
        handler = ExcelHandler(cfg.INPUT_EXCEL_FILE)
        # Load with data_only=True to get values, not formulas
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True)

        if sheet is None:
            # Error already logged by ExcelHandler
            raise RuntimeError(f"Failed to load sheet from '{cfg.INPUT_EXCEL_FILE}'. Exiting.")

        # --- 2. Find All Header Rows ---
        logging.info("Searching for all header rows...")
        # Use the function to find all potential header rows
        header_rows = sheet_parser.find_all_header_rows(
            sheet,
            cfg.HEADER_IDENTIFICATION_PATTERN,
            cfg.HEADER_SEARCH_ROW_RANGE, # Use configured search depth
            cfg.HEADER_SEARCH_COL_RANGE
        )

        # Check if any headers were found
        if not header_rows:
             # Warning/Error already logged by find_all_header_rows
             raise RuntimeError("Could not find any header rows matching the pattern. Cannot proceed.")

        logging.info(f"Found {len(header_rows)} header row(s) at: {header_rows}")

        # --- 3. Map Columns (Based on the FIRST header) ---
        # Assumption: All tables have the same column structure defined by the first header.
        logging.info(f"Mapping columns based on the first header row ({header_rows[0]})...")
        column_mapping = sheet_parser.map_columns_to_headers(
            sheet,
            header_rows[0], # Use the first found header row for mapping
            cfg.HEADER_SEARCH_COL_RANGE
        )

        # Check if mapping was successful (found essential columns)
        if not column_mapping:
             # Error/Warning logged by map_columns_to_headers
             raise RuntimeError("Failed to map required columns from the first header. Check header names and config.")
        logging.info(f"Successfully mapped {len(column_mapping)} columns:")
        logging.info(pprint.pformat(column_mapping)) # Log the mapping for verification


        # --- 4. Extract Data for All Tables ---
        logging.info("Extracting data for all identified tables...")
        # Use the extraction function which returns a dict of tables
        all_tables_data = sheet_parser.extract_multiple_tables(
            sheet,
            header_rows, # Pass the list of all found header rows
            column_mapping # Pass the single mapping derived from the first header
        )

        if not all_tables_data:
             logging.warning("No data was extracted from any tables. The sheet might be empty below headers or stop condition met early.")
             # Depending on requirements, this might be ok or an error.


        # --- 5. Process Each Table Individually ---
        processed_tables = {} # Dictionary to store processed data for each table
        logging.info("--- Starting Data Processing for Each Table ---")
        for table_index, raw_data in all_tables_data.items():
            logging.info(f"--- Processing Table {table_index} ---")

            # Check if the raw_data dictionary actually contains any non-empty lists or actual data
            if not any(raw_data.values()):
                logging.warning(f"Table {table_index} contains no data columns/rows. Skipping processing.")
                processed_tables[table_index] = raw_data # Store empty data if needed
                continue
            # More robust check: Ensure there's at least one non-None value in any list
            has_actual_data = any(v is not None for v_list in raw_data.values() for v in v_list)
            if not has_actual_data:
                 logging.warning(f"Table {table_index} extracted but appears to contain only empty cells. Skipping processing.")
                 processed_tables[table_index] = raw_data # Store empty data
                 continue


            # --- 5a. Pre-process: Calculate CBM --- ####################################
            logging.info(f"Pre-processing Table {table_index}: Calculating CBM values...")
            # process_cbm_column modifies raw_data in place and returns it
            try:
                 raw_data_with_cbm = data_processor.process_cbm_column(raw_data) # CALL THE NEW FUNCTION
            except Exception as cbm_e:
                 logging.error(f"Error during CBM calculation for Table {table_index}: {cbm_e}", exc_info=True)
                 # Decide how to handle CBM errors: skip table, continue without CBM, etc.
                 # For now, let's continue but log the error and potentially skip distribution for CBM
                 raw_data_with_cbm = raw_data # Fallback to original data if CBM calc fails severely
                 logging.warning(f"Proceeding for Table {table_index} without successfully calculated CBM values.")
            ###########################################################################


            # --- 5b. Process: Distribute Values ---
            logging.info(f"Processing Table {table_index}: Distributing values...")
            try:
                # Pass the data (now potentially with calculated CBM) to distribute_values
                processed_data = data_processor.distribute_values(
                    raw_data_with_cbm,         # Use the result from CBM processing step
                    cfg.COLUMNS_TO_DISTRIBUTE, # From config.py (e.g., ["net", "gross", "cbm"])
                    cfg.DISTRIBUTION_BASIS_COLUMN # From config.py (e.g., "pcs")
                )
                processed_tables[table_index] = processed_data # Store final processed data for this table
                logging.info(f"--- Finished Processing Table {table_index} ---")
                # Optionally log the processed data for this table (can be verbose)
                logging.debug(f"Processed Data for Table {table_index}:\n{pprint.pformat(processed_data)}")

            except data_processor.ProcessingError as pe:
                 # Log error specific to this table during distribution
                 logging.error(f"Value distribution failed for Table {table_index}: {pe}")
                 # Option 1: Store raw data instead
                 # processed_tables[table_index] = raw_data_with_cbm
                 # Option 2: Skip this table
                 # processed_tables[table_index] = None # Or some indicator of failure
                 # Option 3: Re-raise to stop everything (current behavior)
                 raise
            except Exception as e:
                # Catch any other unexpected errors during distribution
                logging.error(f"An unexpected error occurred during value distribution for Table {table_index}: {e}", exc_info=True)
                raise # Stop execution on unexpected errors


        # --- 6. Output / Further Steps ---
        logging.info("--- All Tables Processed ---")
        # You now have the 'processed_tables' dictionary containing the processed data
        # for each table, keyed by table_index (1, 2, ...).
        logging.info(f"Final processed data structure contains {len(processed_tables)} tables.")

        # Example: Print the processed data for inspection (can be large)
        logging.info("--- Final Processed Data (All Tables) ---")
        # Use pprint for readable output of the nested dictionary structure
        logging.info("\n" + pprint.pformat(processed_tables))

        # Future steps:
        # - Save processed_tables to a pickle file (cfg.OUTPUT_PICKLE_FILE)
        #   import pickle
        #   try:
        #       with open(cfg.OUTPUT_PICKLE_FILE, 'wb') as f_pickle:
        #           pickle.dump(processed_tables, f_pickle)
        #       logging.info(f"Successfully saved processed data to {cfg.OUTPUT_PICKLE_FILE}")
        #   except Exception as pickle_e:
        #       logging.error(f"Failed to save data to pickle file {cfg.OUTPUT_PICKLE_FILE}: {pickle_e}")
        #
        # - Convert each table's data to a Pandas DataFrame for analysis/output
        # - Generate separate output files (CSV, Excel) per table
        # - Combine data in a specific way before output

        logging.info("--- Invoice Automation Finished Successfully ---")

    except FileNotFoundError as e:
        logging.error(f"Input file error: {e}")
    except RuntimeError as e:
        # Catch specific errors raised for flow control (e.g., no headers found)
        logging.error(f"Processing halted: {e}")
    except Exception as e:
        # Catch any other unexpected errors during setup or extraction
        logging.error(f"An unexpected error occurred: {e}", exc_info=True)
    finally:
        # --- Clean up ---
        if handler:
            handler.close() # Release file handle reference
        logging.info("--- Automation Run Complete ---")


if __name__ == "__main__":
    run_invoice_automation()

# --- END OF FILE: main.py ---