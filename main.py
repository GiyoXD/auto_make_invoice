# --- START OF FILE main.py ---

import logging
import pprint # For nicely printing dictionaries

# Import from our refactored modules
import config as cfg # Use alias for clarity
from excel_handler import ExcelHandler
import sheet_parser
import data_processor

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run_invoice_automation():
    """Main function to run the invoice data extraction and processing."""
    logging.info("--- Starting Invoice Automation ---")
    handler = None  # Initialize handler

    try:
        # --- 1. Load Excel Sheet ---
        handler = ExcelHandler(cfg.INPUT_EXCEL_FILE)
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True) # data_only=True usually preferred

        if sheet is None:
            raise RuntimeError(f"Failed to load sheet from '{cfg.INPUT_EXCEL_FILE}'.")

        # --- 2. Parse Sheet Structure ---
        logging.info("Finding header row...")
        header_row = sheet_parser.find_header_row(
            sheet,
            cfg.HEADER_IDENTIFICATION_PATTERN,
            cfg.HEADER_SEARCH_ROW_RANGE,
            cfg.HEADER_SEARCH_COL_RANGE
        )
        if header_row is None:
             raise RuntimeError("Could not find the header row. Check config pattern and search range.")

        logging.info("Mapping columns to headers...")
        column_mapping = sheet_parser.map_columns_to_headers(
            sheet,
            header_row,
            cfg.HEADER_SEARCH_COL_RANGE
        )
        if not column_mapping:
             raise RuntimeError("Failed to map required columns. Check header names in Excel and config.")
        logging.info(f"Column Mapping: {column_mapping}")


        # --- 3. Extract Raw Data ---
        logging.info("Extracting raw data...")
        raw_data = sheet_parser.extract_raw_data(
            sheet,
            header_row,
            column_mapping
        )
        if not any(raw_data.values()): # Check if any data was actually extracted
             logging.warning("Raw data extraction resulted in empty lists.")
             # Decide if this is an error or just an empty sheet section
             # return # Or raise error depending on expectation

        # logging.debug("--- Raw Extracted Data ---")
        # logging.debug(pprint.pformat(raw_data))


        # --- 4. Process Data (Distribute Values) ---
        logging.info("Processing data: Distributing values...")
        try:
            processed_data = data_processor.distribute_values(
                raw_data,
                cfg.COLUMNS_TO_DISTRIBUTE,
                cfg.DISTRIBUTION_BASIS_COLUMN
            )
            logging.info("--- Processed Data ---")
            # Use pprint with logging for better readability
            logging.info("\n" + pprint.pformat(processed_data))

            # --- 5. Further Steps (Future) ---
            # - Validate processed data
            # - Aggregate data if necessary
            # - Export to Pickle/CSV/Invoice Format
            # logging.info("Exporting processed data...")
            # export_data(processed_data, cfg.OUTPUT_PICKLE_FILE)

        except data_processor.ProcessingError as pe:
             logging.error(f"Data processing failed: {pe}")
             # Decide how to handle: stop, try alternative, etc.
             raise # Re-raise if it's critical

        logging.info("--- Invoice Automation Finished Successfully ---")

    except FileNotFoundError as e:
        logging.error(f"Input file error: {e}")
    except RuntimeError as e:
        logging.error(f"Runtime error during processing: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True) # Log traceback for unexpected errors
    finally:
        # --- Clean up ---
        if handler:
            handler.close()
        logging.info("--- Exiting ---")


if __name__ == "__main__":
    run_invoice_automation()

# --- END OF FILE main.py ---