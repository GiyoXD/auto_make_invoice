# --- START OF FULL FILE: main.py ---
# --- Fixed datetime JSON serialization ---

import logging
import pprint
import re
import decimal
import os
import json # Added for JSON output
import datetime # <<< ADDED IMPORT for datetime handling
import argparse # <<< ADDED IMPORT for argument parsing
from pathlib import Path # <<< ADDED IMPORT for pathlib
from typing import Dict, List, Any, Optional, Tuple, Union

# Import from our refactored modules
try:
    import config as cfg # Keep config for fallback and other settings
except ImportError:
    logging.error("Failed to import config.py. Please ensure it exists and is configured.")
    # Define dummy cfg values if needed for script to load, but it will likely fail later
    class DummyConfig:
        INPUT_EXCEL_FILE = "fallback_excel.xlsx" # Example placeholder
        SHEET_NAME = "Sheet1"
        HEADER_IDENTIFICATION_PATTERN = r"PO#" # Example
        HEADER_SEARCH_ROW_RANGE = (1, 20) # Example
        HEADER_SEARCH_COL_RANGE = (1, 30) # Example
        COLUMNS_TO_DISTRIBUTE = [] # Example
        DISTRIBUTION_BASIS_COLUMN = "SQFT" # Example
        CUSTOM_AGGREGATION_WORKBOOK_PREFIXES = ["CUST"] # eeExample
    cfg = DummyConfig()
    logging.warning("Using dummy config values due to import failure.")


from excel_handler import ExcelHandler
import sheet_parser
import data_processor # Includes all processing functions

# Configure logging (Set level as needed, DEBUG is useful)
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs (for DEBUG)

# --- Constants for FOB Compounding Formatting ---
FOB_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\\PO2)
FOB_INTRA_CHUNK_SEPARATOR = "/"  # Separator within a group (e.g., DOUBLE BACKSLASH)
FOB_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline)

# Type alias for the two possible initial aggregation structures
# UPDATED Type Alias to reflect new key structures
InitialAggregationResults = Union[
    Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]], # Standard Result (PO, Item, Price, Desc)
    Dict[Tuple[Any, Any, Optional[str], None], Dict[str, decimal.Decimal]]                             # Custom Result (PO, Item, Desc, None) - UPDATED
]
# Type alias for the FOB compounding result structure
FobCompoundingResult = Dict[str, Union[str, decimal.Decimal]]

# Type alias for the final FOB result (ALWAYS a split dict, but structure varies)
FinalFobResultType = Dict[str, FobCompoundingResult]


# *** FOB Compounding Function with Chunking ***
def perform_fob_compounding(
    initial_results: InitialAggregationResults, # Type hint updated
    aggregation_mode: str # 'standard' or 'custom' -> Needed to parse input keys correctly
) -> Optional[FinalFobResultType]: # <<< Return type is always Dict[str, ...]
    """
    Performs FOB Compounding from standard or custom aggregation results.
    - If description data IS present: Performs BUFFALO split (Groups "1" & "2").
      Uses FOB_CHUNK_SIZE=2 and FOB_INTRA_CHUNK_SEPARATOR='\\'.
    - If description data IS NOT present: Performs PO Count split (Groups "1", "2", ...).
      Uses NO_DESC_SPLIT_CHUNK_SIZE=8 and FOB_INTRA_CHUNK_SEPARATOR='\\'.
      Calculates chunk-specific totals.

    Args:
        initial_results: The dictionary from EITHER standard OR custom aggregation.
        aggregation_mode: String ('standard' or 'custom') indicating key structure.
    Returns:
        - A dictionary keyed by group/chunk index ("1", "2", ...).
        - Default structure (empty groups "1", "2") if input is empty.
        - None on critical internal errors.
    """
    prefix = "[perform_fob_compounding]"
    logging.info(f"{prefix} Starting FOB Compounding. Checking for descriptions to determine split type.")

    # Helper function for creating a default empty group result
    def default_group_result() -> FobCompoundingResult:
        return {
            'combined_po': '',
            'combined_item': '',
            'combined_description': '',
            'total_sqft': decimal.Decimal(0),
            'total_amount': decimal.Decimal(0)
        }

    # Handle empty input consistently -> returns default BUFFALO split dict
    if not initial_results:
        logging.warning(f"{prefix} Input aggregation results map is empty. Returning default empty FOB groups.")
        return {
            "1": default_group_result(), # Buffalo group
            "2": default_group_result()  # Non-Buffalo group
        }

    # --- Check if any description data exists ---
    any_description_present = False
    for key in initial_results.keys():
        desc_key_val = None
        try:
            if aggregation_mode == 'standard' and len(key) == 4:
                desc_key_val = key[3]
            elif aggregation_mode == 'custom' and len(key) == 4:
                desc_key_val = key[3]
            elif len(key) >= 4: desc_key_val = key[3]
            if desc_key_val is not None and str(desc_key_val).strip():
                any_description_present = True
                logging.debug(f"{prefix} Found description data. Will perform BUFFALO split.")
                break
        except (IndexError, TypeError): continue

    # Reusable helper function for formatting chunks
    def format_chunks(items: List[str], chunk_size: int, intra_sep: str, inter_sep: str) -> str:
        if not items:
            return ""
        processed_chunks = []
        for i in range(0, len(items), chunk_size):
            chunk = [str(item) for item in items[i:i + chunk_size]]
            joined_chunk = intra_sep.join(chunk)
            processed_chunks.append(joined_chunk)
        return inter_sep.join(processed_chunks)

    # --- Decide Execution Path --- #

    if any_description_present:
        # --- Path 1: Descriptions ARE present -> BUFFALO Split Aggregation (Chunk Size 2) ---
        logging.info(f"{prefix} Performing BUFFALO split aggregation (Chunk Size: {FOB_CHUNK_SIZE}).")
        # Initialize accumulators for BUFFALO group ("1")
        buffalo_pos = set()
        buffalo_items = set()
        buffalo_descriptions = set()
        buffalo_sqft = decimal.Decimal(0)
        buffalo_amount = decimal.Decimal(0)
        # Initialize accumulators for NON-BUFFALO group ("2")
        non_buffalo_pos = set()
        non_buffalo_items = set()
        non_buffalo_descriptions = set()
        non_buffalo_sqft = decimal.Decimal(0)
        non_buffalo_amount = decimal.Decimal(0)

        logging.debug(f"{prefix} Processing {len(initial_results)} entries for BUFFALO split.")
        for key, sums_dict in initial_results.items():
             po_key_val, item_key_val, desc_key_val = None, None, None
             try: # Extract PO, Item, Desc
                 if aggregation_mode == 'standard' and len(key) == 4:
                     po_key_val, item_key_val, _, desc_key_val = key
                 elif aggregation_mode == 'custom' and len(key) == 4:
                     po_key_val, item_key_val, _, desc_key_val = key
                 else:
                     if len(key) != 4: logging.warning(f"{prefix} Unexpected key length ({len(key)}) for key {key} in BUFFALO split mode. Trying heuristic.")
                     if len(key) >= 2: po_key_val, item_key_val = key[0], key[1]
                     if len(key) >= 4: desc_key_val = key[3]
                     if po_key_val is None or item_key_val is None:
                         logging.warning(f"{prefix} Cannot extract PO/Item/Desc reliably from key {key} in BUFFALO split mode. Skipping.")
                         continue
             except (ValueError, TypeError, IndexError) as e:
                 logging.warning(f"{prefix} Error unpacking key {key} (BUFFALO split mode): {e}. Skipping.")
                 continue

             po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
             item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
             desc_str = str(desc_key_val).strip() if desc_key_val is not None else ""
             is_buffalo = False
             if desc_str and "BUFFALO" in desc_str.upper(): is_buffalo = True
             sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
             amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))
             if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
             if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)

             if is_buffalo:
                 buffalo_pos.add(po_str)
                 buffalo_items.add(item_str)
                 buffalo_descriptions.add(desc_str)
                 buffalo_sqft += sqft_sum
                 buffalo_amount += amount_sum
             else:
                 non_buffalo_pos.add(po_str)
                 non_buffalo_items.add(item_str)
                 if desc_str: non_buffalo_descriptions.add(desc_str)
                 non_buffalo_sqft += sqft_sum
                 non_buffalo_amount += amount_sum

        logging.debug(f"{prefix} Finished processing entries for BUFFALO split.")

        # Format BUFFALO Group ("1")
        sorted_buffalo_pos = sorted(list(buffalo_pos))
        sorted_buffalo_items = sorted(list(buffalo_items))
        sorted_buffalo_descriptions = sorted([d for d in buffalo_descriptions if d])
        buffalo_result: FobCompoundingResult = {
            'combined_po': format_chunks(sorted_buffalo_pos, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR),
            'combined_item': format_chunks(sorted_buffalo_items, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR),
            'combined_description': format_chunks(sorted_buffalo_descriptions, 1, "", "\n"),
            'total_sqft': buffalo_sqft,
            'total_amount': buffalo_amount
        }
        # Format NON-BUFFALO Group ("2")
        sorted_non_buffalo_pos = sorted(list(non_buffalo_pos))
        sorted_non_buffalo_items = sorted(list(non_buffalo_items))
        sorted_non_buffalo_descriptions = sorted([d for d in non_buffalo_descriptions if d])
        non_buffalo_result: FobCompoundingResult = {
            'combined_po': format_chunks(sorted_non_buffalo_pos, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR),
            'combined_item': format_chunks(sorted_non_buffalo_items, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR),
            'combined_description': format_chunks(sorted_non_buffalo_descriptions, 1, "", "\n"),
            'total_sqft': non_buffalo_sqft,
            'total_amount': non_buffalo_amount
        }
        # Construct Final Result Dictionary for BUFFALO Split Case
        final_buffalo_split_result: FinalFobResultType = {
            "1": buffalo_result,
            "2": non_buffalo_result
        }
        logging.info(f"{prefix} BUFFALO split FOB Compounding complete.")
        return final_buffalo_split_result
        # --- End Path 1 (BUFFALO Split) --- #

    else:
        # --- Path 2: Descriptions are NOT present -> PO Count Split Aggregation ---
        # Totals are calculated based on conceptual groups of 8 POs.
        # Final string formatting uses chunk size 2.
        PO_GROUPING_FOR_TOTALS = 5 # Define the size for grouping totals
        logging.info(f"{prefix} No description data found. Performing PO count split aggregation.")
        logging.info(f"{prefix}   - Totals calculated per group of {PO_GROUPING_FOR_TOTALS} POs.")
        logging.info(f"{prefix}   - String formatting uses chunk size {FOB_CHUNK_SIZE} and separator '{FOB_INTRA_CHUNK_SEPARATOR}'.")

        # Step 1: Aggregate data by PO
        po_data_aggregation: Dict[str, Dict[str, Union[set, decimal.Decimal]]] = {}
        logging.debug(f"{prefix} Pass 1: Aggregating SQFT/Amount/Items per PO.")
        for key, sums_dict in initial_results.items():
             po_key_val, item_key_val = None, None
             try: # Extract PO/Item
                 if len(key) >= 2: po_key_val, item_key_val = key[0], key[1]
                 else: continue
             except (TypeError, IndexError) as e: continue # Ignore errors in pass 1

             po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
             item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
             sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
             amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))
             if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
             if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)

             if po_str not in po_data_aggregation:
                 po_data_aggregation[po_str] = {'sqft_total': decimal.Decimal(0), 'amount_total': decimal.Decimal(0), 'items': set()}
             po_data_aggregation[po_str]['sqft_total'] += sqft_sum # type: ignore
             po_data_aggregation[po_str]['amount_total'] += amount_sum # type: ignore
             po_data_aggregation[po_str]['items'].add(item_str) # type: ignore

        if not po_data_aggregation:
            logging.warning(f"{prefix} No valid PO data found for PO count splitting. Returning empty dict.")
            return {}

        # Step 2: Get sorted list of unique POs
        sorted_pos = sorted(list(po_data_aggregation.keys()))

        # Step 3: Iterate through POs in conceptual groups of 8 for total calculation
        final_po_count_split_result: FinalFobResultType = {}
        # Calculate number of output chunks based on the total grouping size
        num_conceptual_chunks = (len(sorted_pos) + PO_GROUPING_FOR_TOTALS - 1) // PO_GROUPING_FOR_TOTALS

        logging.debug(f"{prefix} Pass 2: Creating {num_conceptual_chunks} output chunks based on conceptual PO groups of {PO_GROUPING_FOR_TOTALS}.")

        for i in range(num_conceptual_chunks):
            # Determine the POs belonging to this conceptual chunk (for totals)
            start_idx = i * PO_GROUPING_FOR_TOTALS
            end_idx = start_idx + PO_GROUPING_FOR_TOTALS
            conceptual_po_chunk = sorted_pos[start_idx:end_idx]

            # Calculate totals and collect items for THIS conceptual chunk
            chunk_sqft_total = decimal.Decimal(0)
            chunk_amount_total = decimal.Decimal(0)
            chunk_items = set()
            po_list_for_formatting = [] # Collect POs in this chunk for formatting

            for po_str in conceptual_po_chunk:
                po_agg_data = po_data_aggregation.get(po_str)
                if po_agg_data:
                    chunk_sqft_total += po_agg_data.get('sqft_total', decimal.Decimal(0)) # type: ignore
                    chunk_amount_total += po_agg_data.get('amount_total', decimal.Decimal(0)) # type: ignore
                    chunk_items.update(po_agg_data.get('items', set())) # type: ignore
                    po_list_for_formatting.append(po_str) # Add the PO itself to the list for formatting
                else:
                     logging.warning(f"{prefix} PO '{po_str}' not found in aggregation data during chunking.")

            # Sort items collected for this chunk
            sorted_chunk_items = sorted(list(chunk_items))

            # Step 4: Format the collected POs and Items using desired format (size 2)
            # --- Add Debugging --- 
            logging.debug(f"{prefix} Chunk {i+1}: Formatting POs. Input list ({len(po_list_for_formatting)} items): {po_list_for_formatting}")
            logging.debug(f"{prefix} Chunk {i+1}: PO Format Params: size={FOB_CHUNK_SIZE}, intra='{FOB_INTRA_CHUNK_SEPARATOR}', inter={repr(FOB_INTER_CHUNK_SEPARATOR)}")
            # --- End Debugging --- 
            formatted_po_chunk = format_chunks(po_list_for_formatting, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR)
            # --- Add Debugging --- 
            logging.debug(f"{prefix} Chunk {i+1}: Formatted POs Result: {repr(formatted_po_chunk)}")
            # --- End Debugging --- 

            # --- Add Debugging --- 
            logging.debug(f"{prefix} Chunk {i+1}: Formatting Items. Input list ({len(sorted_chunk_items)} items): {sorted_chunk_items}")
            logging.debug(f"{prefix} Chunk {i+1}: Item Format Params: size={FOB_CHUNK_SIZE}, intra='{FOB_INTRA_CHUNK_SEPARATOR}', inter={repr(FOB_INTER_CHUNK_SEPARATOR)}")
            # --- End Debugging --- 
            formatted_item_chunk = format_chunks(sorted_chunk_items, FOB_CHUNK_SIZE, FOB_INTRA_CHUNK_SEPARATOR, FOB_INTER_CHUNK_SEPARATOR)
            # --- Add Debugging --- 
            logging.debug(f"{prefix} Chunk {i+1}: Formatted Items Result: {repr(formatted_item_chunk)}")
            # --- End Debugging --- 

            # Create the result dictionary for this chunk index
            chunk_result: FobCompoundingResult = {
                'combined_po': formatted_po_chunk,
                'combined_item': formatted_item_chunk,
                'combined_description': '', # No descriptions in this path
                'total_sqft': chunk_sqft_total,    # Use CHUNK total (calculated based on group of 8)
                'total_amount': chunk_amount_total # Use CHUNK total (calculated based on group of 8)
            }
            chunk_index_str = str(i + 1)
            final_po_count_split_result[chunk_index_str] = chunk_result
            logging.debug(f"{prefix} Created output chunk {chunk_index_str}: {len(conceptual_po_chunk)} POs contributed totals, SQFT={chunk_sqft_total}, Amount={chunk_amount_total}")

        logging.info(f"{prefix} PO count split FOB Compounding complete ({len(final_po_count_split_result)} chunks created).")
        return final_po_count_split_result
        # --- End Path 2 (PO Count Split) --- #


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

# <<< MODIFIED FUNCTION SIGNATURE >>>
def run_invoice_automation(input_excel_override: Optional[str] = None, output_dir_override: Optional[str] = None):
    """Main function to find tables, extract, and process data for each.
       Uses input_excel_override if provided, otherwise falls back to cfg.INPUT_EXCEL_FILE.
       Saves output JSON to output_dir_override if provided, otherwise uses CWD.
    """
    logging.info("--- Starting Invoice Automation ---")
    handler = None
    actual_sheet_name = None
    input_filename = "Unknown"
    input_filepath = None
    output_dir = None

    # --- Determine Input Excel File ---
    if input_excel_override:
        input_filepath = input_excel_override
        logging.info(f"Using input Excel path from command line: {input_filepath}")
    else:
        try:
            # Fallback to config if no override provided
            input_filepath = cfg.INPUT_EXCEL_FILE
            logging.info(f"Using input Excel path from config.py: {input_filepath}")
        except AttributeError:
             logging.error("INPUT_EXCEL_FILE not found in config.py and no command-line override provided.")
             raise RuntimeError("Input Excel file path is missing.")
        except Exception as e:
            logging.error(f"Error accessing INPUT_EXCEL_FILE from config.py: {e}")
            raise RuntimeError(f"Could not determine input Excel file path: {e}")

    # Check if the determined filepath exists (relative to CWD or absolute)
    if not os.path.isfile(input_filepath):
         # Try resolving relative to the script's directory if not found in CWD
        script_dir = os.path.dirname(__file__)
        potential_path = os.path.join(script_dir, input_filepath)
        if os.path.isfile(potential_path):
            input_filepath = potential_path
            logging.info(f"Resolved relative input path to script directory: {input_filepath}")
        else:
            logging.error(f"Input Excel file not found at path: {input_filepath}")
            # Log CWD for debugging
            logging.error(f"Current working directory: {os.getcwd()}")
            if script_dir != os.getcwd():
                 logging.error(f"Script directory: {script_dir}")
            raise FileNotFoundError(f"Input Excel file not found: {input_filepath}")

    # Get just the filename for logging and output naming
    input_filename = os.path.basename(input_filepath)
    logging.info(f"Processing workbook: {input_filename}")
    # --- End Determine Input Excel File ---

    # --- Determine Output Directory ---
    if output_dir_override:
        output_dir = Path(output_dir_override).resolve()
        logging.info(f"Using output directory from command line: {output_dir}")
        # Ensure the directory exists
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
             logging.error(f"Could not create or access output directory '{output_dir}': {e}")
             raise RuntimeError(f"Invalid output directory specified: {output_dir}")
    else:
        # Default to current working directory if no override
        output_dir = Path(os.getcwd())
        logging.info(f"Using default output directory (CWD): {output_dir}")
    # --- End Determine Output Directory ---


    processed_tables: Dict[int, Dict[str, Any]] = {}
    all_tables_data: Dict[int, Dict[str, List[Any]]] = {}

    # Global dictionaries for initial aggregation results
    global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]] = {}
    global_custom_aggregation_results: Dict[Tuple[Any, Any, Optional[str], None], Dict[str, decimal.Decimal]] = {}
    # Global variable for the final FOB compounded result -> Type updated
    global_fob_compounded_result: Optional[FinalFobResultType] = None

    aggregation_mode_used = "standard" # Default, determines WHICH aggregation feeds FOB

    # --- Determine Initial Aggregation Strategy (based on the ACTUAL input filename) ---
    use_custom_aggregation_for_fob = False # Determines which map feeds FOB
    try:
        # Use the now determined input_filename
        logging.info(f"Checking workbook filename '{input_filename}' to determine PRIMARY aggregation mode for FOB compounding.")
        # Ensure CUSTOM_AGGREGATION_WORKBOOK_PREFIXES is defined in cfg
        custom_prefixes = getattr(cfg, 'CUSTOM_AGGREGATION_WORKBOOK_PREFIXES', [])
        if not isinstance(custom_prefixes, list):
            logging.warning("cfg.CUSTOM_AGGREGATION_WORKBOOK_PREFIXES is not a list. Using empty list.")
            custom_prefixes = []

        for prefix in custom_prefixes:
             if input_filename.startswith(prefix):
                use_custom_aggregation_for_fob = True # This workbook primarily uses custom for FOB
                aggregation_mode_used = "custom"
                logging.info(f"Workbook filename matches prefix '{prefix}'. Will use CUSTOM aggregation results for FOB compounding.")
                break
        if not use_custom_aggregation_for_fob:
             logging.info(f"Workbook filename does not match custom prefixes. Will use STANDARD aggregation results for FOB compounding.")
             aggregation_mode_used = "standard"
    except Exception as e:
        logging.error(f"Error during aggregation strategy determination for filename '{input_filename}': {e}")
        logging.warning("Defaulting to STANDARD aggregation for FOB compounding due to error.")
        aggregation_mode_used = "standard"
        use_custom_aggregation_for_fob = False
    # ---------------------------------------------

    try:
        # --- Steps 1-4: Load, Find Headers, Map Columns, Extract Data ---
        # <<< USE THE DETERMINED input_filepath >>>
        logging.info(f"Loading workbook from: {input_filepath}")
        handler = ExcelHandler(input_filepath)
        sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True)
        if sheet is None: raise RuntimeError(f"Failed to load sheet from '{input_filepath}'.")
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


        # --- Log Final FOB Compounded Result (INFO Level) - Simplified to expect split result --- #
        logging.info(f"--- Final FOB Compounded Result (Workbook: '{input_filename}', Based on '{aggregation_mode_used.upper()}' Input) ---")
        if global_fob_compounded_result is not None and isinstance(global_fob_compounded_result, dict) and "1" in global_fob_compounded_result:
            # Assume it's the BUFFALO split result ("1" and "2")
            logging.info(f"FOB result is split into BUFFALO (1) and NON-BUFFALO (2) groups.")
            for chunk_index, chunk_data in sorted(global_fob_compounded_result.items()):
                logging.info(f"--- FOB Group {chunk_index} --- ")
                if chunk_data and isinstance(chunk_data, dict):
                    po_string_value = chunk_data.get('combined_po', '<Not Found>')
                    item_string_value = chunk_data.get('combined_item', '<Not Found>')
                    desc_string_value = chunk_data.get('combined_description', '<Not Found>')
                    total_sqft_value = chunk_data.get('total_sqft', 'N/A')
                    total_amount_value = chunk_data.get('total_amount', 'N/A')

                    logging.info(f"  Combined POs:\n{po_string_value}")
                    logging.info(f"  Combined Items:\n{item_string_value}")
                    logging.info(f"  Combined Descriptions:\n{desc_string_value}")
                    logging.info(f"  Total SQFT: {total_sqft_value} (Type: {type(total_sqft_value)})")
                    logging.info(f"  Total Amount: {total_amount_value} (Type: {type(total_amount_value)})")
                else:
                    logging.info(f"  Group {chunk_index} data not found or invalid.")
            logging.info("-" * 30)

        elif global_fob_compounded_result is None:
            logging.error("FOB Compounding result is None or was not set.")
        else:
            # Handle unexpected type if necessary (e.g., empty dict if input was empty and aggregation failed)
             logging.warning(f"FOB Compounding result has unexpected structure/type: {type(global_fob_compounded_result)}")

        # --- End Final Logging ---


        # --- 8. Generate JSON Output ---
        logging.info("--- Preparing Data for JSON Output ---")
        try:
            # Create the structure to be converted to JSON
            # Use the helper function to ensure serializability
            final_json_structure = {
                 "metadata": {
                    "workbook_filename": input_filename, # Use the actual input filename
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

            # --- MODIFIED: Save JSON using output_dir and simplified filename ---
            input_stem = Path(input_filename).stem # Get filename without extension
            json_output_filename = f"{input_stem}.json" # Simplified filename
            output_json_path = output_dir / json_output_filename # Combine output dir and filename
            logging.info(f"Determined output JSON path: {output_json_path}")
            # --- END MODIFICATION ---
            try:
                with open(output_json_path, 'w', encoding='utf-8') as f_json:
                     f_json.write(json_output_string)
                logging.info(f"Successfully saved JSON output to '{output_json_path}'")
            except IOError as io_err:
                logging.error(f"Failed to write JSON output to file '{output_json_path}': {io_err}")
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
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description="Process an Excel invoice file to generate JSON data.")
    parser.add_argument(
        "--input-excel",
        type=str,
        default=None, # Default to None, indicating fallback to config.py
        help="Path to the input Excel file. Overrides the value in config.py if provided."
    )
    # --- ADDED: Output directory argument ---
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None, # Default to None, indicating use CWD
        help="Directory to save the output JSON file. Defaults to the current working directory."
    )
    # --- END ADD ---
    args = parser.parse_args()
    # --- End Argument Parsing ---

    # --- Run the main logic ---
    # Pass the parsed arguments to the main function
    run_invoice_automation(
        input_excel_override=args.input_excel,
        output_dir_override=args.output_dir # Pass the output dir argument
    )
    # --- End Run Logic ---

# --- END OF FULL FILE: main.py ---