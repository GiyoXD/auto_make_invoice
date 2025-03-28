import logging
from typing import Dict, List, Any, Optional
import decimal # Use Decimal for precise calculations
import re

# Import config values (consider passing as arguments)
# DISTRIBUTION_BASIS_COLUMN is needed here, COLUMNS_TO_DISTRIBUTE is used in main.py loop
from config import DISTRIBUTION_BASIS_COLUMN

# Set precision for Decimal calculations
decimal.getcontext().prec = 28 # Default precision, adjust if needed
# Define precision specifically for CBM results (e.g., 4 decimal places)
CBM_DECIMAL_PLACES = decimal.Decimal('0.0001')

class ProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass

def _convert_to_decimal(value: Any, context: str = "") -> Optional[decimal.Decimal]:
    """Safely convert a value to Decimal, logging errors."""
    # --- Added handling if input is already Decimal ---
    if isinstance(value, decimal.Decimal):
        return value
    # -------------------------------------------------
    if value is None:
        return None
    value_str = str(value).strip()
    if not value_str:
        return None
    try:
        # Handle potential strings with commas, etc., if necessary
        # cleaned_value = value_str.replace(',', '') # Example
        return decimal.Decimal(value_str)
    except (decimal.InvalidOperation, TypeError, ValueError) as e:
        logging.warning(f"Could not convert '{value}' to Decimal {context}: {e}")
        return None

# --- NEW FUNCTION TO CALCULATE CBM FROM STRING ---
def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.

    Args:
        cbm_value: The value from the CBM cell (can be string, number, None).
        row_index: The 0-based index of the row for logging purposes.

    Returns:
        The calculated CBM as a Decimal, or None if parsing fails or input is invalid.
    """
    if cbm_value is None:
        return None

    # If it's already a number (potentially from a previous run or direct input), just convert to Decimal and quantize
    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        calculated = _convert_to_decimal(cbm_value, f"for pre-numeric CBM at row {row_index + 1}")
        # Quantize even pre-numeric values for consistency
        return calculated.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP) if calculated is not None else None


    if not isinstance(cbm_value, str):
        logging.warning(f"Unexpected type '{type(cbm_value).__name__}' for CBM value '{cbm_value}' at row {row_index + 1}. Cannot calculate.")
        return None

    cbm_str = cbm_value.strip()
    if not cbm_str:
        return None

    # Try splitting by '*' first
    parts = cbm_str.split('*')

    # If not 3 parts, try splitting by 'x' or 'X' (case-insensitive)
    if len(parts) != 3:
        if '*' not in cbm_str and ('x' in cbm_str.lower()):
             parts = re.split(r'[xX]', cbm_str) # Split by 'x' or 'X'

    # Check if we have exactly 3 parts after trying separators
    if len(parts) != 3:
        logging.warning(f"Invalid CBM format: '{cbm_str}' at row {row_index + 1}. Expected 3 parts separated by '*' or 'x'. Found {len(parts)} parts.")
        return None

    try:
        # Convert each part to Decimal
        dim1 = _convert_to_decimal(parts[0], f"for CBM part 1 at row {row_index + 1}")
        dim2 = _convert_to_decimal(parts[1], f"for CBM part 2 at row {row_index + 1}")
        dim3 = _convert_to_decimal(parts[2], f"for CBM part 3 at row {row_index + 1}")

        # Check if any conversion failed
        if dim1 is None or dim2 is None or dim3 is None:
            # _convert_to_decimal already logged the specific conversion error
            logging.warning(f"Failed to convert one or more dimensions for CBM '{cbm_str}' at row {row_index + 1}.")
            return None

        # Check for non-positive dimensions if necessary (optional)
        if dim1 <= 0 or dim2 <= 0 or dim3 <= 0:
             logging.warning(f"Non-positive dimension found in CBM '{cbm_str}' at row {row_index + 1}. Result will be non-positive.")

        # Calculate volume and quantize to the defined CBM precision
        volume = (dim1 * dim2 * dim3).quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
        logging.debug(f"Calculated CBM at row {row_index + 1}: {volume} from '{cbm_str}'")
        return volume

    except Exception as e:
        # Catch any unexpected error during calculation
        logging.error(f"Unexpected error calculating CBM from '{cbm_str}' at row {row_index + 1}: {e}", exc_info=True)
        return None

# --- NEW FUNCTION TO PROCESS THE ENTIRE CBM COLUMN ---
def process_cbm_column(raw_data: Dict[str, List[Any]]) -> Dict[str, List[Any]]:
    """
    Iterates through the 'cbm' list in raw_data, calculates numeric CBM values
    from strings (L*W*H or LxWxH format), and updates the list in place.

    Args:
        raw_data: The dictionary containing lists of extracted data for a single table.

    Returns:
        The modified raw_data dictionary with the 'cbm' list processed (contains Decimals or Nones).
    """
    cbm_key = 'cbm' # Canonical name defined in config.py

    if cbm_key not in raw_data:
        logging.debug(f"No '{cbm_key}' column found in this table's extracted data. Skipping CBM calculation.")
        return raw_data # No CBM column to process

    original_cbm_list = raw_data.get(cbm_key, []) # Get list safely
    if not original_cbm_list:
        logging.debug(f"'{cbm_key}' column is present but empty. Skipping CBM calculation.")
        return raw_data # Empty list, nothing to process

    logging.info(f"Processing '{cbm_key}' column for volume calculations...")
    calculated_cbm_list = []

    # Process each value in the original list
    for i, value in enumerate(original_cbm_list):
        calculated_value = _calculate_single_cbm(value, i) # Calculate volume
        calculated_cbm_list.append(calculated_value) # Add Decimal or None

    # Replace the original list in the dictionary with the newly calculated list
    raw_data[cbm_key] = calculated_cbm_list
    logging.info(f"Finished processing '{cbm_key}' column. It now contains calculated values.")
    return raw_data
# --- END OF NEW CBM FUNCTIONS ---


# --- UPDATED distribute_values function ---
def distribute_values(
    raw_data: Dict[str, List[Any]],
    columns_to_distribute: List[str],
    basis_column: str
) -> Dict[str, List[Any]]:
    """
    Distributes values in specified columns based on proportions in the basis column.
    Operates on the input raw_data (which might have pre-calculated CBM).
    Handles pre-calculated CBM decimals correctly.
    """
    if not raw_data:
        logging.warning("Received empty raw_data for distribution.")
        return {}

    processed_data = raw_data # Modify in place

    # --- Input Validation ---
    if basis_column not in processed_data:
        logging.error(f"Basis column '{basis_column}' not found in data for distribution. Cannot distribute.")
        return processed_data

    valid_columns_to_distribute = []
    for col in columns_to_distribute:
        if col not in processed_data:
             logging.warning(f"Column '{col}' specified for distribution but not found in this table's data. Skipping.")
        else:
            valid_columns_to_distribute.append(col)

    if not valid_columns_to_distribute:
         logging.warning("No valid columns found to perform distribution on in this table.")
         return processed_data

    basis_values_list = processed_data.get(basis_column, [])
    num_rows = len(basis_values_list)
    if num_rows == 0:
        logging.info("No data rows found in the basis column. Skipping distribution.")
        return processed_data

    logging.info(f"Starting value distribution for columns: {valid_columns_to_distribute} based on '{basis_column}'.")

    basis_values_dec: List[Optional[decimal.Decimal]] = [
        _convert_to_decimal(val, f"in basis column '{basis_column}' at row {i+1}")
        for i, val in enumerate(basis_values_list)
    ]

    # --- Process each column ---
    for col_name in valid_columns_to_distribute:
        logging.debug(f"Processing column for distribution: {col_name}")
        original_col_values = processed_data.get(col_name, [])
        if len(original_col_values) != num_rows:
             logging.error(f"Row count mismatch: basis '{basis_column}' ({num_rows}) vs '{col_name}' ({len(original_col_values)}). Skipping distribution for '{col_name}'.")
             continue

        processed_col_values: List[Any] = [None] * num_rows

        i = 0
        while i < num_rows:
            current_val_dec = _convert_to_decimal(original_col_values[i], f"in column '{col_name}' at row {i+1}")

            if current_val_dec is not None and current_val_dec != 0: # Found a non-zero value to distribute
                processed_col_values[i] = current_val_dec # Store original value (as Decimal)

                # Look ahead for distribution block
                j = i + 1
                distribution_rows_indices = []
                while j < num_rows:
                     # Check original next value
                     original_next_val_dec = _convert_to_decimal(original_col_values[j], f"in column '{col_name}' lookahead at row {j+1}")
                     # Stop lookahead if non-empty/non-zero
                     if original_next_val_dec is not None and original_next_val_dec != 0:
                          break

                     # Check basis value for this potential row
                     if basis_values_dec[j] is not None:
                          distribution_rows_indices.append(j)
                     else:
                          logging.warning(f"  Skipping row {j+1} in block for '{col_name}' (from row {i+1}) due to missing basis. Assigning 0.")
                          processed_col_values[j] = decimal.Decimal(0) # Assign 0 immediately
                     j += 1

                if distribution_rows_indices:
                    # Found a block
                    block_indices = [i] + distribution_rows_indices
                    logging.debug(f"  Found distribution block in '{col_name}': Rows indices {block_indices} (1-based: {[x+1 for x in block_indices]})")

                    # Sum basis for the block
                    total_basis_in_block = decimal.Decimal(0)
                    valid_basis_found = False
                    for k in block_indices:
                        basis_val = basis_values_dec[k]
                        if basis_val is not None and basis_val > 0:
                            total_basis_in_block += basis_val
                            valid_basis_found = True
                        elif basis_val is not None:
                             logging.warning(f"  Basis value is zero/negative ({basis_val}) in row {k+1} for block from row {i+1}. Excluding from total.")

                    # Distribute if possible
                    if total_basis_in_block > 0 and valid_basis_found:
                         distributed_sum_check = decimal.Decimal(0)
                         # --- USE CORRECT PRECISION ---
                         dist_precision = CBM_DECIMAL_PLACES if col_name == 'cbm' else decimal.Decimal('0.0001') # Default precision
                         # -----------------------------
                         logging.debug(f"    Total basis: {total_basis_in_block}. Distributing {current_val_dec} with precision {dist_precision}")

                         for k in block_indices:
                             basis_val = basis_values_dec[k]
                             if basis_val is not None and basis_val > 0:
                                 proportion = basis_val / total_basis_in_block
                                 distributed_value = (current_val_dec * proportion).quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                                 processed_col_values[k] = distributed_value
                                 distributed_sum_check += distributed_value
                                 logging.debug(f"    Row {k+1}: Basis={basis_val}, Prop={proportion:.6f}, Dist Val={distributed_value}")
                             elif processed_col_values[k] is None: # Handle rows skipped earlier or with 0 basis
                                 processed_col_values[k] = decimal.Decimal(0)
                                 log_reason = "missing basis" if basis_val is None else f"zero/negative basis ({basis_val})"
                                 logging.warning(f"    Row {k+1}: Assigning 0 to distributed value for '{col_name}' due to {log_reason}.")


                         # Distribution check
                         tolerance = dist_precision / 2
                         if not abs(distributed_sum_check - current_val_dec) <= tolerance:
                              logging.warning(f"  Distribution check failed for block from row {i+1}, column '{col_name}'. Original: {current_val_dec}, Sum: {distributed_sum_check}, Diff: {distributed_sum_check - current_val_dec}")
                         else:
                              logging.debug(f"  Distribution check passed for block from row {i+1}. Original: {current_val_dec}, Sum: {distributed_sum_check}")

                    else:
                        # Cannot distribute
                        logging.warning(f"  Cannot distribute value '{current_val_dec}' from row {i+1} in '{col_name}'. Total positive basis is zero or invalid. Keeping original at row {i+1}, setting others in block to 0.")
                        for k in distribution_rows_indices:
                            if processed_col_values[k] is None: # Ensure rows skipped due to missing basis are 0
                                processed_col_values[k] = decimal.Decimal(0)

                    i = j # Move index past the processed block
                else:
                    # No empty/zero cells followed this value
                    i += 1
            else:
                # Current original cell value is empty/zero/invalid.
                if processed_col_values[i] is None: # Check if not filled by end of a previous block
                     processed_col_values[i] = decimal.Decimal(0) # Set to 0 if empty/invalid
                i += 1

        # Update the dictionary with the processed list
        processed_data[col_name] = processed_col_values
        logging.debug(f"Finished processing column for distribution: {col_name}")

    logging.info("Value distribution processing complete for this table.")
    return processed_data
# --- END OF UPDATED distribute_values function ---

# --- END OF FILE: data_processor.py ---