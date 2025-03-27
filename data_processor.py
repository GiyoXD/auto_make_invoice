# --- START OF FILE data_processor.py ---

import logging
from typing import Dict, List, Any, Optional
import decimal # Use Decimal for precise calculations, especially with currency/weights

# Import config values (consider passing as arguments)
from config import COLUMNS_TO_DISTRIBUTE, DISTRIBUTION_BASIS_COLUMN

# Set precision for Decimal calculations
decimal.getcontext().prec = 28 # Default precision, adjust if needed

class ProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass


def _convert_to_decimal(value: Any, context: str = "") -> Optional[decimal.Decimal]:
    """Safely convert a value to Decimal, logging errors."""
    if value is None or str(value).strip() == "":
        return None
    try:
        # Handle potential strings with commas, etc., if necessary
        # cleaned_value = str(value).replace(',', '') # Example
        return decimal.Decimal(value)
    except (decimal.InvalidOperation, TypeError, ValueError) as e:
        logging.warning(f"Could not convert '{value}' to Decimal {context}: {e}")
        return None


def distribute_values(
    raw_data: Dict[str, List[Any]],
    columns_to_distribute: List[str],
    basis_column: str
) -> Dict[str, List[Any]]:
    """
    Distributes values in specified columns based on proportions in the basis column.

    Operates on a copy of the raw_data to avoid modifying the original dict directly.
    Handles sequences where a value is followed by None/empty cells.

    Args:
        raw_data: Dictionary of lists {'col_name': [val1, val2, None, val4,...]}
        columns_to_distribute: List of column names (keys in raw_data) to process.
        basis_column: The name of the column (key in raw_data) containing the
                      proportional values (e.g., 'pcs').

    Returns:
        A new dictionary with the distributed values filled in.

    Raises:
        ProcessingError: If essential columns are missing or data issues occur.
    """
    if not raw_data:
        logging.warning("Received empty raw_data for distribution.")
        return {}

    # Make a deep copy if you need to preserve the original raw_data elsewhere
    # For now, we'll modify a shallow copy (modifies lists in place)
    processed_data = raw_data.copy()

    # --- Input Validation ---
    if basis_column not in processed_data:
        raise ProcessingError(f"Basis column '{basis_column}' not found in data.")
    for col in columns_to_distribute:
        if col not in processed_data:
             raise ProcessingError(f"Column to distribute '{col}' not found in data.")

    num_rows = len(processed_data[basis_column])
    if num_rows == 0:
        logging.info("No data rows to process for distribution.")
        return processed_data # Return the empty structure

    logging.info(f"Starting value distribution for columns: {columns_to_distribute} based on '{basis_column}'.")

    # Convert basis column to Decimals first for efficiency
    basis_values_dec: List[Optional[decimal.Decimal]] = [
        _convert_to_decimal(val, f"in basis column '{basis_column}' at row {i+1}")
        for i, val in enumerate(processed_data[basis_column])
    ]

    # --- Process each column that needs distribution ---
    for col_name in columns_to_distribute:
        logging.debug(f"Processing column: {col_name}")
        original_col_values = processed_data[col_name]
        processed_col_values: List[Any] = [None] * num_rows # Initialize with Nones

        i = 0
        while i < num_rows:
            current_val = original_col_values[i]
            current_val_dec = _convert_to_decimal(current_val, f"in column '{col_name}' at row {i+1}")

            if current_val_dec is not None:
                # Found a value. Check for subsequent empty cells to distribute over.
                processed_col_values[i] = current_val_dec # Store the original value (as Decimal)

                # Look ahead for the distribution block
                j = i + 1
                distribution_rows_indices = []
                while j < num_rows and (original_col_values[j] is None or str(original_col_values[j]).strip() == ""):
                    distribution_rows_indices.append(j)
                    j += 1

                if distribution_rows_indices:
                    # We have a block to distribute over (rows i to j-1)
                    block_indices = [i] + distribution_rows_indices
                    logging.debug(f"  Found block for distribution: Rows {min(block_indices)+1} to {max(block_indices)+1}")

                    # Sum the basis values (e.g., pcs) for this block
                    total_basis_in_block = decimal.Decimal(0)
                    valid_basis_found = False
                    for k in block_indices:
                        basis_val = basis_values_dec[k]
                        if basis_val is not None and basis_val > 0: # Ensure basis is positive
                            total_basis_in_block += basis_val
                            valid_basis_found = True
                        elif basis_val is not None:
                             logging.warning(f"  Basis value is zero or negative ({basis_val}) in row {k+1} for block starting at row {i+1}. Skipping this row in proportion calc.")


                    if total_basis_in_block > 0 and valid_basis_found:
                         # Distribute current_val_dec proportionally
                         distributed_sum_check = decimal.Decimal(0)
                         for k in block_indices:
                             basis_val = basis_values_dec[k]
                             if basis_val is not None and basis_val > 0:
                                 proportion = basis_val / total_basis_in_block
                                 distributed_value = (current_val_dec * proportion).quantize(decimal.Decimal('0.0001'), rounding=decimal.ROUND_HALF_UP) # Adjust precision/rounding as needed
                                 processed_col_values[k] = distributed_value
                                 distributed_sum_check += distributed_value
                                 logging.debug(f"    Row {k+1}: Basis={basis_val}, Prop={proportion:.4f}, Dist={distributed_value}")
                             else:
                                 # Handle rows with invalid/missing basis within the block (assign 0 or None?)
                                 processed_col_values[k] = decimal.Decimal(0) # Or None
                                 logging.warning(f"    Row {k+1}: Invalid or missing basis value. Assigning 0 to distributed value for '{col_name}'.")


                         # Optional: Check if distributed sum roughly matches original
                         if not abs(distributed_sum_check - current_val_dec) < decimal.Decimal('0.01'): # Allow small tolerance
                              logging.warning(f"  Distribution check failed for block starting row {i+1}, column '{col_name}'. "
                                              f"Original: {current_val_dec}, Distributed Sum: {distributed_sum_check}")

                    else:
                        # Cannot distribute (total basis is zero or invalid basis values)
                        logging.warning(f"  Cannot distribute value '{current_val_dec}' from row {i+1} in column '{col_name}'. "
                                        f"Total basis for block (Rows {min(block_indices)+1}-{max(block_indices)+1}) is zero or invalid.")
                        # Keep original value at index i, leave others as None or fill with 0?
                        for k in distribution_rows_indices:
                            processed_col_values[k] = None # Or decimal.Decimal(0)

                    # Move index past the processed block
                    i = j
                else:
                    # No empty cells followed, just move to the next row
                    i += 1
            else:
                # Current cell is empty, leave it as None in processed list (unless it was part of a block handled above)
                if processed_col_values[i] is None: # Avoid overwriting if filled by a distribution block
                     processed_col_values[i] = None
                i += 1

        # Update the dictionary with the processed list (containing Decimals or Nones)
        processed_data[col_name] = processed_col_values
        logging.debug(f"Finished processing column: {col_name}")


    # Optional: Convert Decimals back to floats or strings if needed for output
    # for col_name in columns_to_distribute:
    #    processed_data[col_name] = [float(v) if isinstance(v, decimal.Decimal) else v for v in processed_data[col_name]]

    logging.info("Value distribution processing complete.")
    return processed_data

# --- END OF FILE data_processor.py ---