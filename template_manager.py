# template_manager.py

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell, MergedCell # <-- Import MergedCell
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
# Import range utils for merge checking
from openpyxl.utils.cell import range_boundaries

import logging
import os
import pprint
import shutil

# Import necessary types from 'typing'
from typing import Dict, List, Any, Union, Optional, Tuple # Added List, Tuple

# --- Constants ---
DEFAULT_ANCHOR_TEXT = "Mark & Nº"
DEFAULT_CELLS_BELOW = 5
DEFAULT_CELLS_RIGHT = 10
MAX_LOG_DICT_LEN = 3500 # Increased slightly for list logging
COPY_SUFFIX = "_copy"
ANCHOR_KEY = 'anchor'

# --- Module Level Cache ---
# Stores extracted template data.
# Keys 'bottom_N' map to Optional[Dict[str, Any]]
# Keys 'right_N' map to Optional[List[Optional[Dict[str, Any]]]]
# Key 'anchor' maps to Optional[Dict[str, Any]] (may include 'merge_info', '_anchor_merge_rows')
# Adjusted type hint for the cache value itself to handle mixed types (Dict or List)
_template_cache: Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]] = {}


# --- Helper Functions ---

# _find_anchor_cell - Merge Aware (Finds top-left of merge)
def _find_anchor_cell(worksheet: Worksheet, anchor_text: str) -> Optional[Cell]:
    """
    Finds the FIRST cell containing the anchor text during a row-by-row scan.
    If the cell found is part of a merged range, it attempts to return the
    TOP-LEFT cell of that merge range.

    Args:
        worksheet: The openpyxl worksheet object.
        anchor_text: The text to search for (case-sensitive).

    Returns:
        The openpyxl Cell object representing the top-left corner of the merge
        (or the cell itself if not merged) containing the text, otherwise None.
    """
    prefix = "[TemplateManager._find_anchor_cell (Merge Aware)]"
    logging.debug(f"{prefix} Searching for FIRST cell containing anchor text '{anchor_text}' in sheet '{worksheet.title}'...")

    # Pre-build map of coordinate -> top-left coordinate of its merge range
    merge_map: Dict[str, str] = {}
    for mc_range in worksheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = mc_range.bounds
        top_left_coord = f"{get_column_letter(min_col)}{min_row}"
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                coord = f"{get_column_letter(c)}{r}"
                merge_map[coord] = top_left_coord

    for row_idx, row in enumerate(worksheet.iter_rows(), start=1):
        for col_idx, cell in enumerate(row, start=1):
            current_coord = cell.coordinate # Use cell's coordinate directly

            # Check value FIRST before checking merge status
            cell_value = cell.value
            if cell_value is None:
                # If value is None, we only care if it's the *top-left* of a merged cell,
                # as that's where the value *might* be stored according to spec,
                # although often it's empty too. A non-top-left merged cell
                # *should* have a None value, so we skip those unless they are
                # the designated top-left.
                is_top_left_of_merge = False
                if current_coord in merge_map and merge_map[current_coord] == current_coord:
                    is_top_left_of_merge = True
                # If it's None AND not the top-left of a merge, skip it.
                if not is_top_left_of_merge:
                    continue
                # If it *is* the top-left of a merge but the value is None,
                # we still proceed to the try/except block below, as we *might*
                # be searching for "None" or an empty string representation.

            try:
                 # Check if anchor text is in the cell value's string representation
                 if anchor_text in str(cell_value):
                     logging.info(f"{prefix} Found anchor text '{anchor_text}' potentially in cell {current_coord}.")
                     # Now, check if this cell is part of a merge range using our map
                     if current_coord in merge_map:
                         top_left_coord = merge_map[current_coord]
                         logging.info(f"{prefix} Cell {current_coord} is part of merged range starting at {top_left_coord}. Returning top-left cell object.")
                         # Get the Cell object for the top-left coordinate
                         if current_coord == top_left_coord:
                             return cell # We found the text in the top-left cell directly
                         else:
                             # We found the text in a cell that is *not* the top-left of its merge.
                             # The actual value resides in the top-left. Return the top-left Cell object.
                             tl_col_str, tl_row_str = coordinate_from_string(top_left_coord)
                             tl_col = column_index_from_string(tl_col_str)
                             tl_row = int(tl_row_str)
                             try:
                                 # Return the actual top-left cell object
                                 return worksheet.cell(row=tl_row, column=tl_col)
                             except IndexError:
                                 logging.warning(f"{prefix} Calculated top-left {top_left_coord} seems out of bounds. Returning originally found cell {current_coord}.")
                                 return cell # Fallback to the cell where text was found
                     else:
                         # Not part of any known merge range
                         logging.info(f"{prefix} Cell {current_coord} contains text and is not merged. Returning this cell.")
                         return cell
            except TypeError:
                 logging.debug(f"{prefix} Could not compare anchor text with cell {current_coord} value '{cell_value}' (Type: {type(cell_value)}). Skipping.")
                 continue
            except Exception as e:
                 logging.warning(f"{prefix} Unexpected error checking cell {current_coord} value '{cell_value}': {e}")
                 continue

    logging.warning(f"{prefix} Anchor text '{anchor_text}' not found in any cell (or merge top-left) in sheet '{worksheet.title}'.")
    return None


# _extract_relative_cells - Extracts Anchor (inc. Merge), Bottom (Single), Right (Full Height Lists)
def _extract_relative_cells(
    worksheet: Worksheet,
    anchor_cell: Cell, # Expected to be top-left if merged
    num_below: int,
    num_right: int
) -> Optional[Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]]]:
    """
    Extracts properties from the anchor cell (inc. merge info) and cells relative.
    'bottom_N': Extracts single cell N rows below anchor.
    'right_N': Extracts a LIST of cells N columns right of anchor, spanning the
               same number of rows as the anchor's merge height (or 1 if not merged).

    Args:
        worksheet: The openpyxl worksheet object.
        anchor_cell: The top-left Cell object of the anchor.
        num_below: Number of cells below anchor to extract.
        num_right: Number of columns right of anchor to extract vertically.

    Returns:
        A dictionary where 'anchor' and 'bottom_N' map to single cell data dicts
        (or None/error), and 'right_N' maps to a LIST of cell data dicts
        (or None/error for each cell in the list). Returns None on major error.
    """
    prefix = "[TemplateManager._extract_relative_cells (Full Height Right)]"
    if not anchor_cell:
        logging.error(f"{prefix} Invalid anchor_cell provided.")
        return None

    # Type hint for the main dictionary value reflects mixed types
    relative_data: Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]] = {}
    anchor_row = anchor_cell.row
    anchor_col = anchor_cell.column
    anchor_coord = anchor_cell.coordinate

    logging.info(f"{prefix} Extracting properties starting at anchor {anchor_coord} (R={anchor_row}, C={anchor_col}).")
    logging.info(f"{prefix}   Below: {num_below}, Right (Full Height): {num_right}")

    # Get default dimensions
    default_row_height = worksheet.sheet_format.defaultRowHeight if worksheet.sheet_format and worksheet.sheet_format.defaultRowHeight else 15.0
    default_col_width = worksheet.sheet_format.defaultColWidth if worksheet.sheet_format and worksheet.sheet_format.defaultColWidth else 8.43

    # --- Helper: Extract single cell properties ---
    def _get_cell_props(cell: Cell, target_r: int, target_c: int, is_anchor: bool = False) -> Dict[str, Any]:
        """Extracts value, dimensions, border, and merge info (if anchor)."""
        props: Dict[str, Any] = {}
        try:
            # 1. Value
            # For MergedCells (not top-left), value will be None. We still store it.
            # The actual value/style comes from the top-left cell of the merge.
            props['value'] = cell.value

            # 2. Height (Row property) - Always read from the target row dimension
            rd = worksheet.row_dimensions.get(target_r)
            props['height'] = rd.height if rd and rd.height is not None else default_row_height

            # 3. Width (Column property) - Always read from the target column dimension
            col_letter = get_column_letter(target_c)
            cd = worksheet.column_dimensions.get(col_letter)
            props['width'] = cd.width if cd and cd.width is not None else default_col_width

            # 4. Border - Read from the specific cell's style
            border_styles: Dict[str, Optional[str]] = {'left': None, 'right': None, 'top': None, 'bottom': None}
            # Use cell.border directly, works even for MergedCell (inherits style)
            if cell.has_style and cell.border:
                border_styles['left'] = cell.border.left.style if cell.border.left else None
                border_styles['right'] = cell.border.right.style if cell.border.right else None
                border_styles['top'] = cell.border.top.style if cell.border.top else None
                border_styles['bottom'] = cell.border.bottom.style if cell.border.bottom else None
            props['border'] = border_styles

            # 5. Merge Info (Only extracted for the ANCHOR cell based on its top-left coords)
            anchor_merge_rows = 1 # Default row span if not anchor or not merged
            if is_anchor:
                merge_info = None
                # Check worksheet's merged ranges against the ANCHOR's coordinates
                for mc_range in worksheet.merged_cells.ranges:
                    # Check if the anchor cell is the top-left corner of this merge range
                    if mc_range.min_row == target_r and mc_range.min_col == target_c:
                        rows_merged = mc_range.max_row - mc_range.min_row + 1
                        cols_merged = mc_range.max_col - mc_range.min_col + 1
                        # Only store merge info if it actually spans more than one cell
                        if rows_merged > 1 or cols_merged > 1:
                             merge_info = {'rows': rows_merged, 'cols': cols_merged}
                             anchor_merge_rows = rows_merged # Store actual merge height
                             logging.debug(f"{prefix}   -> Anchor cell {cell.coordinate} is top-left of merge: {rows_merged}r x {cols_merged}c.")
                        break # Found the relevant merge range for the anchor
                if merge_info:
                     props['merge_info'] = merge_info
                # Store merge height directly on anchor props for easier access later
                # Use internal key to avoid conflict if user names a field '_anchor_merge_rows'
                props['_anchor_merge_rows'] = anchor_merge_rows

            # Logging
            border_log_repr = f"L={border_styles['left']}, R={border_styles['right']}, T={border_styles['top']}, B={border_styles['bottom']}"
            merge_log_repr = f", Merged={props['merge_info']}" if 'merge_info' in props else ""
            val_repr = str(props.get('value', ''))[:30] # Limit value length in log
            cell_type_log = f"(Type: {type(cell).__name__})" # Log if it's Cell or MergedCell
            logging.debug(
                f"{prefix}     -> Extracted (R={target_r}, C={target_c}, Cell: {cell.coordinate} {cell_type_log}): "
                f"Val='{val_repr}...', H={props['height']:.1f}, W={props['width']:.1f}, "
                f"Border={{{border_log_repr}}}{merge_log_repr}"
            )
            return props
        except Exception as e:
            logging.error(f"{prefix} Error reading properties for cell at R={target_r}, C={target_c}: {e}", exc_info=False)
            return {'error': f"Error reading cell properties: {e}"}

    # --- Extract Anchor Cell Properties ---
    logging.debug(f"{prefix} Extracting ANCHOR cell ({ANCHOR_KEY}) at {anchor_coord}...")
    anchor_props = _get_cell_props(anchor_cell, anchor_row, anchor_col, is_anchor=True)
    relative_data[ANCHOR_KEY] = anchor_props
    # Get anchor merge height (default 1 if not merged or error getting props)
    anchor_merge_height = 1
    if isinstance(anchor_props, dict): # Check it's a dict before accessing keys
        # Use the internal key we stored
        anchor_merge_height = anchor_props.get('_anchor_merge_rows', 1)
    logging.debug(f"{prefix}   Anchor merge height determined: {anchor_merge_height} rows.")


    # --- Extract Cells Below (Single Cells) ---
    logging.debug(f"{prefix} Extracting cells BELOW anchor (single cells)...")
    for i in range(1, num_below + 1):
        target_row = anchor_row + i
        target_col = anchor_col
        key_name = f'bottom_{i}'
        if target_row > worksheet.max_row or target_col > worksheet.max_column:
            logging.warning(f"{prefix} Target {key_name} (R={target_row}, C={target_col}) out of bounds. Storing None.")
            relative_data[key_name] = None
            continue
        try:
             # Get the cell object first
             cell_to_extract = worksheet.cell(row=target_row, column=target_col)
             # Extract its properties (is_anchor=False)
             relative_data[key_name] = _get_cell_props(cell_to_extract, target_row, target_col)
        except IndexError:
             logging.warning(f"{prefix} Could not access cell for {key_name} at (R={target_row}, C={target_col}). Storing error.")
             relative_data[key_name] = {'error': f"IndexError accessing cell at R={target_row}, C={target_col}"}
        except Exception as exc:
             logging.error(f"{prefix} Error getting props for {key_name} at (R={target_row}, C={target_col}): {exc}")
             relative_data[key_name] = {'error': f"Error getting props: {exc}"}


    # --- Extract Cells To The RIGHT (Full Height) ---
    logging.debug(f"{prefix} Extracting cells to the RIGHT of anchor (Full Height: {anchor_merge_height} rows each)...")
    for i in range(1, num_right + 1):
        target_base_col = anchor_col + i
        key_name = f'right_{i}'
        # List to hold properties for cells in this vertical slice
        cells_in_column: List[Optional[Dict[str, Any]]] = []

        logging.debug(f"{prefix}   Processing '{key_name}' (Column Index: {target_base_col})...")

        # Loop vertically based on anchor's merge height
        for j in range(anchor_merge_height):
            target_row = anchor_row + j
            target_col = target_base_col # Column is fixed for this 'right_i'

            # Check bounds for this specific cell in the vertical slice
            if target_row > worksheet.max_row or target_col > worksheet.max_column:
                logging.warning(f"{prefix}     Cell at (R={target_row}, C={target_col}) for '{key_name}' (row offset {j}) out of bounds. Storing None in list.")
                cells_in_column.append(None)
                continue

            # Try to get cell and extract properties
            try:
                # Get the cell object first - this might be a Cell or a MergedCell
                cell_to_extract = worksheet.cell(row=target_row, column=target_col)

                # Log if this cell is part of another merge (but not top-left)
                # This check is mostly informational during extraction.
                if isinstance(cell_to_extract, MergedCell):
                     logging.debug(f"{prefix}       Cell at (R={target_row}, C={target_col}) is a MergedCell (part of a merged range, not top-left). Extracting its properties (value will be None, style inherited).")
                # The _get_cell_props function handles both Cell and MergedCell types correctly for value/style extraction.

                # Extract properties from the cell object (is_anchor=False)
                props = _get_cell_props(cell_to_extract, target_row, target_col)
                cells_in_column.append(props)

            except IndexError:
                logging.warning(f"{prefix}     IndexError accessing cell for '{key_name}' at (R={target_row}, C={target_col}). Storing error in list.")
                cells_in_column.append({'error': f"IndexError accessing cell at R={target_row}, C={target_col}"})
            except Exception as cell_exc:
                 logging.error(f"{prefix}     Unexpected error processing cell for '{key_name}' at (R={target_row}, C={target_col}): {cell_exc}. Storing error in list.")
                 cells_in_column.append({'error': f"Unexpected error processing cell: {cell_exc}"})

        # Store the list of cell data for this 'right_i' key
        relative_data[key_name] = cells_in_column
        # Log the structure of the list briefly for debugging
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
            list_repr_items = []
            for item_idx, d in enumerate(cells_in_column):
                # Limit the number of items shown for long lists
                if item_idx >= 3 and len(cells_in_column) > 4 :
                    if item_idx == 3: list_repr_items.append("...")
                    if item_idx < len(cells_in_column) -1: continue
                # Represent the item concisely
                if isinstance(d, dict): list_repr_items.append(f"Dict({len(d)})" if 'error' not in d else "ErrDict")
                else: list_repr_items.append(str(d))
            list_repr = "[" + ", ".join(list_repr_items) + "]"
            logging.debug(f"{prefix}   Stored '{key_name}' as List (Length={len(cells_in_column)}): {list_repr}")


    logging.info(f"{prefix} Finished extracting properties. Found data for {len(relative_data)} positions ('right_N' keys contain lists).")
    return relative_data


# --- Public Interface Functions ---

# load_template - Uses revised extraction (full height right)
def load_template(
    workbook_path: str,
    sheet_identifier: Union[str, int],
    template_name: str,
    anchor_text: str = DEFAULT_ANCHOR_TEXT,
    num_below: int = DEFAULT_CELLS_BELOW,
    num_right: int = DEFAULT_CELLS_RIGHT,
    force_reload: bool = False,
    create_copy: bool = False
) -> bool:
    """
    Loads template, finds anchor (merge-aware), extracts properties.
    'right_N' extraction captures cells vertically spanning the anchor's merge height.

    Args:
        workbook_path: Path to the .xlsx file.
        sheet_identifier: Name (str) or index (int) of the sheet.
        template_name: Unique name for storing the template.
        anchor_text: Text to find the anchor cell.
        num_below: Number of cells below anchor to extract.
        num_right: Number of columns right of anchor to extract (full height).
        force_reload: If True, reload even if in cache.
        create_copy: If True, create a copy and read from it.

    Returns:
        True if template loaded/found, False otherwise.
    """
    prefix = f"[TemplateManager.load_template(name='{template_name}')]"
    # --- Cache check ---
    if not force_reload and template_name in _template_cache:
        logging.info(f"{prefix} Template '{template_name}' already found in cache. Skipping reload.")
        return True

    # --- File existence check ---
    if not os.path.exists(workbook_path):
        logging.error(f"{prefix} Original workbook file not found: '{workbook_path}'.")
        return False

    # --- Handle Copy Creation ---
    path_to_load = workbook_path
    copied_path_created = None # Variable defined here, local to this function call
    if create_copy:
        try:
            base, ext = os.path.splitext(workbook_path)
            path_to_load = f"{base}{COPY_SUFFIX}{ext}"
            if os.path.exists(path_to_load):
                 try:
                      os.remove(path_to_load)
                      logging.debug(f"{prefix} Removed existing copy '{path_to_load}' before creating new one.")
                 except Exception as remove_err:
                      logging.warning(f"{prefix} Could not remove existing copy '{path_to_load}': {remove_err}.")
            shutil.copy2(workbook_path, path_to_load) # copy2 preserves metadata
            copied_path_created = path_to_load # Store path if copy succeeded
            logging.info(f"{prefix} Created temporary copy for processing: '{path_to_load}'")
        except Exception as copy_err:
            logging.error(f"{prefix} Failed to create copy of '{workbook_path}': {copy_err}", exc_info=True)
            # Clean up potentially partially created copy if error occurs during creation itself
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass # Ignore cleanup error
            return False

    # --- Load Workbook ---
    workbook = None
    try:
        logging.info(f"{prefix} Attempting to load workbook: '{path_to_load}'")
        try:
            # data_only=True reads cell values, not formulas.
            # read_only=False allows potential modification (though we don't modify here, it avoids some issues)
            # keep_vba=False might be useful if VBA is not needed.
            workbook = openpyxl.load_workbook(path_to_load, data_only=True, read_only=False)
            logging.debug(f"{prefix} Workbook loaded (read_only=False, data_only=True).")
        except PermissionError:
            logging.error(f"{prefix} Permission denied trying to read '{path_to_load}'. Is it open elsewhere?")
            # Cleanup copy if it was created
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
            return False
        except FileNotFoundError:
            logging.error(f"{prefix} File not found at path specified for loading: '{path_to_load}'.")
            # No copy to clean up if the source wasn't found for copying initially
            return False
        except Exception as load_err:
            logging.error(f"{prefix} Failed to load workbook '{path_to_load}': {load_err}", exc_info=True)
            if workbook:
                try: workbook.close() # Attempt to close if object exists
                except Exception: pass
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
            return False

        if not workbook:
             logging.error(f"{prefix} Workbook object is None or invalid after load attempt.")
             if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
             return False

        # --- Find Worksheet ---
        worksheet = None
        sheet_title_for_log = ""
        try:
            if isinstance(sheet_identifier, str):
                if sheet_identifier in workbook.sheetnames:
                    worksheet = workbook[sheet_identifier]
                    sheet_title_for_log = worksheet.title
                else:
                    logging.error(f"{prefix} Sheet name '{sheet_identifier}' not found in '{path_to_load}'. Available: {workbook.sheetnames}")
                    workbook.close(); return False # Exit early, cleanup handled in finally
            elif isinstance(sheet_identifier, int):
                if 0 <= sheet_identifier < len(workbook.sheetnames):
                    worksheet = workbook.worksheets[sheet_identifier]
                    sheet_title_for_log = worksheet.title
                else:
                    logging.error(f"{prefix} Sheet index {sheet_identifier} out of range ({len(workbook.sheetnames)} sheets) in '{path_to_load}'.")
                    workbook.close(); return False # Exit early, cleanup handled in finally
            else:
                logging.error(f"{prefix} Invalid sheet_identifier type '{type(sheet_identifier)}'. Must be str or int.")
                workbook.close(); return False # Exit early, cleanup handled in finally
        except Exception as sheet_access_err:
             logging.error(f"{prefix} Error accessing sheet using identifier '{sheet_identifier}': {sheet_access_err}", exc_info=True)
             try: workbook.close()
             except Exception: pass
             return False # Exit early, cleanup handled in finally

        if not isinstance(worksheet, Worksheet):
             logging.error(f"{prefix} Could not obtain a valid Worksheet object from '{path_to_load}' using identifier '{sheet_identifier}'.")
             try: workbook.close()
             except Exception: pass
             return False # Exit early, cleanup handled in finally
        logging.info(f"{prefix} Successfully accessed sheet: '{sheet_title_for_log}'.")

        # --- Find Anchor Cell ---
        logging.info(f"{prefix} Searching for anchor cell '{anchor_text}' (using merge-aware finder)...")
        anchor_cell = _find_anchor_cell(worksheet, anchor_text)
        if not anchor_cell:
            logging.error(f"{prefix} Failed to find anchor cell '{anchor_text}' in sheet '{sheet_title_for_log}'.")
            workbook.close(); return False # Exit early, cleanup handled in finally

        # --- Extract Properties ---
        logging.info(f"{prefix} Anchor found at {anchor_cell.coordinate}. Extracting properties (full height right)...")
        relative_data = _extract_relative_cells(worksheet, anchor_cell, num_below, num_right)

        # --- Store in Cache ---
        if relative_data is not None:
            # Log a sample of the data if debug is enabled
            if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
                 log_data_sample = {}
                 for k, v in relative_data.items():
                      if isinstance(v, list):
                           log_data_sample[k] = f"List(len={len(v)})" # Summarize lists
                      else:
                           log_data_sample[k] = v # Keep dicts/None as is
                 log_str = pprint.pformat(log_data_sample, indent=2, width=120)
                 if len(log_str) > MAX_LOG_DICT_LEN:
                      log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... [truncated]"
                 logging.debug(f"{prefix} Extracted data structure sample:\n{log_str}")

            _template_cache[template_name] = relative_data
            logging.info(f"{prefix} Properties extracted and stored in cache as '{template_name}'.")
            # Success path
            workbook.close()
            # NOTE: Cleanup of the copy (if created) happens *outside* this function
            # in the __main__ block or calling code AFTER load_template returns.
            return True
        else:
            logging.error(f"{prefix} Failed extraction process in sheet '{sheet_title_for_log}' (_extract_relative_cells returned None).")
            workbook.close()
            return False # Extraction failed, cleanup handled in finally

    except Exception as e:
        logging.error(f"{prefix} Unexpected error during template loading/extraction: {e}", exc_info=True)
        if workbook:
             try: workbook.close()
             except Exception: pass
        # Indicate failure
        return False
    finally:
        # --- IMPORTANT: Cleanup Copy ---
        # This block ensures the copy is removed if it was created, regardless of
        # whether the function succeeded or failed (unless an error occurred during copy creation itself).
        if copied_path_created and os.path.exists(copied_path_created):
            try:
                os.remove(copied_path_created)
                logging.info(f"{prefix} Cleaned up temporary copy file: '{copied_path_created}'")
            except Exception as remove_err:
                logging.warning(f"{prefix} Could not remove temporary copy '{copied_path_created}' during cleanup: {remove_err}")


# get_template - Retrieves template (structure reflects full height right)
def get_template(template_name: str) -> Optional[Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]]]:
    """
    Retrieves template data from cache. 'right_N' keys map to lists of cell data dicts.
    Returns a shallow copy.

    Args:
        template_name: The name of the template to retrieve.

    Returns:
        Template data dictionary or None if not found.
    """
    prefix = f"[TemplateManager.get_template(name='{template_name}')]"
    template = _template_cache.get(template_name)
    if template is not None:
        logging.info(f"{prefix} Template '{template_name}' found in cache.")
        # Return a shallow copy to prevent accidental modification of the cached original
        return template.copy()
    else:
        logging.warning(f"{prefix} Template '{template_name}' not found in cache.")
        return None

# clear_template_cache - Clears the entire cache
def clear_template_cache():
    """Clears all loaded templates from the in-memory cache."""
    prefix = "[TemplateManager.clear_template_cache]"
    global _template_cache
    count = len(_template_cache)
    _template_cache = {}
    logging.info(f"{prefix} Cleared {count} template(s) from the cache.")


# create_xlsx_from_template - Handles Anchor Merge & Right Lists (FIXED FOR MERGEDCELL)
def create_xlsx_from_template(
    template_data: Optional[Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]]],
    output_path: str,
    start_cell_coord: str = "A1",
    sheet_name: str = "Generated Template"
) -> bool:
    """
    Creates XLSX file from template data. Handles anchor merge and places
    vertical lists of cells from 'right_N' keys correctly relative to anchor.
    Correctly handles writing to sheets with merged cells.

    Args:
        template_data: Template data dict (potentially from get_template()).
        output_path: Path for the new XLSX file.
        start_cell_coord: Top-left cell for placing the anchor's data/merge.
        sheet_name: Name for the sheet in the new workbook.

    Returns:
        True if successful, False otherwise.
    """
    prefix = "[TemplateManager.create_xlsx_from_template (Handles Right Lists)]"

    if not template_data:
        logging.error(f"{prefix} Input template_data is None or empty. Cannot create file.")
        return False

    wb = None
    try:
        # Create workbook and sheet
        wb = Workbook()
        ws = wb.active
        if ws is None:
            logging.error(f"{prefix} Could not get active worksheet from new workbook.")
            return False
        ws.title = sheet_name
        logging.info(f"{prefix} Created new workbook with sheet '{sheet_name}'.")

        # Determine start coordinates
        try:
            start_col_str, start_row_str = coordinate_from_string(start_cell_coord)
            start_col_anchor = column_index_from_string(start_col_str)
            start_row_anchor = int(start_row_str)
            logging.info(f"{prefix} Anchor/relative data placement starts at '{start_cell_coord}' (R={start_row_anchor}, C={start_col_anchor}).")
        except (ValueError, TypeError) as e:
            logging.error(f"{prefix} Invalid start_cell_coord '{start_cell_coord}': {e}. Using A1 as default.", exc_info=False)
            start_col_anchor, start_row_anchor = 1, 1

        # --- Helper: Apply single cell properties ---
        def _apply_props(target_cell: Cell, cell_data_dict: Dict[str, Any], target_r: int, target_c: int):
             """Applies properties from dict to cell. Skips MergedCells."""
             # --- VITAL CHECK for MergedCell ---
             # If the target cell is part of a merge (and not the top-left),
             # openpyxl returns a MergedCell object. We cannot set properties
             # directly on these. They inherit from the top-left cell.
             if isinstance(target_cell, MergedCell):
                 logging.debug(f"{prefix} Skipping property application for MergedCell at {target_cell.coordinate}. Properties are handled by its top-left cell.")
                 return # Do nothing for this cell

             # --- Proceed only if it's a regular cell and data is valid ---
             if not isinstance(cell_data_dict, dict):
                 logging.warning(f"{prefix} Invalid cell_data_dict type ({type(cell_data_dict)}) for cell {target_cell.coordinate}. Skipping property application.")
                 return

             if 'error' in cell_data_dict:
                 logging.warning(f"{prefix} Skipping property application for cell {target_cell.coordinate} due to extraction error: {cell_data_dict['error']}")
                 return

             # Apply Value
             target_cell.value = cell_data_dict.get('value')

             # Apply Height (Row Property)
             height = cell_data_dict.get('height')
             if height is not None:
                 try:
                     height_f = float(height)
                     if height_f >= 0:
                         # Avoid setting if already correct to potentially improve performance slightly
                         current_rd = ws.row_dimensions.get(target_r)
                         current_height = current_rd.height if current_rd and current_rd.height is not None else None
                         # Use a small tolerance for floating point comparison if necessary
                         if current_height is None or abs(current_height - height_f) > 1e-6:
                             ws.row_dimensions[target_r].height = height_f
                             # logging.debug(f"{prefix} Set row {target_r} height to {height_f:.2f}")
                     else:
                         logging.warning(f"{prefix} Ignored negative height {height_f} for row {target_r}")
                 except (ValueError, TypeError):
                     logging.warning(f"{prefix} Invalid height value '{height}' for row {target_r}")

             # Apply Width (Column Property)
             width = cell_data_dict.get('width')
             if width is not None:
                 col_letter = get_column_letter(target_c)
                 try:
                     width_f = float(width)
                     if width_f >= 0:
                         # Avoid setting if already correct
                         current_cd = ws.column_dimensions.get(col_letter)
                         current_width = current_cd.width if current_cd and current_cd.width is not None else None
                         # Use a small tolerance for floating point comparison if necessary
                         if current_width is None or abs(current_width - width_f) > 1e-6:
                            ws.column_dimensions[col_letter].width = width_f
                            # logging.debug(f"{prefix} Set col {col_letter} width to {width_f:.2f}")
                     else:
                         logging.warning(f"{prefix} Ignored negative width {width_f} for col {col_letter}")
                 except (ValueError, TypeError):
                     logging.warning(f"{prefix} Invalid width value '{width}' for col {col_letter}")

             # Apply Border
             border_styles = cell_data_dict.get('border')
             # Make sure border_styles is a dictionary before proceeding
             if isinstance(border_styles, dict):
                 # Create a new Border object based on styles, applying defaults if missing
                 new_border = Border(
                    left=Side(style=border_styles.get('left')) if border_styles.get('left') else Side(),
                    right=Side(style=border_styles.get('right')) if border_styles.get('right') else Side(),
                    top=Side(style=border_styles.get('top')) if border_styles.get('top') else Side(),
                    bottom=Side(style=border_styles.get('bottom')) if border_styles.get('bottom') else Side()
                 )
                 # Only apply if the new border is different from the current one
                 # This avoids applying default empty borders unnecessarily
                 current_border = target_cell.border if target_cell.has_style else Border()
                 if new_border != current_border:
                      target_cell.border = new_border
                      # logging.debug(f"{prefix} Applied border to {target_cell.coordinate}")
             elif target_cell.has_style and target_cell.border != Border():
                 # Explicitly clear border if template had none but cell has one
                 target_cell.border = Border()
                 # logging.debug(f"{prefix} Cleared border from {target_cell.coordinate}")

        # --- Process template data ---
        processed_cells = 0
        max_target_row, max_target_col = 0, 0

        # Sort keys for predictable processing order (Anchor -> Bottom -> Right)
        # Ensure integer conversion handles potential errors gracefully
        def sort_key_func(k):
            if k == ANCHOR_KEY: return (0, 0)
            parts = k.split('_')
            type_order = 99
            index = 0
            if len(parts) == 2 and parts[1].isdigit():
                index = int(parts[1])
                if parts[0] == 'bottom': type_order = 1
                elif parts[0] == 'right': type_order = 2
            return (type_order, index)

        sorted_keys = sorted(template_data.keys(), key=sort_key_func)

        for key in sorted_keys:
            cell_value_or_list = template_data[key]

            if cell_value_or_list is None:
                logging.debug(f"{prefix} Skipping key '{key}': Value is None.")
                continue

            # --- Handle Anchor Cell ---
            if key == ANCHOR_KEY:
                if isinstance(cell_value_or_list, dict) and 'error' not in cell_value_or_list:
                    anchor_data = cell_value_or_list
                    target_row, target_col = start_row_anchor, start_col_anchor
                    # Get the top-left cell for applying props and potentially merging
                    target_cell = ws.cell(row=target_row, column=target_col)
                    logging.debug(f"{prefix} Processing '{key}': Base Target Cell {target_cell.coordinate}")

                    # Apply properties FIRST to the top-left cell
                    _apply_props(target_cell, anchor_data, target_row, target_col)
                    processed_cells += 1
                    max_target_row = max(max_target_row, target_row)
                    max_target_col = max(max_target_col, target_col)

                    # Handle merging AFTER applying props to the top-left cell
                    merge_info = anchor_data.get('merge_info')
                    target_end_row, target_end_col = target_row, target_col # Initialize end coords
                    if isinstance(merge_info, dict):
                        rows_to_merge = merge_info.get('rows', 1)
                        cols_to_merge = merge_info.get('cols', 1)
                        if rows_to_merge > 1 or cols_to_merge > 1:
                            target_end_row = target_row + rows_to_merge - 1
                            target_end_col = target_col + cols_to_merge - 1
                            try:
                                merge_range_str = f"{target_cell.coordinate}:{get_column_letter(target_end_col)}{target_end_row}"
                                logging.info(f"{prefix} Applying merge for anchor: {merge_range_str}")
                                # Ensure we don't merge over existing merged cells if logic requires
                                # (basic merge_cells handles simple cases)
                                ws.merge_cells(start_row=target_row, start_column=target_col,
                                               end_row=target_end_row, end_column=target_end_col)
                                # Update max bounds based on merge
                                max_target_row = max(max_target_row, target_end_row)
                                max_target_col = max(max_target_col, target_end_col)
                            except Exception as merge_err:
                                logging.error(f"{prefix} Failed to apply merge for anchor ({merge_range_str}): {merge_err}", exc_info=True)
                                # Reset end coords if merge failed
                                target_end_row, target_end_col = target_row, target_col
                else:
                     logging.warning(f"{prefix} Skipping key '{key}': Invalid data format ({type(cell_value_or_list)}) or contains error.")
                continue # Move to next key

            # --- Handle Bottom Cells (Single Dict) ---
            elif key.startswith('bottom_'):
                 if isinstance(cell_value_or_list, dict) and 'error' not in cell_value_or_list:
                     cell_data_dict = cell_value_or_list
                     try:
                         offset = int(key.split('_')[1])
                         target_row = start_row_anchor + offset
                         target_col = start_col_anchor
                         # Basic bounds check
                         if target_row <= 0 or target_col <= 0: continue

                         target_coord_log = f"{get_column_letter(target_col)}{target_row}"
                         logging.debug(f"{prefix} Processing '{key}': Target {target_coord_log}")
                         # Get cell *before* applying props to handle potential MergedCell
                         target_cell = ws.cell(row=target_row, column=target_col)
                         _apply_props(target_cell, cell_data_dict, target_row, target_col)
                         # Only count if not skipped by _apply_props (e.g., MergedCell)
                         if not isinstance(target_cell, MergedCell): processed_cells += 1
                         max_target_row = max(max_target_row, target_row)
                         max_target_col = max(max_target_col, target_col)
                     except (ValueError, IndexError, TypeError) as e:
                         logging.warning(f"{prefix} Skipping '{key}': Error calculating position or applying props: {e}")
                 else:
                     logging.warning(f"{prefix} Skipping key '{key}': Invalid data format ({type(cell_value_or_list)}) or contains error.")
                 continue # Move to next key

            # --- Handle Right Cells (List of Dicts) ---
            elif key.startswith('right_'):
                 if isinstance(cell_value_or_list, list):
                     cell_data_list = cell_value_or_list
                     logging.debug(f"{prefix} Processing '{key}' (List length={len(cell_data_list)})...")
                     try:
                         offset_col = int(key.split('_')[1])
                         target_base_col = start_col_anchor + offset_col
                         # Basic bounds check
                         if target_base_col <= 0: continue

                         # Iterate through the vertical list for this column
                         for row_offset, cell_data_item in enumerate(cell_data_list):
                             # Skip if the item itself is None or invalid
                             if cell_data_item is None: continue
                             if not isinstance(cell_data_item, dict) or 'error' in cell_data_item:
                                 logging.warning(f"{prefix}   Skipping row offset {row_offset} for '{key}': Invalid item data ({cell_data_item}).")
                                 continue

                             target_row = start_row_anchor + row_offset
                             target_col = target_base_col
                             # Basic bounds check
                             if target_row <= 0: continue

                             cell_coord_log = f"{get_column_letter(target_col)}{target_row}"
                             logging.debug(f"{prefix}   Applying data for row offset {row_offset} to {cell_coord_log}")
                             # Get cell *before* applying props
                             target_cell = ws.cell(row=target_row, column=target_col)
                             _apply_props(target_cell, cell_data_item, target_row, target_col)
                             # Only count if not skipped by _apply_props
                             if not isinstance(target_cell, MergedCell): processed_cells += 1
                             max_target_row = max(max_target_row, target_row)
                             max_target_col = max(max_target_col, target_col)

                     except (ValueError, IndexError, TypeError) as e:
                         logging.warning(f"{prefix} Skipping '{key}' list processing: Error calculating position or applying props: {e}")
                 else:
                      logging.warning(f"{prefix} Skipping key '{key}': Expected a list for 'right_N', but got {type(cell_value_or_list)}.")
                 continue # Move to next key

            else:
                logging.warning(f"{prefix} Skipping unrecognized key '{key}' in template data.")

        # --- Save the workbook ---
        if processed_cells > 0 or template_data.get(ANCHOR_KEY): # Save even if only anchor was processed
            logging.info(f"{prefix} Processed relevant data (max R={max_target_row}, C={max_target_col}). Saving workbook to '{output_path}'...")
            wb.save(output_path)
            logging.info(f"{prefix} Successfully saved workbook to '{output_path}'.")
            return True
        else:
            logging.warning(f"{prefix} No valid, non-merged cells processed from template data. Output file '{output_path}' might be empty or only contain merged cells.")
            # Still attempt to save, as merged cells might be the desired output
            try:
                wb.save(output_path)
                logging.info(f"{prefix} Saved workbook (potentially empty or only merges) to '{output_path}'.")
                return True
            except Exception as save_err:
                 logging.error(f"{prefix} Failed to save workbook to '{output_path}': {save_err}", exc_info=True)
                 return False

    except Exception as e:
        logging.error(f"{prefix} Failed to create or save workbook at '{output_path}': {e}", exc_info=True)
        return False
    finally:
         if wb:
              try: wb.close()
              except Exception as close_err: logging.warning(f"{prefix} Error closing workbook object: {close_err}")


# --- Example Usage ---
if __name__ == "__main__":
    # Set level to DEBUG to see detailed extraction, list processing, merge handling
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s')

    logging.info("--- Running Template Manager Example (Handles Full Height Right Extraction & Creation) ---")
    logging.info(f"Current Working Directory: {os.getcwd()}")

    # >>> IMPORTANT: Verify this path is correct for your system <<<
    # Use raw string (r"...") or double backslashes ("\\") for Windows paths
    # target_workbook = r"C:\path\to\your\dap.xlsx"
    target_workbook = "dap.xlsx" # Assumes it's in the same directory as the script

    # --- Critical Check: Ensure workbook exists ---
    if not os.path.exists(target_workbook):
        logging.critical(f"*** CRITICAL ERROR: Workbook not found at '{os.path.abspath(target_workbook)}'. Example cannot run. Please correct the path. ***")
        exit(1) # Exit with a non-zero code

    # --- Parameters ---
    target_sheet = "Packing list"             # The exact name of the sheet
    target_anchor = "Mark & Nº"             # The text to search for as anchor
    cells_to_extract_below = 5              # How many single cells *below* anchor
    cells_to_extract_right = 8              # How many columns *to the right* (full height slices)
    template_key = "DapPackingListFullHeightRight_Final" # Unique name for this template run
    should_create_copy = False # Set to True to test copy creation/cleanup

    # --- Output file ---
    output_filename = "dummy_generated_template_full_height_right_Final.xlsx" # Updated output name
    start_coordinate_in_output = "C3" # Where anchor's top-left maps to in output

    # Path for potential copy file (used for initial cleanup check)
    # Construct the expected copy name based on the target_workbook path
    target_wb_base, target_wb_ext = os.path.splitext(target_workbook)
    target_copy_file = f"{target_wb_base}{COPY_SUFFIX}{target_wb_ext}"

    # --- Optional: Cleanup any previous output/copy files ---
    files_to_clean = [target_copy_file, output_filename]
    logging.info("--- Pre-run Cleanup ---")
    for f_path in files_to_clean:
        abs_f_path = os.path.abspath(f_path)
        if os.path.exists(abs_f_path):
            try:
                os.remove(abs_f_path)
                logging.info(f"Removed existing file: '{abs_f_path}'")
            except Exception as e:
                logging.warning(f"Could not remove pre-existing file '{abs_f_path}': {e}")
        else:
            logging.debug(f"Pre-existing file not found (no cleanup needed): '{abs_f_path}'")

    # --- Load Template (should handle merge, extract lists for right) ---
    logging.info(f"\n>>> Attempting Load (Full Height Right): Source '{target_workbook}'...")
    load_success = load_template(
        workbook_path=target_workbook,
        sheet_identifier=target_sheet,
        template_name=template_key,
        anchor_text=target_anchor,
        num_below=cells_to_extract_below,
        num_right=cells_to_extract_right,
        force_reload=True,      # Force reload for demonstration
        create_copy=should_create_copy # Use the variable defined above
    )

    # --- Process Result & Generate Dummy XLSX ---
    if load_success:
        logging.info(f"Successfully loaded template as '{template_key}'. Retrieving data...")
        retrieved_data = get_template(template_key)

        if retrieved_data is not None:
            # Debug: Print structure of anchor and an example 'right_N' key if needed
            if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
                print(f"\n--- Extracted Data Sample ('{template_key}') ---")
                anchor_data_sample = retrieved_data.get(ANCHOR_KEY)
                print("Anchor Data:")
                pprint.pprint(anchor_data_sample, indent=2, width=100)

                # Determine a valid 'right_N' key that exists, if possible
                example_right_key = None
                for i in range(1, cells_to_extract_right + 1):
                    key_to_check = f'right_{i}'
                    if key_to_check in retrieved_data:
                        example_right_key = key_to_check
                        break # Found one

                if example_right_key:
                    right_data_sample = retrieved_data.get(example_right_key)
                    print(f"\nData for '{example_right_key}' (should be a list):")
                    # Limit print length for potentially long lists
                    if isinstance(right_data_sample, list):
                         pprint.pprint(right_data_sample[:5], indent=2, width=100) # Print first 5 items
                         if len(right_data_sample) > 5: print("  ...")
                    else:
                         pprint.pprint(right_data_sample, indent=2, width=100) # Print if not a list (unexpected)

                print("--- End Sample ---")

            # --- Call function to create the dummy XLSX ---
            logging.info(f"\n>>> Attempting to create dummy XLSX from extracted template data...")
            create_success = create_xlsx_from_template(
                template_data=retrieved_data,
                output_path=output_filename,
                start_cell_coord=start_coordinate_in_output,
                sheet_name=f"Generated_{target_sheet}" # Make sheet name dynamic
            )

            if create_success:
                logging.info(f"Successfully created dummy template file: '{output_filename}'")
                absolute_output_path = os.path.abspath(output_filename)
                # Provide clear user feedback
                print(f"\n+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                print(f"+++ Dummy XLSX file created successfully!")
                print(f"+++ Location: {absolute_output_path}")
                print(f"+++ Anchor merge (if any) and cells to the right (full height)")
                print(f"+++ should be placed starting relative to anchor at: {start_coordinate_in_output}")
                print(f"+++ Please open the file to visually inspect the result.")
                print(f"+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            else:
                logging.error(f"Failed to create dummy template file '{output_filename}'. Check logs for specific errors.")

        else:
            # This case should ideally not happen if load_success was True
            logging.error(f"CRITICAL INTERNAL ERROR: Failed to retrieve template data ('{template_key}') using get_template after load_template reported success.")
    else:
        logging.error(f"Failed to load template '{template_key}' from '{target_workbook}'. Cannot create dummy file.")

    # --- Final Cleanup (Checks the expected copy path again) ---
    # This is mainly relevant if create_copy was True and something failed *after*
    # the copy was created but *before* the 'finally' block in load_template ran,
    # or if the user manually stops the script mid-execution.
    # The 'finally' block in load_template *should* handle cleanup in normal operation.
    logging.info("--- Post-run Cleanup Check ---")
    abs_target_copy_file = os.path.abspath(target_copy_file)
    logging.debug(f"Checking for cleanup target: '{abs_target_copy_file}'")
    if should_create_copy and os.path.exists(abs_target_copy_file):
        try:
            os.remove(abs_target_copy_file)
            logging.info(f"Cleaned up copy file '{abs_target_copy_file}' after run (final check).")
        except Exception as e:
            logging.warning(f"Could not remove copy file '{abs_target_copy_file}' during final cleanup check: {e}")
    else:
        if should_create_copy:
            logging.debug(f"Copy file '{abs_target_copy_file}' not found during final check (already cleaned up by load_template).")
        else:
            logging.debug(f"No copy file expected or found at '{abs_target_copy_file}' during final cleanup (create_copy=False).")

    logging.info("\n--- Template Manager Example Finished ---")