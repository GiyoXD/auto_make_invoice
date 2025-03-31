# template_manager.py

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell, MergedCell
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
# Import range utils for merge checking
from openpyxl.utils.cell import range_boundaries

import logging
import os
import pprint
import shutil

# Import necessary types from 'typing'
from typing import Dict, List, Any, Union, Optional, Tuple

# --- Constants ---
DEFAULT_ANCHOR_TEXT = "Mark & Nº"
DEFAULT_CELLS_BELOW = 5
DEFAULT_CELLS_RIGHT = 10
MAX_LOG_DICT_LEN = 2000 # Reduced as detailed dict logging is now DEBUG level
COPY_SUFFIX = "_copy"
ANCHOR_KEY = 'anchor'

# --- Module Level Cache ---
_template_cache: Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]] = {}


# --- Helper Functions ---

# _find_anchor_cell - Merge Aware
def _find_anchor_cell(worksheet: Worksheet, anchor_text: str) -> Optional[Cell]:
    """
    Finds the FIRST cell containing the anchor text during a row-by-row scan.
    If the cell found is part of a merged range, it returns the TOP-LEFT cell.

    Args:
        worksheet: The openpyxl worksheet object.
        anchor_text: The text to search for (case-sensitive).

    Returns:
        The top-left Cell object of the merge (or the cell itself) containing
        the text, otherwise None.
    """
    prefix = "[TemplateManager._find_anchor_cell]"
    # This log is useful if anchor isn't found, keep as DEBUG
    logging.debug(f"{prefix} Searching for FIRST cell containing anchor text '{anchor_text}' in sheet '{worksheet.title}'...")

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
            current_coord = cell.coordinate
            cell_value = cell.value

            if cell_value is None:
                is_top_left_of_merge = current_coord in merge_map and merge_map[current_coord] == current_coord
                if not is_top_left_of_merge:
                    continue

            try:
                 if anchor_text in str(cell_value):
                     # Log detection and merge check at DEBUG level
                     logging.debug(f"{prefix} Found anchor text potentially in cell {current_coord}.")
                     if current_coord in merge_map:
                         top_left_coord = merge_map[current_coord]
                         # Log merge detail at DEBUG level
                         logging.debug(f"{prefix} Cell {current_coord} is part of merged range starting at {top_left_coord}. Returning top-left cell object.")
                         if current_coord == top_left_coord:
                             return cell
                         else:
                             tl_col_str, tl_row_str = coordinate_from_string(top_left_coord)
                             tl_col = column_index_from_string(tl_col_str)
                             tl_row = int(tl_row_str)
                             try:
                                 return worksheet.cell(row=tl_row, column=tl_col)
                             except IndexError:
                                 logging.warning(f"{prefix} Calculated top-left {top_left_coord} seems out of bounds. Returning originally found cell {current_coord}.")
                                 return cell
                     else:
                         # Log non-merge detail at DEBUG level
                         logging.debug(f"{prefix} Cell {current_coord} contains text and is not merged. Returning this cell.")
                         return cell
            except TypeError:
                 # Low-level type issue, keep as DEBUG
                 logging.debug(f"{prefix} Could not compare anchor text with cell {current_coord} value '{cell_value}' (Type: {type(cell_value)}). Skipping.")
                 continue
            except Exception as e:
                 logging.warning(f"{prefix} Unexpected error checking cell {current_coord} value '{cell_value}': {e}")
                 continue

    # Failure to find is important, keep as WARNING
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
    prefix = "[TemplateManager._extract_relative_cells]"
    if not anchor_cell:
        logging.error(f"{prefix} Invalid anchor_cell provided.")
        return None

    relative_data: Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]] = {}
    anchor_row = anchor_cell.row
    anchor_col = anchor_cell.column
    anchor_coord = anchor_cell.coordinate

    # Keep this INFO log - shows start of extraction and parameters
    logging.info(f"{prefix} Extracting properties relative to anchor {anchor_coord} (Below={num_below}, Right={num_right}).")

    # Get default dimensions
    default_row_height = worksheet.sheet_format.defaultRowHeight if worksheet.sheet_format and worksheet.sheet_format.defaultRowHeight else 15.0
    default_col_width = worksheet.sheet_format.defaultColWidth if worksheet.sheet_format and worksheet.sheet_format.defaultColWidth else 8.43

    # --- Helper: Extract single cell properties (Keep detailed logging as DEBUG) ---
    def _get_cell_props(cell: Cell, target_r: int, target_c: int, is_anchor: bool = False) -> Dict[str, Any]:
        """Extracts value, dimensions, border, alignment, and merge info (if anchor)."""
        props: Dict[str, Any] = {}
        try:
            props['value'] = cell.value
            rd = worksheet.row_dimensions.get(target_r)
            props['height'] = rd.height if rd and rd.height is not None else default_row_height
            col_letter = get_column_letter(target_c)
            cd = worksheet.column_dimensions.get(col_letter)
            props['width'] = cd.width if cd and cd.width is not None else default_col_width
            border_styles: Dict[str, Optional[str]] = {'left': None, 'right': None, 'top': None, 'bottom': None}
            if cell.has_style and cell.border:
                border_styles['left'] = cell.border.left.style if cell.border.left else None
                border_styles['right'] = cell.border.right.style if cell.border.right else None
                border_styles['top'] = cell.border.top.style if cell.border.top else None
                border_styles['bottom'] = cell.border.bottom.style if cell.border.bottom else None
            props['border'] = border_styles
            alignment_styles: Dict[str, Any] = {'horizontal': None, 'vertical': None, 'wrap_text': None}
            if cell.has_style and cell.alignment:
                alignment_styles['horizontal'] = cell.alignment.horizontal
                alignment_styles['vertical'] = cell.alignment.vertical
                alignment_styles['wrap_text'] = cell.alignment.wrap_text
            props['alignment'] = alignment_styles
            anchor_merge_rows = 1
            if is_anchor:
                merge_info = None
                for mc_range in worksheet.merged_cells.ranges:
                    if mc_range.min_row == target_r and mc_range.min_col == target_c:
                        rows_merged = mc_range.max_row - mc_range.min_row + 1
                        cols_merged = mc_range.max_col - mc_range.min_col + 1
                        if rows_merged > 1 or cols_merged > 1:
                            merge_info = {'rows': rows_merged, 'cols': cols_merged}
                            anchor_merge_rows = rows_merged
                            # Detailed merge info is DEBUG level
                            logging.debug(f"{prefix}   -> Anchor cell {cell.coordinate} is top-left of merge: {rows_merged}r x {cols_merged}c.")
                        break
                if merge_info:
                    props['merge_info'] = merge_info
                props['_anchor_merge_rows'] = anchor_merge_rows

            # Very detailed extraction log - keep as DEBUG
            if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
                border_log_repr = f"L={border_styles['left']}, R={border_styles['right']}, T={border_styles['top']}, B={border_styles['bottom']}"
                alignment_log_repr = f"H={alignment_styles['horizontal']}, V={alignment_styles['vertical']}, Wrap={alignment_styles['wrap_text']}"
                merge_log_repr = f", Merged={props['merge_info']}" if 'merge_info' in props else ""
                val_repr = str(props.get('value', ''))[:30]
                cell_type_log = f"(Type: {type(cell).__name__})"
                logging.debug(
                    f"{prefix}     -> Extracted (R={target_r}, C={target_c}, Cell: {cell.coordinate} {cell_type_log}): "
                    f"Val='{val_repr}...', H={props['height']:.1f}, W={props['width']:.1f}, "
                    f"Border={{{border_log_repr}}}, Alignment={{{alignment_log_repr}}}{merge_log_repr}"
                )
            return props
        except Exception as e:
            # Error reading cell is important, keep as ERROR
            logging.error(f"{prefix} Error reading properties for cell at R={target_r}, C={target_c}: {e}", exc_info=False)
            return {'error': f"Error reading cell properties: {e}"}

    # --- Extract Anchor Cell Properties ---
    # Detail of *which* cell is being processed is DEBUG
    logging.debug(f"{prefix} Extracting ANCHOR cell ({ANCHOR_KEY}) at {anchor_coord}...")
    anchor_props = _get_cell_props(anchor_cell, anchor_row, anchor_col, is_anchor=True)
    relative_data[ANCHOR_KEY] = anchor_props
    anchor_merge_height = 1
    if isinstance(anchor_props, dict):
        anchor_merge_height = anchor_props.get('_anchor_merge_rows', 1)
    # Anchor height detail is DEBUG
    logging.debug(f"{prefix}   Anchor merge height determined: {anchor_merge_height} rows.")


    # --- Extract Cells Below (Single Cells) ---
    # Detail of starting the 'below' loop is DEBUG
    logging.debug(f"{prefix} Extracting cells BELOW anchor (single cells)...")
    for i in range(1, num_below + 1):
        target_row = anchor_row + i
        target_col = anchor_col
        key_name = f'bottom_{i}'
        if target_row > worksheet.max_row or target_col > worksheet.max_column:
            # Out of bounds is a potential issue, keep as WARNING
            logging.warning(f"{prefix} Target {key_name} (R={target_row}, C={target_col}) out of bounds. Storing None.")
            relative_data[key_name] = None
            continue
        try:
             cell_to_extract = worksheet.cell(row=target_row, column=target_col)
             relative_data[key_name] = _get_cell_props(cell_to_extract, target_row, target_col)
        except IndexError:
             # Failure to access cell is a potential issue, keep as WARNING
             logging.warning(f"{prefix} Could not access cell for {key_name} at (R={target_row}, C={target_col}). Storing error.")
             relative_data[key_name] = {'error': f"IndexError accessing cell at R={target_row}, C={target_col}"}
        except Exception as exc:
             # Other errors during extraction are important, keep as ERROR
             logging.error(f"{prefix} Error getting props for {key_name} at (R={target_row}, C={target_col}): {exc}")
             relative_data[key_name] = {'error': f"Error getting props: {exc}"}


    # --- Extract Cells To The RIGHT (Full Height) ---
    # Detail of starting the 'right' loop is DEBUG
    logging.debug(f"{prefix} Extracting cells to the RIGHT of anchor (Full Height: {anchor_merge_height} rows each)...")
    for i in range(1, num_right + 1):
        target_base_col = anchor_col + i
        key_name = f'right_{i}'
        cells_in_column: List[Optional[Dict[str, Any]]] = []

        # Detail of processing each 'right_N' column is DEBUG
        logging.debug(f"{prefix}   Processing '{key_name}' (Column Index: {target_base_col})...")

        for j in range(anchor_merge_height):
            target_row = anchor_row + j
            target_col = target_base_col

            if target_row > worksheet.max_row or target_col > worksheet.max_column:
                # Out of bounds per cell is a potential issue, keep as WARNING
                logging.warning(f"{prefix}     Cell at (R={target_row}, C={target_col}) for '{key_name}' (row offset {j}) out of bounds. Storing None in list.")
                cells_in_column.append(None)
                continue

            try:
                cell_to_extract = worksheet.cell(row=target_row, column=target_col)

                # Detail about MergedCell detection is DEBUG
                if isinstance(cell_to_extract, MergedCell):
                     logging.debug(f"{prefix}       Cell at (R={target_row}, C={target_col}) is a MergedCell. Extracting its properties (value=None, style inherited).")

                props = _get_cell_props(cell_to_extract, target_row, target_col)
                cells_in_column.append(props)

            except IndexError:
                # Failure to access cell is a potential issue, keep as WARNING
                logging.warning(f"{prefix}     IndexError accessing cell for '{key_name}' at (R={target_row}, C={target_col}). Storing error in list.")
                cells_in_column.append({'error': f"IndexError accessing cell at R={target_row}, C={target_col}"})
            except Exception as cell_exc:
                 # Other errors during extraction are important, keep as ERROR
                 logging.error(f"{prefix}     Unexpected error processing cell for '{key_name}' at (R={target_row}, C={target_col}): {cell_exc}. Storing error in list.")
                 cells_in_column.append({'error': f"Unexpected error processing cell: {cell_exc}"})

        relative_data[key_name] = cells_in_column
        # Log list structure only if DEBUG enabled
        if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
            list_repr_items = []
            for item_idx, d in enumerate(cells_in_column):
                if item_idx >= 3 and len(cells_in_column) > 4 :
                    if item_idx == 3: list_repr_items.append("...")
                    if item_idx < len(cells_in_column) -1: continue
                if isinstance(d, dict): list_repr_items.append(f"Dict({len(d)})" if 'error' not in d else "ErrDict")
                else: list_repr_items.append(str(d))
            list_repr = "[" + ", ".join(list_repr_items) + "]"
            logging.debug(f"{prefix}   Stored '{key_name}' as List (Length={len(cells_in_column)}): {list_repr}")

    # Keep this INFO log - shows end of extraction and overall result
    logging.info(f"{prefix} Finished extracting properties. Found data for {len(relative_data)} relative positions.")
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
    # Cache hit is useful INFO
    if not force_reload and template_name in _template_cache:
        logging.info(f"{prefix} Template '{template_name}' already found in cache. Skipping reload.")
        return True

    # File not found is critical ERROR
    if not os.path.exists(workbook_path):
        logging.error(f"{prefix} Original workbook file not found: '{workbook_path}'.")
        return False

    path_to_load = workbook_path
    copied_path_created = None
    if create_copy:
        try:
            base, ext = os.path.splitext(workbook_path)
            path_to_load = f"{base}{COPY_SUFFIX}{ext}"
            if os.path.exists(path_to_load):
                 try:
                      os.remove(path_to_load)
                      # Removing existing copy is a notable action, keep as INFO
                      logging.info(f"{prefix} Removed existing copy '{path_to_load}' before creating new one.")
                 except Exception as remove_err:
                      # Failure to remove is a WARNING
                      logging.warning(f"{prefix} Could not remove existing copy '{path_to_load}': {remove_err}.")
            shutil.copy2(workbook_path, path_to_load)
            copied_path_created = path_to_load
            # Creating copy is notable INFO
            logging.info(f"{prefix} Created temporary copy for processing: '{path_to_load}'")
        except Exception as copy_err:
            # Failure to copy is an ERROR
            logging.error(f"{prefix} Failed to create copy of '{workbook_path}': {copy_err}", exc_info=True)
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
            return False

    workbook = None
    try:
        # Attempting load is useful INFO
        logging.info(f"{prefix} Attempting to load workbook: '{path_to_load}'")
        try:
            workbook = openpyxl.load_workbook(path_to_load, data_only=True, read_only=False)
            # Load details are DEBUG
            logging.debug(f"{prefix} Workbook loaded (read_only=False, data_only=True).")
        except PermissionError:
            # Permission error is critical ERROR
            logging.error(f"{prefix} Permission denied trying to read '{path_to_load}'. Is it open elsewhere?")
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
            return False
        except FileNotFoundError:
            # File not found (should be caught earlier, but good resilience) is ERROR
            logging.error(f"{prefix} File not found at path specified for loading: '{path_to_load}'.")
            return False
        except Exception as load_err:
            # General load failure is ERROR
            logging.error(f"{prefix} Failed to load workbook '{path_to_load}': {load_err}", exc_info=True)
            if workbook:
                try: workbook.close()
                except Exception: pass
            if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
            return False

        if not workbook:
             # Should not happen if load didn't raise error, but check anyway - ERROR
             logging.error(f"{prefix} Workbook object is None or invalid after load attempt.")
             if copied_path_created and os.path.exists(copied_path_created):
                 try: os.remove(copied_path_created)
                 except: pass
             return False

        worksheet = None
        sheet_title_for_log = ""
        try:
            if isinstance(sheet_identifier, str):
                if sheet_identifier in workbook.sheetnames:
                    worksheet = workbook[sheet_identifier]
                    sheet_title_for_log = worksheet.title
                else:
                    # Sheet not found is ERROR
                    logging.error(f"{prefix} Sheet name '{sheet_identifier}' not found in '{path_to_load}'. Available: {workbook.sheetnames}")
                    workbook.close()
                    return False
            elif isinstance(sheet_identifier, int):
                if 0 <= sheet_identifier < len(workbook.sheetnames):
                    worksheet = workbook.worksheets[sheet_identifier]
                    sheet_title_for_log = worksheet.title
                else:
                    # Index out of range is ERROR
                    logging.error(f"{prefix} Sheet index {sheet_identifier} out of range ({len(workbook.sheetnames)} sheets) in '{path_to_load}'.")
                    workbook.close()
                    return False
            else:
                # Invalid type is ERROR
                logging.error(f"{prefix} Invalid sheet_identifier type '{type(sheet_identifier)}'. Must be str or int.")
                workbook.close()
                return False
        except Exception as sheet_access_err:
             # General sheet access error is ERROR
             logging.error(f"{prefix} Error accessing sheet using identifier '{sheet_identifier}': {sheet_access_err}", exc_info=True)
             try: workbook.close()
             except Exception: pass
             return False

        if not isinstance(worksheet, Worksheet):
             # Should not happen, but check anyway - ERROR
             logging.error(f"{prefix} Could not obtain a valid Worksheet object from '{path_to_load}' using identifier '{sheet_identifier}'.")
             try: workbook.close()
             except Exception: pass
             return False
        # Success accessing sheet is useful INFO
        logging.info(f"{prefix} Successfully accessed sheet: '{sheet_title_for_log}'.")

        # Search start is useful INFO
        logging.info(f"{prefix} Searching for anchor cell '{anchor_text}'...")
        anchor_cell = _find_anchor_cell(worksheet, anchor_text)
        if not anchor_cell:
            # Failure to find anchor is important ERROR
            logging.error(f"{prefix} Failed to find anchor cell '{anchor_text}' in sheet '{sheet_title_for_log}'.")
            workbook.close()
            return False

        # Anchor found and extraction starting is useful INFO
        logging.info(f"{prefix} Anchor found at {anchor_cell.coordinate}. Extracting properties...")
        relative_data = _extract_relative_cells(worksheet, anchor_cell, num_below, num_right)

        if relative_data is not None:
            # Log data sample only if DEBUG enabled
            if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
                 log_data_sample = {}
                 for k, v in relative_data.items():
                      if isinstance(v, list):
                           log_data_sample[k] = f"List(len={len(v)})"
                      else:
                           log_data_sample[k] = v
                 log_str = pprint.pformat(log_data_sample, indent=2, width=120)
                 if len(log_str) > MAX_LOG_DICT_LEN:
                      log_str = log_str[:MAX_LOG_DICT_LEN] + "\n... [truncated]"
                 logging.debug(f"{prefix} Extracted data structure sample:\n{log_str}")

            _template_cache[template_name] = relative_data
            # Overall success is useful INFO
            logging.info(f"{prefix} Properties extracted and stored in cache as '{template_name}'.")
            workbook.close()
            return True
        else:
            # Extraction failure is important ERROR
            logging.error(f"{prefix} Failed extraction process in sheet '{sheet_title_for_log}' (_extract_relative_cells returned None).")
            workbook.close()
            return False

    except Exception as e:
        # Catch-all unexpected error is ERROR
        logging.error(f"{prefix} Unexpected error during template loading/extraction: {e}", exc_info=True)
        if workbook:
             try: workbook.close()
             except Exception: pass
        return False
    finally:
        # Cleanup copy in finally is important INFO/WARNING
        if copied_path_created and os.path.exists(copied_path_created):
            try:
                os.remove(copied_path_created)
                logging.info(f"{prefix} Cleaned up temporary copy file: '{copied_path_created}'")
            except Exception as remove_err:
                logging.warning(f"{prefix} Could not remove temporary copy '{copied_path_created}' during cleanup: {remove_err}")


# get_template - Retrieves template
def get_template(template_name: str) -> Optional[Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]]]:
    """
    Retrieves template data from cache. Returns a shallow copy.

    Args:
        template_name: The name of the template to retrieve.

    Returns:
        Template data dictionary or None if not found.
    """
    prefix = f"[TemplateManager.get_template(name='{template_name}')]"
    template = _template_cache.get(template_name)
    if template is not None:
        # Cache hit is useful INFO
        logging.info(f"{prefix} Template '{template_name}' found in cache.")
        return template.copy()
    else:
        # Cache miss is potential issue, keep as WARNING
        logging.warning(f"{prefix} Template '{template_name}' not found in cache.")
        return None

# clear_template_cache - Clears the entire cache
def clear_template_cache():
    """Clears all loaded templates from the in-memory cache."""
    prefix = "[TemplateManager.clear_template_cache]"
    global _template_cache
    count = len(_template_cache)
    _template_cache = {}
    # Confirmation of cache clear is useful INFO
    logging.info(f"{prefix} Cleared {count} template(s) from the cache.")


# create_xlsx_from_template - Handles Anchor Merge & Right Lists
def create_xlsx_from_template(
    template_data: Optional[Dict[str, Union[Optional[Dict[str, Any]], Optional[List[Optional[Dict[str, Any]]]]]]],
    output_path: str,
    start_cell_coord: str = "A1",
    sheet_name: str = "Generated Template"
) -> bool:
    """
    Creates XLSX file from template data. Handles anchor merge and places
    vertical lists of cells from 'right_N' keys correctly relative to anchor.

    Args:
        template_data: Template data dict (potentially from get_template()).
        output_path: Path for the new XLSX file.
        start_cell_coord: Top-left cell for placing the anchor's data/merge.
        sheet_name: Name for the sheet in the new workbook.

    Returns:
        True if successful, False otherwise.
    """
    prefix = "[TemplateManager.create_xlsx_from_template]"

    if not template_data:
        # Invalid input is important ERROR
        logging.error(f"{prefix} Input template_data is None or empty. Cannot create file.")
        return False

    wb = None
    try:
        wb = Workbook()
        ws = wb.active
        if ws is None:
            ws = wb.create_sheet(title=sheet_name, index=0)
            if ws is None:
                # Failure to get sheet is critical ERROR
                logging.error(f"{prefix} Could not get or create active worksheet in new workbook.")
                return False
        else:
            ws.title = sheet_name
        # Sheet creation/access confirmation is useful INFO
        logging.info(f"{prefix} Created/accessed sheet '{ws.title}' in new workbook.")

        try:
            start_col_str, start_row_str = coordinate_from_string(start_cell_coord)
            start_col_anchor = column_index_from_string(start_col_str)
            start_row_anchor = int(start_row_str)
            # Start coordinate is useful INFO
            logging.info(f"{prefix} Anchor/relative data placement starts at '{start_cell_coord}' (R={start_row_anchor}, C={start_col_anchor}).")
        except (ValueError, TypeError) as e:
            # Invalid start coord is an ERROR, but we default
            logging.error(f"{prefix} Invalid start_cell_coord '{start_cell_coord}': {e}. Using A1 as default.", exc_info=False)
            start_col_anchor, start_row_anchor = 1, 1

        # --- Helper: Apply single cell properties (Keep detailed logging as DEBUG) ---
        def _apply_props(target_ws: Worksheet, target_cell: Cell, cell_data_dict: Dict[str, Any], target_r: int, target_c: int):
             """Applies properties from dict to cell. Skips MergedCells."""
             # Skipping merged cells is detailed logic, keep as DEBUG
             if isinstance(target_cell, MergedCell):
                 logging.debug(f"{prefix} Skipping property application for MergedCell at {target_cell.coordinate}.")
                 return

             if not isinstance(cell_data_dict, dict):
                 # Invalid data format is a WARNING
                 logging.warning(f"{prefix} Invalid cell_data_dict type ({type(cell_data_dict)}) for cell {target_cell.coordinate}. Skipping.")
                 return

             if 'error' in cell_data_dict:
                 # Skipping due to prior error is a WARNING
                 logging.warning(f"{prefix} Skipping property application for cell {target_cell.coordinate} due to extraction error: {cell_data_dict['error']}")
                 return

             target_cell.value = cell_data_dict.get('value')
             height = cell_data_dict.get('height')
             if height is not None:
                 try:
                     height_f = float(height)
                     if height_f >= 0:
                         current_rd = target_ws.row_dimensions.get(target_r)
                         current_height = current_rd.height if current_rd and current_rd.height is not None else None
                         if current_height is None or abs(current_height - height_f) > 1e-6:
                             target_ws.row_dimensions[target_r].height = height_f
                     else: # Negative height is a WARNING
                         logging.warning(f"{prefix} Ignored negative height {height_f} for row {target_r}")
                 except (ValueError, TypeError): # Invalid height is a WARNING
                     logging.warning(f"{prefix} Invalid height value '{height}' for row {target_r}")

             width = cell_data_dict.get('width')
             if width is not None:
                 col_letter = get_column_letter(target_c)
                 try:
                     width_f = float(width)
                     if width_f >= 0:
                         current_cd = target_ws.column_dimensions.get(col_letter)
                         current_width = current_cd.width if current_cd and current_cd.width is not None else None
                         if current_width is None or abs(current_width - width_f) > 1e-6:
                            target_ws.column_dimensions[col_letter].width = width_f
                     else: # Negative width is a WARNING
                         logging.warning(f"{prefix} Ignored negative width {width_f} for col {col_letter}")
                 except (ValueError, TypeError): # Invalid width is a WARNING
                     logging.warning(f"{prefix} Invalid width value '{width}' for col {col_letter}")

             border_styles = cell_data_dict.get('border')
             if isinstance(border_styles, dict):
                 new_border = Border(
                    left=Side(style=border_styles.get('left')) if border_styles.get('left') else Side(),
                    right=Side(style=border_styles.get('right')) if border_styles.get('right') else Side(),
                    top=Side(style=border_styles.get('top')) if border_styles.get('top') else Side(),
                    bottom=Side(style=border_styles.get('bottom')) if border_styles.get('bottom') else Side()
                 )
                 current_border = target_cell.border if target_cell.has_style else Border()
                 if new_border != current_border:
                      target_cell.border = new_border
             elif target_cell.has_style and target_cell.border != Border():
                 target_cell.border = Border()

             alignment_styles = cell_data_dict.get('alignment')
             if isinstance(alignment_styles, dict):
                 new_alignment = Alignment(
                     horizontal=alignment_styles.get('horizontal'),
                     vertical=alignment_styles.get('vertical'),
                     wrap_text=alignment_styles.get('wrap_text')
                 )
                 current_alignment = target_cell.alignment if target_cell.has_style else Alignment()
                 if new_alignment != current_alignment:
                    if new_alignment != Alignment():
                        target_cell.alignment = new_alignment
                        # Applying alignment is detailed, keep as DEBUG
                        # logging.debug(f"{prefix} Applied alignment to {target_cell.coordinate}")
             elif target_cell.has_style and target_cell.alignment != Alignment():
                 target_cell.alignment = Alignment()
                 # Clearing alignment is detailed, keep as DEBUG
                 # logging.debug(f"{prefix} Cleared alignment from {target_cell.coordinate}")

        # --- Process template data ---
        processed_cells = 0
        max_target_row, max_target_col = 0, 0

        def sort_key_func(k):
            if k == ANCHOR_KEY: return (0, 0)
            parts = k.split('_')
            type_order = 99; index = 0
            if len(parts) == 2 and parts[1].isdigit():
                try:
                    index = int(parts[1])
                    if parts[0] == 'bottom': type_order = 1
                    elif parts[0] == 'right': type_order = 2
                except ValueError: pass
            return (type_order, index)
        valid_template_keys = [k for k in template_data.keys() if not k.startswith('_')]
        sorted_keys = sorted(valid_template_keys, key=sort_key_func)

        for key in sorted_keys:
            cell_value_or_list = template_data[key]

            # Skipping None is detailed, keep as DEBUG
            if cell_value_or_list is None:
                logging.debug(f"{prefix} Skipping key '{key}': Value is None.")
                continue

            if key == ANCHOR_KEY:
                if isinstance(cell_value_or_list, dict) and 'error' not in cell_value_or_list:
                    anchor_data = cell_value_or_list
                    target_row, target_col = start_row_anchor, start_col_anchor
                    target_cell = ws.cell(row=target_row, column=target_col)
                    # Processing anchor is detailed, keep as DEBUG
                    logging.debug(f"{prefix} Processing '{key}': Base Target Cell {target_cell.coordinate}")
                    _apply_props(ws, target_cell, anchor_data, target_row, target_col)
                    processed_cells += 1
                    max_target_row = max(max_target_row, target_row)
                    max_target_col = max(max_target_col, target_col)
                    merge_info = anchor_data.get('merge_info')
                    target_end_row, target_end_col = target_row, target_col
                    if isinstance(merge_info, dict):
                        rows_to_merge = merge_info.get('rows', 1)
                        cols_to_merge = merge_info.get('cols', 1)
                        if rows_to_merge > 1 or cols_to_merge > 1:
                            target_end_row = target_row + rows_to_merge - 1
                            target_end_col = target_col + cols_to_merge - 1
                            try:
                                merge_range_str = f"{target_cell.coordinate}:{get_column_letter(target_end_col)}{target_end_row}"
                                # Applying merge is significant, keep as INFO
                                logging.info(f"{prefix} Applying merge for anchor: {merge_range_str}")
                                ws.merge_cells(start_row=target_row, start_column=target_col,
                                               end_row=target_end_row, end_column=target_end_col)
                                max_target_row = max(max_target_row, target_end_row)
                                max_target_col = max(max_target_col, target_end_col)
                            except Exception as merge_err:
                                # Merge failure is important ERROR
                                logging.error(f"{prefix} Failed to apply merge for anchor ({merge_range_str}): {merge_err}", exc_info=True)
                                target_end_row, target_end_col = target_row, target_col
                else: # Invalid anchor data is a WARNING
                     logging.warning(f"{prefix} Skipping key '{key}': Invalid data format ({type(cell_value_or_list)}) or contains error.")
                continue

            elif key.startswith('bottom_'):
                 if isinstance(cell_value_or_list, dict) and 'error' not in cell_value_or_list:
                     cell_data_dict = cell_value_or_list
                     try:
                         offset = int(key.split('_')[1])
                         target_row = start_row_anchor + offset
                         target_col = start_col_anchor
                         if target_row <= 0 or target_col <= 0: continue
                         # Processing bottom cells is detailed, keep as DEBUG
                         target_coord_log = f"{get_column_letter(target_col)}{target_row}"
                         logging.debug(f"{prefix} Processing '{key}': Target {target_coord_log}")
                         target_cell = ws.cell(row=target_row, column=target_col)
                         _apply_props(ws, target_cell, cell_data_dict, target_row, target_col)
                         if not isinstance(target_cell, MergedCell): processed_cells += 1
                         max_target_row = max(max_target_row, target_row)
                         max_target_col = max(max_target_col, target_col)
                     except (ValueError, IndexError, TypeError) as e: # Calculation errors are WARNING
                         logging.warning(f"{prefix} Skipping '{key}': Error calculating position or applying props: {e}")
                 else: # Invalid bottom data is a WARNING
                     logging.warning(f"{prefix} Skipping key '{key}': Invalid data format ({type(cell_value_or_list)}) or contains error.")
                 continue

            elif key.startswith('right_'):
                 if isinstance(cell_value_or_list, list):
                     cell_data_list = cell_value_or_list
                     # Processing right lists is detailed, keep as DEBUG
                     logging.debug(f"{prefix} Processing '{key}' (List length={len(cell_data_list)})...")
                     try:
                         offset_col = int(key.split('_')[1])
                         target_base_col = start_col_anchor + offset_col
                         if target_base_col <= 0: continue
                         for row_offset, cell_data_item in enumerate(cell_data_list):
                             if cell_data_item is None: continue
                             if not isinstance(cell_data_item, dict) or 'error' in cell_data_item:
                                 # Skipping invalid items is a WARNING
                                 logging.warning(f"{prefix}   Skipping row offset {row_offset} for '{key}': Invalid item data ({cell_data_item}).")
                                 continue
                             target_row = start_row_anchor + row_offset
                             target_col = target_base_col
                             if target_row <= 0: continue
                             # Processing individual right cells is detailed, keep as DEBUG
                             cell_coord_log = f"{get_column_letter(target_col)}{target_row}"
                             logging.debug(f"{prefix}   Applying data for row offset {row_offset} to {cell_coord_log}")
                             target_cell = ws.cell(row=target_row, column=target_col)
                             _apply_props(ws, target_cell, cell_data_item, target_row, target_col)
                             if not isinstance(target_cell, MergedCell): processed_cells += 1
                             max_target_row = max(max_target_row, target_row)
                             max_target_col = max(max_target_col, target_col)
                     except (ValueError, IndexError, TypeError) as e: # Calculation errors are WARNING
                         logging.warning(f"{prefix} Skipping '{key}' list processing: Error calculating position or applying props: {e}")
                 else: # Expected list but got something else is a WARNING
                      logging.warning(f"{prefix} Skipping key '{key}': Expected a list for 'right_N', but got {type(cell_value_or_list)}.")
                 continue

            else: # Unrecognized keys are WARNING
                logging.warning(f"{prefix} Skipping unrecognized key '{key}' in template data.")

        if processed_cells > 0 or template_data.get(ANCHOR_KEY):
            # Saving is important INFO
            logging.info(f"{prefix} Processed relevant data (max R={max_target_row}, C={max_target_col}). Saving workbook to '{output_path}'...")
            wb.save(output_path)
            # Success saving is important INFO
            logging.info(f"{prefix} Successfully saved workbook to '{output_path}'.")
            return True
        else:
            # Saving an empty/merge-only file is a potential issue, keep as WARNING
            logging.warning(f"{prefix} No valid, non-merged cells processed from template data. Output file '{output_path}' might be empty or only contain merged cells.")
            try:
                wb.save(output_path)
                # Still log INFO that it was saved, even if potentially empty
                logging.info(f"{prefix} Saved workbook (potentially empty or only merges) to '{output_path}'.")
                return True
            except Exception as save_err:
                 # Failure to save is critical ERROR
                 logging.error(f"{prefix} Failed to save workbook to '{output_path}': {save_err}", exc_info=True)
                 return False

    except Exception as e:
        # Catch-all creation error is critical ERROR
        logging.error(f"{prefix} Failed to create or save workbook at '{output_path}': {e}", exc_info=True)
        return False
    finally:
         if wb:
              try: wb.close()
              # Error closing is minor, keep as WARNING
              except Exception as close_err: logging.warning(f"{prefix} Error closing workbook object: {close_err}")


# --- Example Usage ---
if __name__ == "__main__":
    # Set level to INFO for standard operation, DEBUG for detailed view
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(module)s - %(message)s') # Use INFO

    logging.info("--- Running Template Manager Example ---")
    logging.info(f"Current Working Directory: {os.getcwd()}")

    target_workbook = "dap.xlsx" # Assumes it's in the same directory

    if not os.path.exists(target_workbook):
        logging.critical(f"*** CRITICAL ERROR: Workbook not found at '{os.path.abspath(target_workbook)}'. Example cannot run. ***")
        exit(1)

    target_sheet = "Packing list"
    target_anchor = "Mark & Nº"
    cells_to_extract_below = 5
    cells_to_extract_right = 8
    template_key = "DapPackingList_v1" # Simplified key name
    should_create_copy = False # Typically False unless testing copy logic

    output_filename = "generated_template_output.xlsx"
    start_coordinate_in_output = "C3"

    target_wb_base, target_wb_ext = os.path.splitext(target_workbook)
    target_copy_file = f"{target_wb_base}{COPY_SUFFIX}{target_wb_ext}"

    logging.info("--- Pre-run Cleanup ---")
    files_to_clean = [target_copy_file, output_filename]
    for f_path in files_to_clean:
        abs_f_path = os.path.abspath(f_path)
        if os.path.exists(abs_f_path):
            try:
                os.remove(abs_f_path)
                # Removing files is useful INFO
                logging.info(f"Removed existing file: '{abs_f_path}'")
            except Exception as e:
                # Failure to remove is a WARNING
                logging.warning(f"Could not remove pre-existing file '{abs_f_path}': {e}")
        else:
             # File not found is detailed, keep as DEBUG
             logging.debug(f"Pre-existing file not found (no cleanup needed): '{abs_f_path}'")

    logging.info(f"\n>>> Attempting Load Template: Source='{target_workbook}', Sheet='{target_sheet}', Anchor='{target_anchor}'")
    load_success = load_template(
        workbook_path=target_workbook,
        sheet_identifier=target_sheet,
        template_name=template_key,
        anchor_text=target_anchor,
        num_below=cells_to_extract_below,
        num_right=cells_to_extract_right,
        force_reload=True,
        create_copy=should_create_copy
    )

    if load_success:
        logging.info(f"Successfully loaded template as '{template_key}'. Retrieving data...")
        retrieved_data = get_template(template_key)

        if retrieved_data is not None:
            # Debug printing of data structure only if DEBUG is enabled
            if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
                print(f"\n--- Extracted Data Sample ('{template_key}') ---")
                anchor_data_sample = retrieved_data.get(ANCHOR_KEY)
                print("Anchor Data:")
                pprint.pprint(anchor_data_sample, indent=2, width=100)
                example_right_key = None
                for i in range(1, cells_to_extract_right + 1):
                    key_to_check = f'right_{i}'
                    if key_to_check in retrieved_data: example_right_key = key_to_check; break
                if example_right_key:
                    right_data_sample = retrieved_data.get(example_right_key)
                    print(f"\nData for '{example_right_key}' (List Sample):")
                    if isinstance(right_data_sample, list):
                         pprint.pprint(right_data_sample[:5], indent=2, width=100)
                         if len(right_data_sample) > 5: print("  ...")
                    else: pprint.pprint(right_data_sample, indent=2, width=100)
                print("--- End Sample ---")

            logging.info(f"\n>>> Attempting to create XLSX from template data: Output='{output_filename}', Start='{start_coordinate_in_output}'")
            create_success = create_xlsx_from_template(
                template_data=retrieved_data,
                output_path=output_filename,
                start_cell_coord=start_coordinate_in_output,
                sheet_name=f"Generated_{target_sheet}"
            )

            if create_success:
                logging.info(f"Successfully created output file: '{output_filename}'")
                absolute_output_path = os.path.abspath(output_filename)
                print(f"\n+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                print(f"+++ Output XLSX file created successfully!")
                print(f"+++ Location: {absolute_output_path}")
                print(f"+++ Open the file to inspect the generated template structure.")
                print(f"+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            else:
                # Creation failure is important ERROR
                logging.error(f"Failed to create output file '{output_filename}'. Check logs for errors.")

        else:
            # Internal error if get_template fails after load success - critical ERROR
            logging.error(f"CRITICAL INTERNAL ERROR: Failed to retrieve template data ('{template_key}') after load reported success.")
    else:
        # Load failure is important ERROR
        logging.error(f"Failed to load template '{template_key}' from '{target_workbook}'. Cannot create output file.")

    # Final cleanup check logging reduced to DEBUG
    logging.debug("--- Post-run Cleanup Check ---")
    abs_target_copy_file = os.path.abspath(target_copy_file)
    logging.debug(f"Checking for cleanup target: '{abs_target_copy_file}'")
    if should_create_copy and os.path.exists(abs_target_copy_file):
        try:
            os.remove(abs_target_copy_file)
            # Post-run cleanup is INFO if it happens
            logging.info(f"Cleaned up copy file '{abs_target_copy_file}' after run (final check).")
        except Exception as e:
            # Failure is WARNING
            logging.warning(f"Could not remove copy file '{abs_target_copy_file}' during final cleanup check: {e}")
    else:
        if should_create_copy:
            # Copy not found is DEBUG
            logging.debug(f"Copy file '{abs_target_copy_file}' not found during final check (already cleaned up).")
        else:
            # No copy expected is DEBUG
            logging.debug(f"No copy file expected or found at '{abs_target_copy_file}' (create_copy=False).")

    logging.info("\n--- Template Manager Example Finished ---")