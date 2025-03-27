# --- START OF FILE excel_handler.py ---

import openpyxl
import os
import logging # Using logging is better than print for info/errors

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ExcelHandler:
    """Handles loading and accessing data from Excel files using openpyxl."""
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            raise FileNotFoundError(f"The file '{file_path}' was not found.")
        self.file_path = file_path
        self.workbook = None
        self.sheet = None
        logging.info(f"Initialized ExcelHandler for: {file_path}")

    def load_sheet(self, sheet_name=None, data_only=True):
        """
        Loads the workbook and a specific sheet.

        Args:
            sheet_name (str, optional): Name of the sheet. Defaults to None (active sheet).
            data_only (bool, optional): Get cell values (True) or formulas (False). Defaults to True.

        Returns:
            openpyxl.worksheet.worksheet.Worksheet: The loaded sheet object, or None on failure.
        """
        try:
            self.workbook = openpyxl.load_workbook(self.file_path, data_only=data_only)
            if sheet_name:
                if sheet_name in self.workbook.sheetnames:
                    self.sheet = self.workbook[sheet_name]
                else:
                    logging.warning(f"Sheet '{sheet_name}' not found in '{self.file_path}'. Loading active sheet.")
                    self.sheet = self.workbook.active
            else:
                self.sheet = self.workbook.active

            logging.info(f"Successfully loaded sheet: '{self.sheet.title}'")
            return self.sheet
        except Exception as e:
            logging.error(f"Failed to load workbook/sheet from '{self.file_path}': {e}", exc_info=True)
            self.workbook = None
            self.sheet = None
            return None

    def get_sheet(self):
        """Returns the currently loaded sheet object."""
        if not self.sheet:
            logging.warning("Sheet not loaded. Call load_sheet() first.")
        return self.sheet

    def close(self):
        """Clears references to the workbook and sheet."""
        # openpyxl doesn't require explicit file closing like open(),
        # but clearing references helps garbage collection.
        if self.workbook:
            self.workbook = None
            self.sheet = None
            logging.info(f"Closed workbook reference for: {self.file_path}")

# --- END OF FILE excel_handler.py ---