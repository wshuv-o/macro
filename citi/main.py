"""
run_excel_macro.py

A comprehensive Python utility to automate running Excel macros on `.xlsm` files via COM automation.

Features:
- Runs specified macro from given workbook.
- Supports batch processing via CSV input.
- Robust error handling and retries.
- Configurable via JSON config file.
- CLI with argument parsing.
- Logging with timestamps and log file export.
- Extracts basic workbook info.
- Placeholder extension points for data preprocessing.
"""

import os
import sys
import time
import json
import argparse
import logging
import csv
from datetime import datetime
from typing import List, Optional

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 module not found. Please install it with 'pip install pywin32'.")
    sys.exit(1)


class ExcelMacroRunner:
    """
    Core class to manage Excel macro execution with retries and logging.
    """

    def __init__(self, excel_path: str, macro_name: str, visible: bool = False,
                 max_retries: int = 3, retry_delay_sec: int = 5):
        self.excel_path = excel_path
        self.macro_name = macro_name
        self.visible = visible
        self.max_retries = max_retries
        self.retry_delay_sec = retry_delay_sec
        self.excel_app = None
        self.workbook = None

        self.logger = logging.getLogger('ExcelMacroRunner')

    def open_excel(self):
        self.logger.debug("Starting Excel COM application.")
        self.excel_app = win32com.client.Dispatch("Excel.Application")
        self.excel_app.Visible = self.visible
        self.excel_app.DisplayAlerts = False
        self.logger.debug("Excel application started.")

    def close_excel(self):
        self.logger.debug("Closing Excel application.")
        try:
            if self.workbook:
                self.workbook.Close(False)
                self.logger.debug("Workbook closed.")
            if self.excel_app:
                self.excel_app.Quit()
                self.logger.debug("Excel quit called.")
            # Clean up COM references
            self.workbook = None
            self.excel_app = None
        except Exception as e:
            self.logger.warning(f"Exception during Excel close: {e}")

    def open_workbook(self):
        self.logger.debug(f"Opening workbook at path: {self.excel_path}")
        self.workbook = self.excel_app.Workbooks.Open(self.excel_path)
        self.logger.info(f"Workbook '{os.path.basename(self.excel_path)}' opened.")

    def save_workbook(self):
        if self.workbook:
            self.logger.debug("Saving workbook.")
            self.workbook.Save()
            self.logger.info("Workbook saved.")

    def run_macro(self):
        self.logger.info(f"Running macro '{self.macro_name}'.")
        self.excel_app.Run(self.macro_name)
        self.logger.info("Macro execution completed.")

    def extract_workbook_info(self):
        """
        Extracts some info from the workbook to demonstrate COM access.
        """
        try:
            sheets = [sheet.Name for sheet in self.workbook.Sheets]
            self.logger.info(f"Workbook sheets: {sheets}")
            # Extract used range address of first sheet
            first_sheet = self.workbook.Sheets(1)
            used_range = first_sheet.UsedRange.Address
            self.logger.info(f"First sheet '{first_sheet.Name}' used range: {used_range}")
        except Exception as e:
            self.logger.warning(f"Failed to extract workbook info: {e}")

    def execute(self):
        attempt = 0
        while attempt < self.max_retries:
            try:
                self.open_excel()
                self.open_workbook()
                self.extract_workbook_info()
                self.run_macro()
                self.save_workbook()
                self.close_excel()
                self.logger.info(f"Macro run successfully on attempt {attempt + 1}.")
                return True
            except Exception as e:
                self.logger.error(f"Error on attempt {attempt + 1}: {e}", exc_info=True)
                self.close_excel()
                attempt += 1
                if attempt < self.max_retries:
                    self.logger.info(f"Retrying in {self.retry_delay_sec} seconds...")
                    time.sleep(self.retry_delay_sec)
                else:
                    self.logger.error(f"All {self.max_retries} attempts failed.")
                    return False


def setup_logging(log_file: Optional[str] = None, verbose: bool = False):
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    formatter = logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    # File handler
    if log_file:
        fh = logging.FileHandler(log_file)
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)


def load_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file '{config_path}' not found.")
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config


def batch_run_from_csv(csv_path: str, macro_name: str, visible: bool, retries: int, delay: int):
    """
    Run macro on multiple files listed in a CSV file with column header: filepath
    """
    logger = logging.getLogger('batch_run_from_csv')

    if not os.path.exists(csv_path):
        logger.error(f"CSV file '{csv_path}' does not exist.")
        return

    success_count = 0
    fail_count = 0

    with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for idx, row in enumerate(reader, 1):
            file_path = row.get('filepath')
            if not file_path:
                logger.warning(f"Row {idx} missing 'filepath' column, skipping.")
                continue
            logger.info(f"Processing row {idx}: {file_path}")

            runner = ExcelMacroRunner(file_path, macro_name, visible=visible, max_retries=retries, retry_delay_sec=delay)
            success = runner.execute()
            if success:
                success_count += 1
            else:
                fail_count += 1

    logger.info(f"Batch run complete. Success: {success_count}, Fail: {fail_count}")


def preprocess_data_placeholder():
    """
    Placeholder for data preprocessing before running macro.
    Extend as needed.
    """
    logger = logging.getLogger('preprocess_data_placeholder')
    logger.info("Preprocessing data... (placeholder)")


def main():
    parser = argparse.ArgumentParser(description="Run Excel macro automation script")
    parser.add_argument('-f', '--file', type=str, help="Path to the Excel (.xlsm) file to process.")
    parser.add_argument('-m', '--macro', type=str, required=True, help="Macro name to run (e.g. Module1.MyMacro).")
    parser.add_argument('-c', '--config', type=str, help="Path to JSON config file for batch mode.")
    parser.add_argument('-b', '--batch', type=str, help="CSV file with list of Excel files to run macro on.")
    parser.add_argument('-v', '--visible', action='store_true', help="Show Excel UI during execution.")
    parser.add_argument('-r', '--retries', type=int, default=3, help="Number of retries on failure.")
    parser.add_argument('-d', '--delay', type=int, default=5, help="Delay (seconds) between retries.")
    parser.add_argument('-l', '--log', type=str, default='macro_runner.log', help="Log file path.")
    parser.add_argument('--verbose', action='store_true', help="Enable verbose debug logging.")

    args = parser.parse_args()

    setup_logging(args.log, args.verbose)
    logger = logging.getLogger('main')
    logger.info("Starting Excel Macro Runner.")

    preprocess_data_placeholder()

    if args.config:
        logger.info(f"Loading configuration from {args.config}")
        try:
            config = load_config(args.config)
            logger.debug(f"Config loaded: {config}")
            # Optionally override CLI args with config values
            file_to_run = config.get('file', args.file)
            macro_name = config.get('macro', args.macro)
            batch_file = config.get('batch', args.batch)
            visible = config.get('visible', args.visible)
            retries = config.get('retries', args.retries)
            delay = config.get('delay', args.delay)
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
            sys.exit(1)
    else:
        file_to_run = args.file
        macro_name = args.macro
        batch_file = args.batch
        visible = args.visible
        retries = args.retries
        delay = args.delay

    if batch_file:
        logger.info(f"Starting batch macro run from CSV: {batch_file}")
        batch_run_from_csv(batch_file, macro_name, visible, retries, delay)
    else:
        if not file_to_run:
            logger.error("No Excel file provided for macro run.")
            sys.exit(1)

        runner = ExcelMacroRunner(file_to_run, macro_name, visible=visible, max_retries=retries, retry_delay_sec=delay)
        success = runner.execute()
        if success:
            logger.info("Macro executed successfully.")
            sys.exit(0)
        else:
            logger.error("Macro execution failed.")
            sys.exit(2)


if __name__ == "__main__":
    main()
