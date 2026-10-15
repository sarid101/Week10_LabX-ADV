# AllCode — Week 10 Agentic Spreadsheet Processor
รวมซอร์สโค้ดทั้งหมดของโปรเจกต์ (สาริษฐ์ บุตรช่วง 663380365-3)

## `main.py`

```python
# agentic-spreadsheet-processor/main.py
import os

# ไฟล์ใน src ใช้ relative import (from .utils ...) จึงต้อง import ผ่านแพ็กเกจ src
from src.utils import setup_logging
from src.config_parser import ConfigParser
from src.spreadsheet_agent import SpreadsheetAgent

logger = setup_logging(__name__)


def main():
    """
    Main entry point for the Agentic Spreadsheet Processor.
    Loads config and runs the agent.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(base_dir, 'configs', 'example_spreadsheet_config.json')

    try:
        os.makedirs(os.path.join(base_dir, 'data'), exist_ok=True)

        logger.info(f"Loading configuration from: {config_file_path}")
        config_parser = ConfigParser(config_file_path)
        config = config_parser.load_config()

        # paths in the config are relative to the project root -> run from the project root
        os.chdir(base_dir)

        if not os.path.exists(config["input_file"]):
            logger.critical(f"Input Excel file not found: '{config['input_file']}'. Run create_input_data.py first.")
            return

        logger.info("Initializing Spreadsheet Agent...")
        agent = SpreadsheetAgent(config)

        logger.info("Running Spreadsheet Agent...")
        agent.run()

    except FileNotFoundError as e:
        logger.critical(f"File system error: {e}")
        logger.info("Please ensure all specified input files and configurations exist.")
    except ValueError as e:
        logger.critical(f"Configuration error: {e}")
        logger.info("Please check the format and required fields in your config file.")
    except Exception as e:
        logger.critical(f"An unexpected critical error occurred during execution: {e}", exc_info=True)


if __name__ == "__main__":
    main()
```

## `create_input_data.py`

```python
# agentic-spreadsheet-processor/create_input_data.py
# สร้างไฟล์ data/input_data.xlsx (ชีต SalesData) ตามตัวอย่างในใบงาน
import os
from openpyxl import Workbook
from openpyxl.styles import Font

ROWS = [
    ("Product Name", "Quantity", "Unit Price"),
    ("Laptop", 2, 1200.50),
    ("Mouse", 5, 25.00),
    ("Keyboard", 3, 75.99),
    ("Monitor", 1, 300.00),
    ("Projector", 1, 850.00),
    ("Webcam", 10, 45.00),
    ("Headset", 4, 99.99),
    ("SSD", 2, 180.00),
    ("Router", 1, 70.00),
]


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base_dir, "data", "input_data.xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "SalesData"
    for row in ROWS:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 12

    wb.save(path)
    print(f"Created {path}")


if __name__ == "__main__":
    main()
```

## `src/__init__.py`

```python
# agentic-spreadsheet-processor/src/__init__.py
# This file indicates that the 'src' directory is a Python package.
```

## `src/utils.py`

```python
# agentic-spreadsheet-processor/src/utils.py
import logging
import json
import os

# Configure basic logging for the utility module
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def setup_logging(name, level=logging.INFO, filename=None):
    """Sets up a logger for specific modules."""
    logger = logging.getLogger(name)
    logger.setLevel(level)

    if not logger.handlers:  # Prevent adding multiple handlers if called multiple times
        # Console handler
        ch = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        logger.propagate = False  # avoid printing each line twice via the root logger

        # File handler (optional)
        if filename:
            fh = logging.FileHandler(filename)
            fh.setFormatter(formatter)
            logger.addHandler(fh)
    return logger


def save_json_report(data, filename, directory='data'):
    """Saves a dictionary or list to a JSON file."""
    os.makedirs(directory, exist_ok=True)
    filepath = os.path.join(directory, filename)
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        logging.info(f"Report saved to {filepath}")
        return True
    except IOError as e:
        logging.error(f"Error saving report to JSON file: {e}")
        return False
```

## `src/config_parser.py`

```python
# agentic-spreadsheet-processor/src/config_parser.py
import json
import os
from .utils import setup_logging

logger = setup_logging(__name__)


class ConfigParser:
    """
    Parses and validates the spreadsheet automation configuration from a JSON file.
    """
    def __init__(self, config_path):
        self.config_path = config_path
        self.config = {}

    def load_config(self):
        """Loads the configuration from the specified JSON file."""
        if not os.path.exists(self.config_path):
            logger.critical(f"Configuration file not found: {self.config_path}")
            raise FileNotFoundError(f"Configuration file not found: {self.config_path}")

        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            self._validate_config()
            logger.info(f"Configuration loaded successfully from {self.config_path}")
            return self.config
        except json.JSONDecodeError as e:
            logger.critical(f"Invalid JSON format in config file: {e}")
            raise ValueError(f"Invalid JSON format in config file: {e}")
        except ValueError:
            raise  # keep validation errors as ValueError so main.py reports them as config errors
        except Exception as e:
            logger.critical(f"An error occurred while loading config: {e}")
            raise Exception(f"An error occurred while loading config: {e}")

    def _validate_config(self):
        """
        Validates the loaded configuration to ensure required fields and task structures are present.
        This validation can be extended based on all possible task types.
        """
        required_top_level_fields = ["input_file", "output_file", "tasks"]
        for field in required_top_level_fields:
            if field not in self.config:
                raise ValueError(f"Missing required field in config: '{field}'")

        if not isinstance(self.config["tasks"], list):
            raise ValueError("Config field 'tasks' must be a list.")

        # Basic validation for each task
        for i, task in enumerate(self.config["tasks"]):
            if "type" not in task:
                raise ValueError(f"Task {i} is missing 'type' field.")

            # copy_data ใช้ source_sheet / dest_sheet แทน sheet จึงตรวจแยก
            if task["type"] == "copy_data":
                if "source_sheet" not in task or "dest_sheet" not in task:
                    raise ValueError(f"Copy task {i} requires 'source_sheet' and 'dest_sheet'.")
            elif "sheet" not in task:
                raise ValueError(f"Task {i} is missing 'sheet' field.")

            if task["type"] == "chart":
                if "chart_type" not in task or "data_range" not in task:
                    raise ValueError(f"Chart task {i} requires 'chart_type' and 'data_range'.")

            if task["type"] == "conditional_format":
                if "range" not in task or "rule_type" not in task:
                    raise ValueError(f"Conditional format task {i} requires 'range' and 'rule_type'.")

        logger.info("Configuration validated.")
```

## `src/google_sheets_utils.py`

```python
# agentic-spreadsheet-processor/src/google_sheets_utils.py
# This file is a placeholder for future integration with Google Sheets API.
# Full implementation requires Google Cloud Project setup, API key, and OAuth.

import logging

logger = logging.getLogger(__name__)


class GoogleSheetsUtils:
    """
    Conceptual class for interacting with Google Sheets.
    Requires 'gspread' and 'google-auth' libraries.
    """
    def __init__(self):
        logger.info("GoogleSheetsUtils initialized (conceptual).")
        # In a real implementation, you would set up authentication here.
        # e.g., self.gc = gspread.service_account(filename='path/to/your/credentials.json')

    def read_sheet(self, spreadsheet_id, range_name):
        """Conceptual method to read data from a Google Sheet."""
        logger.info(f"Attempting to read range '{range_name}' from Google Sheet ID: {spreadsheet_id}")
        # Placeholder: In reality, use self.gc.open_by_key(spreadsheet_id).worksheet('Sheet1').get(range_name)
        return [["Header1", "Header2"], ["Value1", "Value2"]]  # Dummy data

    def write_sheet(self, spreadsheet_id, sheet_name, data, start_cell="A1"):
        """Conceptual method to write data to a Google Sheet."""
        logger.info(f"Attempting to write data to Google Sheet ID: {spreadsheet_id}, sheet '{sheet_name}' from {start_cell}")
        # Placeholder: In reality, use self.gc.open_by_key(spreadsheet_id).worksheet(sheet_name).update(start_cell, data)
        return True

    def create_sheet(self, spreadsheet_id, sheet_name):
        """Conceptual method to create a new sheet in a Google Sheet."""
        logger.info(f"Attempting to create sheet '{sheet_name}' in Google Sheet ID: {spreadsheet_id}")
        # Placeholder: In reality, use self.gc.open_by_key(spreadsheet_id).add_worksheet(sheet_name, rows=100, cols=20)
        return True
```

## `src/spreadsheet_agent.py`

```python
# agentic-spreadsheet-processor/src/spreadsheet_agent.py
import json
import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter, range_boundaries

from .utils import setup_logging  # Use utils for consistent logging setup

logger = setup_logging(__name__)


class SpreadsheetAgent:
    """
    An agent that performs a series of operations on an Excel workbook
    based on a provided configuration.
    """
    def __init__(self, config):
        self.config = config
        self.workbook = None
        self.audit_log = []
        self.start_time = datetime.datetime.now()
        logger.info("Spreadsheet Agent initialized.")

    def _log_audit(self, status, message, task_type="N/A", details=None):
        """Adds an entry to the audit log."""
        log_entry = {
            "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
            "status": status,
            "task_type": task_type,
            "message": message,
            "details": details if details else {}
        }
        self.audit_log.append(log_entry)
        if status == "ERROR" or status == "CRITICAL":
            logger.error(f"AUDIT LOG - {status}: {message} | Task: {task_type} | Details: {details}")
        else:
            logger.info(f"AUDIT LOG - {status}: {message} | Task: {task_type}")

    def _load_workbook(self):
        """Loads the input workbook."""
        input_file = self.config["input_file"]
        try:
            self.workbook = openpyxl.load_workbook(input_file)
            self._log_audit("SUCCESS", f"Workbook '{input_file}' loaded.", "load_workbook")
            logger.info(f"Loaded workbook: {input_file}")
            return True
        except FileNotFoundError:
            self._log_audit("ERROR", f"Input file not found: {input_file}", "load_workbook")
            logger.critical(f"Input file not found: {input_file}")
            return False
        except Exception as e:
            self._log_audit("ERROR", f"Failed to load workbook '{input_file}': {e}", "load_workbook")
            logger.critical(f"Error loading workbook: {e}")
            return False

    def _get_sheet(self, sheet_name, create=True):
        """Safely gets a worksheet from the workbook (creates it when allowed)."""
        if sheet_name in self.workbook.sheetnames:
            return self.workbook[sheet_name]
        if not create:
            return None
        self._log_audit("WARNING", f"Sheet '{sheet_name}' not found. Creating new sheet.", "get_sheet")
        logger.warning(f"Sheet '{sheet_name}' not found. Creating new sheet.")
        return self.workbook.create_sheet(sheet_name)

    def _process_copy_data_task(self, task):
        """Copies data from a source sheet to a destination sheet."""
        src_sheet_name = task.get("source_sheet")
        dest_sheet_name = task.get("dest_sheet")

        if not src_sheet_name or not dest_sheet_name:
            self._log_audit("ERROR", "Missing source_sheet or dest_sheet for copy_data task.", "copy_data", task)
            return

        # source sheet must already exist — don't silently create an empty one
        src_sheet = self._get_sheet(src_sheet_name, create=False)
        if src_sheet is None:
            self._log_audit("ERROR", f"Source sheet '{src_sheet_name}' not available for copying.", "copy_data", task)
            return

        dest_sheet = self._get_sheet(dest_sheet_name)

        for row in src_sheet.iter_rows():
            for cell in row:
                dest_sheet[cell.coordinate].value = cell.value
        # keep the header bold and the column widths readable
        for cell in dest_sheet[1]:
            cell.font = Font(bold=True)
        for col_letter, dim in src_sheet.column_dimensions.items():
            dest_sheet.column_dimensions[col_letter].width = dim.width

        self._log_audit("SUCCESS", f"Data copied from '{src_sheet_name}' to '{dest_sheet_name}'.", "copy_data")
        logger.info(f"Copied data from {src_sheet_name} to {dest_sheet_name}")

    def _process_calculate_column_task(self, task):
        """Calculates a new column based on a formula or existing columns."""
        sheet_name = task.get("sheet")
        target_column_letter = task.get("target_column")
        start_row = task.get("start_row", 2)  # Assume header in row 1
        formula = task.get("formula")  # e.g., "=B{row}*C{row}"

        if not sheet_name or not target_column_letter or not formula:
            self._log_audit("ERROR", "Missing required fields for calculate_column task.", "calculate_column", task)
            return

        sheet = self._get_sheet(sheet_name)

        # Add header for the new column
        header_name = task.get("header", "Calculated Column")
        sheet[f"{target_column_letter}1"].value = header_name
        sheet[f"{target_column_letter}1"].font = Font(bold=True)
        sheet.column_dimensions[target_column_letter].width = max(14, len(header_name) + 4)

        for row_num in range(start_row, sheet.max_row + 1):
            cell_coord = f"{target_column_letter}{row_num}"
            try:
                sheet[cell_coord] = formula.replace("{row}", str(row_num))
                if task.get("number_format"):
                    sheet[cell_coord].number_format = task["number_format"]
            except Exception as e:
                self._log_audit("ERROR", f"Failed to apply formula to {cell_coord}: {e}", "calculate_column", task)
                logger.error(f"Failed to apply formula to {cell_coord}: {e}")
        self._log_audit("SUCCESS", f"Column '{target_column_letter}' calculated in '{sheet_name}'.", "calculate_column")
        logger.info(f"Calculated column {target_column_letter} in {sheet_name}")

    def _process_conditional_format_task(self, task):
        """Applies conditional formatting to a range."""
        sheet_name = task.get("sheet")
        data_range = task.get("range")
        rule_type = task.get("rule_type")

        if not sheet_name or not data_range or not rule_type:
            self._log_audit("ERROR", "Missing required fields for conditional_format task.", "conditional_format", task)
            return

        sheet = self._get_sheet(sheet_name)

        try:
            if rule_type == "color_scale":
                # 3-color scale: red (min) -> yellow (50th percentile) -> green (max)
                rule = ColorScaleRule(start_type='min', start_color=task.get("min_color", "FF0000"),
                                      mid_type='percentile', mid_value=50, mid_color=task.get("mid_color", "FFFF00"),
                                      end_type='max', end_color=task.get("max_color", "00FF00"))
                sheet.conditional_formatting.add(data_range, rule)
            elif rule_type == "data_bar":
                rule = DataBarRule(start_type='min', end_type='max', color=task.get("color", "0070C0"))
                sheet.conditional_formatting.add(data_range, rule)
            elif rule_type == "icon_set":
                rule = IconSetRule(task.get("icon_set", '3TrafficLights1'),
                                   task.get("types", 'percent'),
                                   task.get("values", [0, 33, 67]))
                sheet.conditional_formatting.add(data_range, rule)
            elif rule_type == "expression":
                formula = task.get("formula", "")
                if formula.startswith("="):
                    formula = formula[1:]  # conditional-format formulas are stored without the leading '='
                fill_color = task.get("fill_color", "FFC7CE")  # Light Red
                font_color = task.get("font_color", "9C0006")  # Dark Red
                dxf = DifferentialStyle(font=Font(color=font_color),
                                        fill=PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid"))
                rule = Rule(type="expression", dxf=dxf, formula=[formula])
                sheet.conditional_formatting.add(data_range, rule)
            else:
                self._log_audit("ERROR", f"Unsupported conditional format rule_type: {rule_type}", "conditional_format", task)
                return

            self._log_audit("SUCCESS", f"Conditional formatting '{rule_type}' applied to '{data_range}' in '{sheet_name}'.", "conditional_format")
            logger.info(f"Applied conditional format {rule_type} to {data_range} in {sheet_name}")

        except Exception as e:
            self._log_audit("ERROR", f"Failed to apply conditional formatting to '{data_range}': {e}", "conditional_format", task)
            logger.error(f"Failed to apply conditional formatting to {data_range}: {e}")

    def _process_chart_task(self, task):
        """Creates a chart based on specified data."""
        sheet_name = task.get("sheet")
        chart_type = task.get("chart_type")
        data_range = task.get("data_range")          # e.g., "D1:D10" (header + values)
        category_range = task.get("category_range")  # e.g., "A2:A10" (labels, no header)
        title = task.get("title", "Chart Title")
        x_axis_title = task.get("x_axis_title", "")
        y_axis_title = task.get("y_axis_title", "")
        top_left_cell = task.get("top_left_cell", "G2")

        if not sheet_name or not chart_type or not data_range:
            self._log_audit("ERROR", "Missing required fields for chart task.", "chart", task)
            return

        sheet = self._get_sheet(sheet_name)

        try:
            if chart_type == "bar":
                chart = BarChart()
            elif chart_type == "line":
                chart = LineChart()
            elif chart_type == "pie":
                chart = PieChart()
            else:
                self._log_audit("ERROR", f"Unsupported chart type: {chart_type}", "chart", task)
                return

            min_col, min_row, max_col, max_row = range_boundaries(data_range)
            data = Reference(sheet, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
            chart.series.append(Series(data, title_from_data=True))  # first row = series title

            if category_range:
                c_min_col, c_min_row, c_max_col, c_max_row = range_boundaries(category_range)
                cats = Reference(sheet, min_col=c_min_col, min_row=c_min_row, max_col=c_max_col, max_row=c_max_row)
                chart.set_categories(cats)

            chart.title = title
            if chart_type != "pie":
                chart.x_axis.title = x_axis_title
                chart.y_axis.title = y_axis_title
                chart.x_axis.delete = False  # show axis labels in newer Excel
                chart.y_axis.delete = False
            chart.width, chart.height = task.get("width", 18), task.get("height", 9)

            sheet.add_chart(chart, top_left_cell)
            self._log_audit("SUCCESS", f"Chart '{title}' of type '{chart_type}' added to '{sheet_name}'.", "chart")
            logger.info(f"Added {chart_type} chart to {sheet_name}")

        except Exception as e:
            self._log_audit("ERROR", f"Failed to create chart '{title}': {e}", "chart", task)
            logger.error(f"Failed to create chart: {e}")

    def _add_audit_log_sheet(self):
        """Adds a sheet to the workbook containing the audit log."""
        if not self.audit_log:
            return

        audit_sheet = self.workbook.create_sheet("Audit Log", 0)  # Create as first sheet
        audit_sheet.append(["Timestamp", "Status", "Task Type", "Message", "Details"])

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
        for col_num in range(1, 6):
            cell = audit_sheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill

        status_colors = {"SUCCESS": "C6EFCE", "WARNING": "FFEB9C", "ERROR": "FFC7CE", "CRITICAL": "FFC7CE"}
        for entry in self.audit_log:
            audit_sheet.append([
                entry["timestamp"],
                entry["status"],
                entry["task_type"],
                entry["message"],
                json.dumps(entry["details"], ensure_ascii=False)  # Store details as JSON string
            ])
            color = status_colors.get(entry["status"])
            if color:
                audit_sheet.cell(row=audit_sheet.max_row, column=2).fill = PatternFill(
                    start_color=color, end_color=color, fill_type="solid")

        # Auto-adjust column widths (capped so the Details column stays readable)
        for col_num in range(1, 6):
            column_letter = get_column_letter(col_num)
            max_length = max((len(str(c.value)) for c in audit_sheet[column_letter] if c.value), default=0)
            audit_sheet.column_dimensions[column_letter].width = min(max_length + 2, 60)
        audit_sheet.freeze_panes = "A2"

        logger.info("Audit log sheet created.")

    def run(self):
        """Executes the defined tasks from the configuration."""
        if not self._load_workbook():
            self._log_audit("CRITICAL", "Agent could not start due to workbook loading failure.", "run")
            return

        try:
            for i, task in enumerate(self.config["tasks"]):
                task_type = task.get("type")
                logger.info(f"Executing Task {i + 1}: Type='{task_type}' Sheet='{task.get('sheet', task.get('dest_sheet'))}'")

                if task_type == "copy_data":
                    self._process_copy_data_task(task)
                elif task_type == "calculate_column":
                    self._process_calculate_column_task(task)
                elif task_type == "conditional_format":
                    self._process_conditional_format_task(task)
                elif task_type == "chart":
                    self._process_chart_task(task)
                else:
                    self._log_audit("WARNING", f"Unsupported task type encountered: {task_type}", "task_execution", task)
                    logger.warning(f"Unsupported task type: {task_type}")

            # Summary + save are logged before the audit sheet is written so they appear in it
            output_file = self.config["output_file"]
            duration = (datetime.datetime.now() - self.start_time).total_seconds()
            self._log_audit("INFO", f"All {len(self.config['tasks'])} tasks processed in {duration:.2f} seconds.", "run_summary")
            self._log_audit("INFO", f"Saving workbook to '{output_file}'.", "save_workbook")
            self._add_audit_log_sheet()

            try:
                self.workbook.save(output_file)
                logger.info(f"Final workbook saved to: {output_file}")
            except Exception as e:
                self._log_audit("CRITICAL", f"Failed to save output workbook '{output_file}': {e}", "save_workbook")
                logger.critical(f"Error saving output workbook: {e}")

        except Exception as e:
            self._log_audit("CRITICAL", f"An unhandled error occurred during agent execution: {e}", "run")
            logger.critical(f"An unhandled error occurred: {e}", exc_info=True)
        finally:
            duration = (datetime.datetime.now() - self.start_time).total_seconds()
            logger.info(f"Agent finished. Total time: {duration:.2f} seconds.")
```

## `configs/example_spreadsheet_config.json`

```json
{
    "input_file": "data/input_data.xlsx",
    "output_file": "data/output_report.xlsx",
    "tasks": [
        {
            "type": "copy_data",
            "source_sheet": "SalesData",
            "dest_sheet": "Raw Data"
        },
        {
            "type": "calculate_column",
            "sheet": "Raw Data",
            "target_column": "D",
            "header": "Total Revenue",
            "start_row": 2,
            "formula": "=B{row}*C{row}",
            "number_format": "#,##0.00"
        },
        {
            "type": "conditional_format",
            "sheet": "Raw Data",
            "range": "D2:D100",
            "rule_type": "color_scale",
            "min_color": "FF0000",
            "mid_color": "FFFF00",
            "max_color": "00FF00"
        },
        {
            "type": "chart",
            "sheet": "Raw Data",
            "chart_type": "bar",
            "data_range": "D1:D5",
            "category_range": "A2:A5",
            "title": "Top 4 Product Revenue",
            "x_axis_title": "Product",
            "y_axis_title": "Revenue",
            "top_left_cell": "G2"
        },
        {
            "type": "calculate_column",
            "sheet": "Raw Data",
            "target_column": "E",
            "header": "Status",
            "start_row": 2,
            "formula": "=IF(D{row}>1000,\"High Value\",\"Standard\")"
        },
        {
            "type": "conditional_format",
            "sheet": "Raw Data",
            "range": "E2:E100",
            "rule_type": "expression",
            "formula": "=$E2=\"High Value\"",
            "fill_color": "92D050",
            "font_color": "000000"
        }
    ]
}
```
