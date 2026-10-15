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
