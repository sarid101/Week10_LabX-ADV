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
