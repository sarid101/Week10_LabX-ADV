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
