# test_ui_automation_logging.py
import os
import sys
from pathlib import Path

# Add parent directory to PYTHONPATH for imports
sys.path.append(str(Path(__file__).parents[2]))

from mseqauto.config import MseqConfig
from mseqauto.core.ui_automation import MseqAutomation

def test_logging():
    """Test the logging setup in MseqAutomation"""
    print("Testing ui_automation logging setup...")

    # Initialize config
    config = MseqConfig()

    # Initialize MseqAutomation without providing a logger (should create its own)
    automation = MseqAutomation(config)

    # Test some log messages
    automation.logger.info("Testing INFO level logging")
    automation.logger.debug("Testing DEBUG level logging")
    automation.logger.warning("Testing WARNING level logging")
    automation.logger.error("Testing ERROR level logging")

    print("Log test complete. Check scripts/logs/ for ui_automation_YYYYMMDD.log")

if __name__ == "__main__":
    test_logging()
