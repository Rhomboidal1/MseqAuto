# logger.py
import logging
from pathlib import Path
from datetime import datetime

def setup_logger(name, log_dir=None):
    # Configure logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Create logs directory if it doesn't exist
    if log_dir is None:
        # Default to a logs directory in the same directory as this module
        log_dir = Path(__file__).resolve().parent / "logs"
    else:
        log_dir = Path(log_dir)
    
    if not log_dir.exists():
        log_dir.mkdir(parents=True)

    # File handler - logs to file
    log_file = log_dir / f"{name}_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)  # Make sure this is DEBUG, not INFO
    
    # Console handler - logs to console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)  # Can keep console at INFO to reduce spam
    
    # Format
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger