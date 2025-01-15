import sys
from loguru import logger

def setup_logging(verbose: bool = False):
    """
    Configure logging with file and console outputs
    
    Args:
        verbose: If True, sets DEBUG level for console output
    """
    # Remove default handler
    logger.remove()
    
    # File logger - no colors
    logger.add(
        "migration.log",
        rotation="1 day",
        retention="7 days",
        level="DEBUG",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {message}",
        colorize=False,
        enqueue=True
    )
    
    # Console logger with colors
    level = "DEBUG" if verbose else "INFO"
    logger.add(
        sys.stdout,
        level=level,
        format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{message}</cyan>",
        colorize=True,
        enqueue=True
    )

def get_logger():
    """
    Get configured logger instance
    """
    return logger