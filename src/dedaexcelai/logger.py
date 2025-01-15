import sys
from loguru import logger

def setup_logging(verbose: bool = False):
    """Configure logging with colored output"""
    # Remove default handler
    logger.remove()
    
    # Set level
    level = "DEBUG" if verbose else "INFO"
    
    # Add handler with simple format
    # Map log levels to emojis
    emojis = {
        "DEBUG": "🔍",
        "INFO": "ℹ️ ",
        "WARNING": "⚠️ ",
        "ERROR": "❌",
        "CRITICAL": "💥"
    }

    def format_record(record):
        # Get the emoji for the current log level
        emoji = emojis.get(record["level"].name, "")
        
        # Format the log message
        return (
            "<green>{time:HH:mm:ss}</green> | "
            "<level>{level: <8}</level>" + emoji + " | "
            "<cyan>{name}:{function}:{line}</cyan> | "
            "{message}\n"
        ).format(**record)

    # Add handler with custom formatter
    logger.add(
        sys.stdout,
        level=level,
        format=format_record,
        colorize=True,
        enqueue=True
    )
    

def get_logger():
    return logger

# Plain text helpers (no color markup needed as loguru handles colors)
def blue(text: str) -> str:
    return text

def green(text: str) -> str:
    return text

def yellow(text: str) -> str:
    return text

def red(text: str) -> str:
    return text

def magenta(text: str) -> str:
    return text

def cyan(text: str) -> str:
    return text

def white(text: str) -> str:
    return text
