from typing import Any, Optional
from decimal import Decimal
from openpyxl.cell import Cell

def get_cell_value(cell: Optional[Any]) -> Any:
    """Helper function to safely get cell value"""
    if cell is None:
        return None
    return cell.value

def is_empty_or_dashes(value: Any) -> bool:
    """Check if a value is empty, None, or contains only dashes"""
    if value is None or value == "":
        return True
    if isinstance(value, str):
        return value.strip().replace("-", "") == ""
    return False

def is_number(value: Any) -> bool:
    """Check if a value is a number (including string representations)"""
    if value is None:
        return False
    try:
        Decimal(str(value))
        return True
    except:
        return False
