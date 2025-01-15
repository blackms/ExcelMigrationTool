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

def extract_cell_references(formula: str, logger) -> list[tuple[str, str]]:
    """
    Extract sheet name and cell references from a formula.
    Example: "=PRIMITIVE!U25*PRIMITIVE!B12" -> [("PRIMITIVE", "U25"), ("PRIMITIVE", "B12")]
    """
    try:
        # Remove the leading = if present
        if formula.startswith('='):
            formula = formula[1:]
        
        refs = []
        parts = formula.split('*')  # Split by multiplication operator
        logger.debug(f"Split formula '{formula}' into parts: {parts}")
        
        for part in parts:
            if '!' in part:  # Contains sheet reference
                sheet, cell = part.split('!')
                refs.append((sheet, cell))
                logger.debug(f"Extracted reference: Sheet={sheet}, Cell={cell}")
        
        return refs
    except Exception as e:
        logger.error(f"Error extracting cell references from '{formula}': {str(e)}")
        return []

def get_cell_value_with_fallback(cell_formulas, cell_data, logger) -> Optional[float]:
    """
    Try to get a numeric value from either a formula cell or data cell.
    Prefers data value over formula value.
    """
    try:
        # Try to get the value from both cells
        formula_value = get_cell_value(cell_formulas)
        data_value = get_cell_value(cell_data)
        
        logger.debug(f"Found values - Formula: {formula_value}, Data: {data_value}")
        
        # Prefer the data value if available
        if data_value is not None and isinstance(data_value, (int, float)):
            logger.info(f"Using data value: {data_value}")
            return float(data_value)
        elif formula_value is not None and isinstance(formula_value, (int, float)):
            logger.info(f"Using formula value: {formula_value}")
            return float(formula_value)
        
        return None
    except Exception as e:
        logger.error(f"Error getting cell value: {str(e)}")
        return None
