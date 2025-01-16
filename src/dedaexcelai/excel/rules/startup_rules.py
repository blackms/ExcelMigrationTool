from typing import Optional, Dict, Any
import openpyxl
from ..utils.cell import get_cell_value
from ...logger import get_logger, blue, yellow, red, green
from .file_rules import FileRules

logger = get_logger()

def get_startup_days_override(filename: str, primitive_data: openpyxl.worksheet.worksheet.Worksheet,
                            primitive_formulas: openpyxl.worksheet.worksheet.Worksheet,
                            formula: str, element_type: str, wbs_type: str) -> Optional[float]:
    """Get startup days override based on special rules."""
    try:
        logger.debug(f"Checking startup days override for file: {filename}")
        
        rules = FileRules.get_startup_days_rule(filename)
        if not rules["enabled"]:
            return None
            
        if not primitive_data or not primitive_formulas:
            logger.warning("No PRIMITIVE sheets provided")
            return None
            
        # Get values based on rules
        values = []
        for cell_info in rules["source_cells"]:
            cell = primitive_data.cell(row=int(cell_info["cell"][1:]), 
                                    column=ord(cell_info["cell"][0]) - ord('A') + 1)
            if not cell:
                return None
            values.append(float(cell.value or 0))
            
        # Apply calculation
        if rules["calculation"]:
            days = rules["calculation"](*values)
            if days is not None:
                logger.info(f"Using calculated value: {days:.2f} days")
                return days
                
        return None
        
    except Exception as e:
        logger.error(f"Error in startup days override: {str(e)}")
        return None
