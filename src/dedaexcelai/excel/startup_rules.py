from typing import Optional, Dict, Any
import openpyxl
from ..excel.cell_operations import get_cell_value
from ..logger import get_logger, blue, yellow, red, green

logger = get_logger()

def get_startup_days_override(filename: str, primitive_data: openpyxl.worksheet.worksheet.Worksheet,
                            primitive_formulas: openpyxl.worksheet.worksheet.Worksheet,
                            formula: str, element_type: str, wbs_type: str) -> Optional[float]:
    """Get startup days override based on special rules."""
    try:
        logger.debug(f"Checking startup days override for file: {filename}")
        logger.debug(f"Element type: {element_type}, WBS type: {wbs_type}")
        
        if not primitive_data or not primitive_formulas:
            logger.warning("No PRIMITIVE sheets provided")
            return None
            
        # Special case for COaaS Schema
        if "COaaS_Schema" in filename:
            logger.info("Found COaaS Schema Prod file, checking U25 value")
            
            # Get values from PRIMITIVE sheet
            s25_cell = primitive_data.cell(row=25, column=19)  # Column S
            b13_cell = primitive_data.cell(row=13, column=2)   # Column B
            
            if not s25_cell or not b13_cell:
                return None
                
            s25_value = float(s25_cell.value or 0)
            b13_value = float(b13_cell.value or 1)
            
            logger.debug(f"S25 value: {s25_value}")
            logger.debug(f"B13 value: {b13_value}")
            
            if b13_value == 0:
                logger.warning("B13 value is 0, cannot divide")
                return None
                
            # Calculate days without rounding
            days = s25_value / b13_value
            logger.debug(f"Calculated value: {s25_value} / {b13_value} = {days}")
            
            logger.info(f"Using calculated value: {days:.2f} days (from S25/B13 = {s25_value}/{b13_value})")
            return days
            
        return None
        
    except Exception as e:
        logger.error(f"Error in startup days override: {str(e)}")
        return None
