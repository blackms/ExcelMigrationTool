from typing import Optional, Dict, Any
import openpyxl
from ..logger import get_logger

logger = get_logger()

class FileRules:
    """Rules manager for specific file processing."""
    
    @staticmethod
    def get_column_mapping(filename: str) -> Dict[str, int]:
        """Get column mapping based on file type."""
        if "COaaS_Schema" in filename:
            return {
                "ru": 6,  # Column F for Resource Unit
                "startup_days": 5,  # Column E for GG Startup
                "profit_center": 18  # Column R for Profit Center
            }
        # Add mappings for other file types here
        return {}
    
    @staticmethod
    def get_resource_unit_rule(filename: str, element_type: str, row_data: Dict[str, Any]) -> Optional[str]:
        """Get Resource Unit value based on file-specific rules."""
        try:
            # Rules for COaaS Schema
            if "COaaS_Schema" in filename:
                if element_type == "Element":
                    logger.debug(f"Setting 'Per ogni contratto' for Element")
                    return "Per Ogni Contratto"
                elif element_type == "SubElement":
                    # Per i SubElement, ritorniamo None per usare il valore originale
                    logger.debug(f"Using original RU value for SubElement")
                    return None
            
            # Default: return None to use original value
            return None
            
        except Exception as e:
            logger.error(f"Error in resource unit rule for {filename}: {str(e)}")
            return None
    
    @staticmethod
    def apply_column_rules(filename: str, element_type: str, column: str, value: Any) -> Optional[Any]:
        """Apply rules for specific columns."""
        try:
            if "COaaS_Schema" in filename:
                if column == "ru" and element_type == "Element":
                    return "Per ogni contratto"
            return value
            
        except Exception as e:
            logger.error(f"Error applying column rules: {str(e)}")
            return value 