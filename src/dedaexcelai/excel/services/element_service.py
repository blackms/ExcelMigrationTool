"""Services for Excel processing."""
from typing import Optional, Tuple, List
import openpyxl
from dedaexcelai.excel.models.elements import Element, ElementType, CostType, CostMapping, FixedCostMapping, FeeCostMapping
from dedaexcelai.excel.utils.cell import get_cell_value
from dedaexcelai.logger import get_logger

logger = get_logger()

class ElementService:
    """Service for handling elements."""
    
    @staticmethod
    def create_element(row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet,
                      primitive_sheet: Optional[openpyxl.worksheet.worksheet.Worksheet] = None) -> Optional[Element]:
        """Create an Element from a worksheet row."""
        try:
            element_type_str = get_cell_value(source_sheet.cell(row=row, column=1))
            if not element_type_str:
                return None
                
            element_type = ElementType(element_type_str)
            name = get_cell_value(source_sheet.cell(row=row, column=2))
            
            if not name or str(name).strip().startswith('-'):
                return None
                
            # Determine cost type
            cost_type = ElementService._determine_cost_type(row, source_sheet, primitive_sheet)
            
            # Calculate length for Elements
            length = None
            if element_type == ElementType.ELEMENT:
                length = len(str(name))
            
            return Element(
                name=name,
                element_type=element_type,
                cost_type=cost_type,
                row=row,
                length=length
            )
            
        except Exception as e:
            logger.error("Error creating element from row {}: {}", row, str(e))
            return None
    
    @staticmethod
    def _determine_cost_type(row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet,
                           primitive_sheet: Optional[openpyxl.worksheet.worksheet.Worksheet] = None) -> CostType:
        """Determine the cost type for a row."""
        from dedaexcelai.excel.utils.structure import determine_cost_type
        cost_type_str = determine_cost_type(row, source_sheet, primitive_sheet)
        return CostType(cost_type_str)
