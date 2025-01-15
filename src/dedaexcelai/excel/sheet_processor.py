from abc import ABC, abstractmethod
import openpyxl
from typing import Optional, Dict, Any
from ..logger import get_logger
from .cell_operations import get_cell_value, is_empty_or_dashes
from .style_manager import set_euro_format
from .file_rules import FileRules

logger = get_logger()

class SheetProcessor(ABC):
    """Base class for sheet processors."""
    
    @abstractmethod
    def process(self, source_sheet: openpyxl.worksheet.worksheet.Worksheet, target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> bool:
        """Process source sheet and write to target sheet."""
        pass

class PrimitiveSheetProcessor(SheetProcessor):
    """Processor for PRIMITIVE sheet."""
    
    def process(self, source_sheet: openpyxl.worksheet.worksheet.Worksheet, target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> bool:
        """Copy PRIMITIVE sheet as-is."""
        try:
            logger.info("Processing PRIMITIVE sheet...")
            for row in source_sheet.values:
                target_sheet.append(row)
            return True
        except Exception as e:
            logger.error(f"Failed to process PRIMITIVE sheet: {str(e)}")
            return False

class SchemaSheetProcessor(SheetProcessor):
    """Processor for SCHEMA sheet."""
    
    def __init__(self, startup_analyzer=None, filename: str = "", 
                 primitive_data=None, primitive_formulas=None):
        self.startup_analyzer = startup_analyzer
        self.filename = filename
        self.primitive_data = primitive_data
        self.primitive_formulas = primitive_formulas
        self.header_row = None
        self.current_output_row = 2  # Start after headers
        
    def setup_headers(self, target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Set up headers with green background."""
        headers = [
            'Lenght', 'Product Element', 'Object', 'Type', 'GG Startup', 'RU', 'RU Qty',
            'RU Unit of measure', 'Q.ty min', 'Q.ty MAX', '%Sconto MAX',
            'Startup Costo', 'Startup Margine', 'Startup Prezzo',
            'Canone Costo Mese', 'Canone Margine', 'Canone Prezzo Mese',
            'Extended Description', 'Profit Center Prevalente', 'Status', 'Note'
        ]
        
        green_fill = openpyxl.styles.PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
        
        for col, header in enumerate(headers, 1):
            cell = target_sheet.cell(row=1, column=col, value=header)
            cell.fill = green_fill
            cell.font = openpyxl.styles.Font(bold=True)
    
    def determine_type(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet) -> str:
        """Determine if cost is Fee or Fixed based on values."""
        # Check WBS column (D) for FIXED
        wbs = get_cell_value(source_sheet.cell(row=row, column=4))
        if isinstance(wbs, str) and 'FIXED' in wbs.upper():
            return 'Fixed Optional'
            
        # Check ricorrente (H) and canone (I)
        ricorrente = get_cell_value(source_sheet.cell(row=row, column=8))
        canone = get_cell_value(source_sheet.cell(row=row, column=9))
        
        # If has ricorrente but no canone, it's a startup cost (Fixed)
        if (ricorrente and str(ricorrente).strip() and not str(ricorrente).strip().startswith('-') and
            (not canone or not str(canone).strip() or str(canone).strip().startswith('-'))):
            return 'Fixed Optional'
            
        # Default to Fee Optional
        return 'Fee Optional'
    
    def process_row(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                   target_sheet: openpyxl.worksheet.worksheet.Worksheet, source_formulas: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process a single row from source to target."""
        # Get type and service element
        element_type = get_cell_value(source_sheet.cell(row=row, column=1))  # Type column
        service_element = get_cell_value(source_sheet.cell(row=row, column=2))  # Service Element column
        
        if service_element is None or is_empty_or_dashes(service_element):
            return
            
        # Set Lenght for Elements
        if element_type == 'Element':
            text_length = len(str(service_element)) if service_element else 0
            target_sheet.cell(row=self.current_output_row, column=1).value = text_length
        
        # Copy Product Element and set formatting
        element_cell = target_sheet.cell(row=self.current_output_row, column=2)
        element_cell.value = service_element
        element_cell.font = openpyxl.styles.Font(bold=(element_type == 'Element'), italic=(element_type == 'SubElement'))
        
        # Set Object (Element/SubElement)
        target_sheet.cell(row=self.current_output_row, column=3).value = element_type
        
        # Set Type (Fee/Fixed)
        cost_type = self.determine_type(row, source_sheet)
        target_sheet.cell(row=self.current_output_row, column=4).value = cost_type
        
        # Process other columns based on mapping
        self.process_mapped_columns(row, source_sheet, target_sheet)
        
        # Handle GG Startup for Fixed costs
        cost_type = self.determine_type(row, source_sheet)
        logger.info(f"Row {row} - Element Type: {element_type}, Cost Type: {cost_type}")
        
        if cost_type.startswith('Fixed'):
            logger.info(f"Processing startup days for Fixed cost in row {row}")
            formula = source_formulas.cell(row=row, column=8).value  # Column H
            logger.info(f"Found formula in column H: {formula}")
            self.process_startup_days(row, source_sheet, target_sheet, source_formulas)
        else:
            logger.debug(f"Skipping startup days for non-Fixed cost in row {row}")
        
        self.current_output_row += 1
    
    def copy_value(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet,
                  target_sheet: openpyxl.worksheet.worksheet.Worksheet, source_col: int, target_col: int) -> None:
        """Copy value from source to target cell."""
        value = get_cell_value(source_sheet.cell(row=row, column=source_col))
        if value is not None:
            target_cell = target_sheet.cell(row=self.current_output_row, column=target_col)
            target_cell.value = value
            # Set number format with 4 decimal places for numeric values
            if isinstance(value, (int, float)):
                target_cell.number_format = '#,##0.0000'
    
    def process_mapped_columns(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                             target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process columns based on mapping."""
        try:
            element_type = source_sheet.cell(row=row, column=1).value
            
            # Get column mapping for this file type
            column_mapping = FileRules.get_column_mapping(self.filename)
            
            # Process Resource Unit (RU)
            source_value = get_cell_value(source_sheet.cell(row=row, column=3))
            ru_value = FileRules.apply_column_rules(self.filename, element_type, "ru", source_value)
            if ru_value is not None:
                target_cell = target_sheet.cell(row=self.current_output_row, column=column_mapping["ru"])
                target_cell.value = ru_value
                logger.debug(f"Set RU value: {ru_value} for row {row}")
            
            # Process other columns...
            
        except Exception as e:
            logger.error(f"Error in process_mapped_columns: {str(e)}")
    
    def process_startup_days(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                            target_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                            source_formulas: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process startup days for the given row."""
        try:
            logger.info(f"Processing startup days for row {row}")
            
            # Get element type and cost type from correct columns
            element_type = get_cell_value(source_sheet.cell(row=row, column=1))  # Column A
            cost_type = self.determine_type(row, source_sheet)
            
            formula = source_formulas.cell(row=row, column=8).value  # H column
            
            if not formula:
                logger.debug(f"No formula found in column H for row {row}")
                return
            
            logger.info(f"Found formula in column H: {formula}")
            
            if not self.startup_analyzer:
                logger.warning("No startup analyzer available")
                return
            
            logger.info("Calling startup analyzer...")
            logger.debug(f"Processing element type: {element_type}")
            logger.debug(f"Processing cost type: {cost_type}")
            
            days = self.startup_analyzer.analyze_startup_days(
                formula=formula,
                primitive_formulas=self.primitive_formulas,
                primitive_data=self.primitive_data,
                row=row,
                schema_sheet=source_sheet,
                schema_formulas=source_formulas,
                filename=self.filename,
                element_type=element_type,
                wbs_type=cost_type
            )
            
            if days is not None:
                logger.info(f"Setting startup days to {days} for row {row}")
                target_cell = target_sheet.cell(row=self.current_output_row, column=5)
                target_cell.value = days
                target_cell.number_format = '#,##0.000000'  # Formato con 6 decimali
            else:
                logger.debug(f"No startup days determined for row {row}")
            
        except Exception as e:
            logger.error(f"Error analyzing startup days for row {row}: {str(e)}")
    
    def process(self, source_sheet: openpyxl.worksheet.worksheet.Worksheet, target_sheet: openpyxl.worksheet.worksheet.Worksheet,
                source_formulas: Optional[openpyxl.worksheet.worksheet.Worksheet] = None) -> bool:
        """Process SCHEMA sheet with new format."""
        try:
            logger.info("Processing SCHEMA sheet...")
            
            # Set up headers
            self.setup_headers(target_sheet)
            
            # Find header row in source
            from .structure_analyzer import find_header_row
            self.header_row = find_header_row(source_sheet)
            
            # Process rows
            for row in range(self.header_row + 1, source_sheet.max_row + 1):
                self.process_row(row, source_sheet, target_sheet, source_formulas)
                
            return True
            
        except Exception as e:
            logger.error(f"Failed to process SCHEMA sheet: {str(e)}")
            return False
