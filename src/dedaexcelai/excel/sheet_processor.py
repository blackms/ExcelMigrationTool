from abc import ABC, abstractmethod
import openpyxl
from typing import Optional, Dict, Any
from ..logger import get_logger
from .cell_operations import get_cell_value, is_empty_or_dashes
from .style_manager import set_euro_format

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
    
    def __init__(self, startup_analyzer=None):
        self.startup_analyzer = startup_analyzer
        self.header_row = None
        self.current_output_row = 2  # Start after headers
        
    def setup_headers(self, target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Set up headers with green background."""
        headers = [
            'Lenght', 'Product Element', 'Type', 'GG Startup', 'RU', 'RU Qty',
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
    
    def process_row(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                   target_sheet: openpyxl.worksheet.worksheet.Worksheet, source_formulas: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process a single row from source to target."""
        # Get type and service element
        element_type = get_cell_value(source_sheet.cell(row=row, column=1))
        service_element = get_cell_value(source_sheet.cell(row=row, column=2))
        
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
        
        # Set Type
        target_sheet.cell(row=self.current_output_row, column=3).value = element_type
        
        # Process other columns based on mapping
        self.process_mapped_columns(row, source_sheet, target_sheet)
        
        # Handle GG Startup for SubElements
        if element_type == 'SubElement':
            self.process_startup_days(row, source_sheet, target_sheet, source_formulas)
        
        self.current_output_row += 1
    
    def process_mapped_columns(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet, 
                             target_sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process columns based on mapping."""
        column_mapping = {
            'Resource Unit': 5,  # RU
            'Profit Center': 18,  # Profit Center Prevalente
            'Unatantu': 11,  # Startup Costo
            'Ricorrente mese': 14,  # Canone Costo Mese
            'Canone': 16,  # Canone Prezzo Mese
        }
        
        for old_col in range(1, source_sheet.max_column + 1):
            old_header = get_cell_value(source_sheet.cell(row=self.header_row, column=old_col))
            if old_header in column_mapping:
                value = get_cell_value(source_sheet.cell(row=row, column=old_col))
                if value is not None:
                    target_sheet.cell(row=self.current_output_row, column=column_mapping[old_header]).value = value
    
    def process_startup_days(self, row: int, source_sheet: openpyxl.worksheet.worksheet.Worksheet,
                           target_sheet: openpyxl.worksheet.worksheet.Worksheet,
                           source_formulas: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Process startup days for a row."""
        if not self.startup_analyzer:
            return
            
        formula_cell = source_formulas.cell(row=row, column=8)
        formula = formula_cell.value if formula_cell else None
        
        if formula and isinstance(formula, str):
            try:
                startup_days = self.startup_analyzer.analyze_startup_days(
                    formula, source_formulas, source_sheet, row, source_sheet, source_formulas
                )
                if startup_days is not None:
                    target_sheet.cell(row=self.current_output_row, column=4).value = startup_days
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
