from copy import copy
import openpyxl
from openpyxl.styles import PatternFill
from typing import Any
from ..logger import get_logger

logger = get_logger()

def copy_cell_style(source_cell: Any, target_cell: Any):
    """
    Safely copy cell style attributes individually to avoid StyleProxy issues
    """
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def set_euro_format(cell: Any):
    """Set cell format to Euro with proper styling"""
    cell.number_format = '#,##0.00 €'
    # Add light blue background
    cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")

def clean_external_references(workbook: openpyxl.Workbook):
    """Remove any external references from the workbook"""
    for sheet in workbook.worksheets:
        logger.debug(f"Cleaning external references in sheet: {sheet.title}")
        # Check for external links in defined names
        for name in workbook.defined_names.values():
            if '[' in name.value:
                logger.debug(f"Removing external reference in defined name: {name.name}")
                workbook.defined_names.delete(name.name)
        
        # Check data validations
        if hasattr(sheet, 'data_validations'):
            for dv in sheet.data_validations.dataValidation:
                if hasattr(dv, 'formula1') and '[' in dv.formula1:
                    logger.debug(f"Removing external reference in data validation: {dv.formula1}")
                    dv.formula1 = dv.formula1.replace('[', '').replace(']', '')
        
        # Check cells for formulas with external references
        for row in sheet.rows:
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value.startswith('='):
                        original = cell.value
                        # Remove external references in formulas
                        cleaned = cell.value.replace('[', '').replace(']', '')
                        if cleaned != original:
                            logger.debug(f"Cleaning formula in {sheet.title}!{cell.coordinate}: {original} -> {cleaned}")
                            cell.value = cleaned
                    elif '[' in cell.value:
                        # Also clean any non-formula text that might contain external references
                        cell.value = cell.value.replace('[', '').replace(']', '')
