import openpyxl
from pathlib import Path
from typing import Dict, Any, Optional
import re
from loguru import logger
from copy import copy
from decimal import Decimal

def get_cell_value(cell: Optional[Any]) -> Any:
    """Helper function to safely get cell value"""
    if cell is None:
        return None
    return cell.value

def is_number(value: Any) -> bool:
    """Check if a value is a number (including string representations)"""
    if value is None:
        return False
    try:
        Decimal(str(value))
        return True
    except:
        return False

def determine_cost_type(row: int, sheet: openpyxl.worksheet.worksheet.Worksheet) -> str:
    """
    Determine the cost type based on which cost column is filled
    Returns: "Fee Optional" for una tantum costs, "Fixed Optional" for monthly recurrent costs
    """
    # Find the columns for Unatantu and Ricorrente Mese in Costo section
    unatantu_value = None
    ricorrente_value = None
    
    # Scan the header row to find the correct columns
    header_row = 1  # Assuming headers are in row 1
    for col in range(1, sheet.max_column + 1):
        header = get_cell_value(sheet.cell(row=header_row, column=col))
        if isinstance(header, str):
            if "unatantu" in header.lower():
                unatantu_value = get_cell_value(sheet.cell(row=row, column=col))
            elif "ricorrente" in header.lower() and "mese" in header.lower():
                ricorrente_value = get_cell_value(sheet.cell(row=row, column=col))

    # Check which type of cost is present
    if is_number(unatantu_value) and float(unatantu_value) > 0:
        logger.debug(f"Row {row}: Found una tantum cost: <yellow>{unatantu_value}</yellow>")
        return "Fee Optional"
    elif is_number(ricorrente_value) and float(ricorrente_value) > 0:
        logger.debug(f"Row {row}: Found monthly recurrent cost: <yellow>{ricorrente_value}</yellow>")
        return "Fixed Optional"
    else:
        logger.warning(f"Row {row}: No valid cost found, defaulting to <yellow>Fee Optional</yellow>")
        return "Fee Optional"

def find_header_row(sheet: openpyxl.worksheet.worksheet.Worksheet) -> int:
    """Find the row containing headers"""
    for row in range(1, sheet.max_row + 1):
        if get_cell_value(sheet.cell(row=row, column=1)) == "Portfolio":
            return row
    return 1

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

def migrate_excel(input_file: str, output_file: str, template_file: str) -> bool:
    """
    Migrate data from input Excel file to output Excel file using the specified template.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path where the output Excel file will be saved
        template_file: Path to the template Excel file
        
    Returns:
        bool: True if migration was successful, False otherwise
    """
    try:
        logger.info(f"Starting migration from <blue>{input_file}</blue> to <blue>{output_file}</blue>")
        logger.info(f"Using template: <blue>{template_file}</blue>")
        
        # Load workbooks without data_only to preserve formulas
        input_wb = openpyxl.load_workbook(input_file, data_only=True)  # Use data_only=True to get values from formulas
        template_wb = openpyxl.load_workbook(template_file)
        
        # Get SCHEMA sheet from input workbook
        if "SCHEMA" not in input_wb.sheetnames:
            logger.error("Input file does not contain a 'SCHEMA' sheet")
            raise ValueError("Input file does not contain a 'SCHEMA' sheet")
        
        input_sheet = input_wb["SCHEMA"]
        output_sheet = template_wb.active
        
        logger.info("<green>Found SCHEMA sheet, processing data...</green>")
        
        # Find header row in input sheet
        header_row = find_header_row(input_sheet)
        logger.debug(f"Header row found at row <yellow>{header_row}</yellow>")
        
        # Process each row after headers
        current_output_row = 2  # Start after headers in template
        rows_processed = 0
        
        for row in range(header_row + 1, input_sheet.max_row + 1):
            # Skip empty rows
            if not get_cell_value(input_sheet.cell(row=row, column=2)):  # Check column B instead of A
                continue
                
            # Skip separator rows
            service_element = get_cell_value(input_sheet.cell(row=row, column=2))
            if not service_element or service_element.startswith("---"):
                continue
            
            logger.debug(f"Processing row <blue>{row}</blue> -> <green>{current_output_row}</green>")
            
            # Copy product element (column B)
            output_sheet.cell(row=current_output_row, column=2).value = service_element
            
            # Determine and set cost type (column C)
            cost_type = determine_cost_type(row, input_sheet)
            output_sheet.cell(row=current_output_row, column=3).value = cost_type
            
            current_output_row += 1
            rows_processed += 1
        
        logger.info(f"Successfully processed <green>{rows_processed}</green> rows")
        
        # Save the output workbook
        template_wb.save(output_file)
        logger.success(f"Migration completed successfully. Output saved to: <blue>{output_file}</blue>")
        return True
        
    except Exception as e:
        logger.exception(f"Migration failed: <red>{str(e)}</red>")
        return False