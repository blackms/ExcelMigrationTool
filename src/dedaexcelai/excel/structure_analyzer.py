import openpyxl
from typing import Optional
from .cell_operations import get_cell_value

def find_header_row(sheet: openpyxl.worksheet.worksheet.Worksheet) -> int:
    """Find the header row in a sheet."""
    for row in range(1, min(10, sheet.max_row + 1)):  # Only check first 10 rows
        cell = sheet.cell(row=row, column=1)
        if get_cell_value(cell) == 'Type':
            return row
    return 1  # Default to first row if not found

def determine_cost_type(row: int, sheet: openpyxl.worksheet.worksheet.Worksheet) -> str:
    """Determine if a row represents a fixed or variable cost."""
    value = get_cell_value(sheet.cell(row=row, column=8))  # Column H
    if isinstance(value, str) and 'FIXED' in value.upper():
        return 'Fixed Mandatory'
    return 'Variable'
