from typing import Any, Optional, List, Tuple
import openpyxl
import re

def get_cell_value(cell: Optional[openpyxl.cell.cell.Cell]) -> Any:
    """Get cell value, handling None cells."""
    return cell.value if cell else None

def is_empty_or_dashes(value: Any) -> bool:
    """Check if value is empty or contains only dashes."""
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip() or all(c == '-' for c in value.strip())
    return False

def is_number(value: Any) -> bool:
    """Check if value can be converted to float."""
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False

def extract_cell_references(formula: str) -> List[Tuple[str, str]]:
    """Extract sheet and cell references from Excel formula."""
    refs = []
    
    # Handle SUM ranges
    sum_pattern = r'SUM\((?:([^!]+)!)?([A-Z]+\d+):([A-Z]+\d+)\)'
    sum_matches = re.finditer(sum_pattern, formula)
    for match in sum_matches:
        sheet = match.group(1) if match.group(1) else 'SCHEMA'  # Default to current sheet if no sheet specified
        start_cell = match.group(2)
        end_cell = match.group(3)
        
        # Convert to column letters and row numbers
        start_col = ''.join(c for c in start_cell if c.isalpha())
        start_row = int(''.join(c for c in start_cell if c.isdigit()))
        end_col = ''.join(c for c in end_cell if c.isalpha())
        end_row = int(''.join(c for c in end_cell if c.isdigit()))
        
        # Add all cells in range
        for row in range(start_row, end_row + 1):
            cell_ref = f"{start_col}{row}"
            refs.append((sheet, cell_ref))
    
    # Handle direct cell references (including multiplication)
    cell_pattern = r"(?:'?([^'!]+)'?|([^!*]+))!(\$?[A-Z]+\$?\d+)"
    cell_matches = re.finditer(cell_pattern, formula)
    for match in cell_matches:
        sheet = match.group(1) or match.group(2)  # Either quoted or unquoted sheet name
        if sheet.startswith('='): sheet = sheet[1:]  # Remove leading = if present
        cell = match.group(3)
        refs.append((sheet, cell.replace('$', '')))  # Remove $ signs
    
    return refs
