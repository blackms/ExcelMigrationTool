import openpyxl

def set_euro_format(cell: openpyxl.cell.cell.Cell) -> None:
    """Set cell format to Euro currency."""
    cell.number_format = '#,##0.00€'

def clean_external_references(workbook: openpyxl.workbook.workbook.Workbook) -> None:
    """Remove any external references from workbook."""
    for sheet in workbook.worksheets:
        for row in sheet.rows:
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    # Remove external references in formulas
                    cell.value = cell.value.replace('[', '').replace(']', '')
