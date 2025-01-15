import os
import tempfile
import shutil
import openpyxl
from typing import Optional

from ..logger import get_logger, blue
from ..llm import StartupDaysAnalyzer
from .cell_operations import get_cell_value, is_empty_or_dashes, is_number
from .structure_analyzer import find_header_row, determine_cost_type
from .style_manager import set_euro_format, clean_external_references

logger = get_logger()

def migrate_excel(input_file: str, output_file: str, template_file: str, openai_key: Optional[str] = None) -> bool:
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
        # Initialize the StartupDaysAnalyzer with OpenAI key
        startup_analyzer = StartupDaysAnalyzer(openai_key)
        logger.info("Initialized StartupDaysAnalyzer with GPT-4 for GG Startup analysis")

        # First check if output file is accessible and try to close it
        if os.path.exists(output_file):
            try:
                # Try to remove the file first
                os.remove(output_file)
                logger.info(f"Removed existing output file: {output_file}")
            except PermissionError:
                error_msg = (
                    f"Cannot access {output_file} - the file is open in Excel or read-only. "
                    "Please close Excel and try again."
                )
                logger.error(error_msg)
                return False
            except Exception as e:
                logger.warning(f"Could not remove existing file: {str(e)}")
            
        logger.info(f"Starting migration from {input_file} to {output_file}")
        logger.info(f"Using template: {template_file}")
        
        # Load input workbook
        logger.info("Loading input workbook...")
        input_wb_data = openpyxl.load_workbook(input_file, data_only=True, keep_links=False)
        input_wb_formulas = openpyxl.load_workbook(input_file, data_only=False, keep_links=False)
        
        # Create new workbook for output
        logger.info("Creating new workbook...")
        output_wb = openpyxl.Workbook()
        
        # Keep PRIMITIVE sheet from input
        logger.info("Copying PRIMITIVE sheet...")
        if 'PRIMITIVE' not in input_wb_data.sheetnames:
            logger.error("PRIMITIVE sheet not found in input workbook")
            raise ValueError("PRIMITIVE sheet not found in input workbook")
            
        # Copy PRIMITIVE sheet
        source_sheet = input_wb_data['PRIMITIVE']
        if 'Sheet' in output_wb.sheetnames:  # Remove default sheet
            del output_wb['Sheet']
        target_sheet = output_wb.create_sheet('PRIMITIVE')
        for row in source_sheet.values:
            target_sheet.append(row)
            
        # Create new SCHEMA sheet with headers
        logger.info("Creating SCHEMA sheet...")
        schema_sheet = output_wb.create_sheet('SCHEMA', 0)  # Make it the first sheet
        
        # Add headers from green background image
        headers = [
            'Lenght', 'Product Element', 'Type', 'GG Startup', 'RU', 'RU Qty',
            'RU Unit of measure', 'Q.ty min', 'Q.ty MAX', '%Sconto MAX',
            'Startup Costo', 'Startup Margine', 'Startup Prezzo',
            'Canone Costo Mese', 'Canone Margine', 'Canone Prezzo Mese',
            'Extended Description', 'Profit Center Prevalente', 'Status', 'Note'
        ]
        
        # Set header background color (green)
        green_fill = openpyxl.styles.PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
        
        for col, header in enumerate(headers, 1):
            cell = schema_sheet.cell(row=1, column=col, value=header)
            cell.fill = green_fill
            cell.font = openpyxl.styles.Font(bold=True)
        
        # Get SCHEMA sheet from input workbook
        if "SCHEMA" not in input_wb_data.sheetnames:
            logger.error("Input file does not contain a 'SCHEMA' sheet")
            raise ValueError("Input file does not contain a 'SCHEMA' sheet")
        
        input_sheet = input_wb_data["SCHEMA"]
        input_sheet_formulas = input_wb_formulas["SCHEMA"]
        
        # Column mapping from old to new schema
        column_mapping = {
            'Type': 3,  # Type in new schema
            'Service Element (Building block)': 2,  # Product Element in new schema
            'Resource Unit': 5,  # RU in new schema
            'WBS': None,  # Not used in new schema
            'Materiale': None,  # Not used in new schema
            'Profit Center': 18,  # Profit Center Prevalente in new schema
            'Unatantu': 11,  # Startup Costo in new schema
            'Ricorrente mese': 14,  # Canone Costo Mese in new schema
            'Canone': 16,  # Canone Prezzo Mese in new schema
        }
        
        logger.info("Found SCHEMA sheet, processing data...")
        
        # Find header row in input sheet
        header_row = find_header_row(input_sheet)
        logger.debug(f"Header row found at row {header_row}")
        
        # Process each row after headers
        current_output_row = 2  # Start after headers
        rows_processed = 0
        
        for row in range(header_row + 1, input_sheet.max_row + 1):
            # Get type and service element
            element_type = get_cell_value(input_sheet.cell(row=row, column=1))  # Type column
            service_element = get_cell_value(input_sheet.cell(row=row, column=2))  # Service Element column
            
            # Skip empty rows
            if service_element is None or is_empty_or_dashes(service_element):
                continue
                
            logger.debug(f"Processing row {row} -> {current_output_row}")
            
            # Set Lenght (first column) only for Elements
            if element_type == 'Element':
                # Get length of text in column B
                text_length = len(str(service_element)) if service_element else 0
                schema_sheet.cell(row=current_output_row, column=1).value = text_length
            else:
                schema_sheet.cell(row=current_output_row, column=1).value = ""
            
            # Copy Product Element
            element_cell = schema_sheet.cell(row=current_output_row, column=2)
            element_cell.value = service_element
            
            # Set Type
            schema_sheet.cell(row=current_output_row, column=3).value = element_type
            
            # Handle GG Startup (column 4) for Fixed costs
            if element_type == 'SubElement':  # Only process startup days for sub-elements
                cost_type = determine_cost_type(row, input_sheet)
                if cost_type in ['Fixed Optional', 'Fixed Mandatory']:
                    formula_cell = input_sheet_formulas.cell(row=row, column=8)
                    formula = formula_cell.value if formula_cell else None
                    if formula and isinstance(formula, str):
                        try:
                            primitive_sheet_data = input_wb_data['PRIMITIVE']
                            primitive_sheet_formulas = input_wb_formulas['PRIMITIVE']
                            startup_days = startup_analyzer.analyze_startup_days(formula, primitive_sheet_formulas, primitive_sheet_data, row, input_sheet, input_sheet_formulas)
                            if startup_days is not None:
                                schema_sheet.cell(row=current_output_row, column=4).value = startup_days
                        except Exception as e:
                            logger.error(f"Error analyzing startup days for row {row}: {str(e)}")
            
            # Copy other mapped columns
            for old_col in range(1, input_sheet.max_column + 1):
                old_header = get_cell_value(input_sheet.cell(row=header_row, column=old_col))
                if old_header in column_mapping and column_mapping[old_header] is not None:
                    new_col = column_mapping[old_header]
                    value = get_cell_value(input_sheet.cell(row=row, column=old_col))
                    if value is not None:
                        schema_sheet.cell(row=current_output_row, column=new_col).value = value
            
            # Set formatting based on type
            if element_type == 'Element':
                element_cell.font = openpyxl.styles.Font(bold=True)
            else:  # SubElement
                element_cell.font = openpyxl.styles.Font(italic=True)
            
            current_output_row += 1
            rows_processed += 1
        
        logger.info(f"Successfully processed {rows_processed} rows")
        
        # Save to temporary file first
        logger.info("Saving output workbook...")
        try:
            # Create temp file in same directory as output file
            temp_dir = os.path.dirname(os.path.abspath(output_file))
            temp_fd, temp_path = tempfile.mkstemp(dir=temp_dir, suffix='.xlsx')
            os.close(temp_fd)  # Close file descriptor
            
            # Save to temp file
            output_wb.save(temp_path)
            logger.debug(f"Saved to temporary file: {temp_path}")
            
            # Try to replace output file with temp file
            try:
                if os.path.exists(output_file):
                    os.remove(output_file)
                shutil.move(temp_path, output_file)
                logger.info(f"Migration completed successfully. Output saved to: {output_file}")
                return True
            except PermissionError:
                error_msg = (
                    f"Cannot access {output_file} - the file is open in Excel or read-only. "
                    "Please close Excel and try again."
                )
                logger.error(error_msg)
                os.remove(temp_path)  # Clean up temp file
                return False
            except Exception as e:
                logger.exception(f"Failed to replace output file: {str(e)}")
                os.remove(temp_path)  # Clean up temp file
                return False
                
        except Exception as e:
            logger.exception(f"Failed to save output file: {str(e)}")
            return False
        
    except Exception as e:
        logger.exception(f"Migration failed during processing: {str(e)}")
        return False
