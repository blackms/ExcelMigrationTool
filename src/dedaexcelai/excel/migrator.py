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
        
        # Load workbooks with additional parameters to handle external links
        # Load one copy with data_only=True for values and one without for formulas
        logger.info("Loading input workbook...")
        input_wb_data = openpyxl.load_workbook(input_file, data_only=True, keep_links=False)
        input_wb_formulas = openpyxl.load_workbook(input_file, data_only=False, keep_links=False)
        
        logger.info("Loading template workbook...")
        template_wb = openpyxl.load_workbook(template_file, keep_links=False)
        
        # Skip EXPORT sheet for now as requested
        logger.info("Skipping EXPORT sheet handling for now...")

        # Clean any existing external references in template
        logger.info("Cleaning any existing external references in template...")
        clean_external_references(template_wb)
        
        # Get SCHEMA sheet from input workbook
        if "SCHEMA" not in input_wb_data.sheetnames:
            logger.error("Input file does not contain a 'SCHEMA' sheet")
            raise ValueError("Input file does not contain a 'SCHEMA' sheet")
        
        input_sheet = input_wb_data["SCHEMA"]
        input_sheet_formulas = input_wb_formulas["SCHEMA"]
        output_sheet = template_wb.active
        
        logger.info("Found SCHEMA sheet, processing data...")
        
        # Find header row in input sheet
        header_row = find_header_row(input_sheet)
        logger.debug(f"Header row found at row {header_row}")
        
        # Process each row after headers
        current_output_row = 2  # Start after headers in template
        rows_processed = 0
        processed_rows = set()  # Keep track of rows we've processed as sub-elements
        
        for row in range(header_row + 1, input_sheet.max_row + 1):
            # Get service element from column B
            service_element = get_cell_value(input_sheet.cell(row=row, column=2))
            
            # Skip completely empty rows or already processed rows
            if service_element is None or row in processed_rows:
                continue
            
            # Handle empty or dash-only service elements differently
            if is_empty_or_dashes(service_element):
                logger.debug(f"Empty row {row}, adding formula to column P")
                # Copy Canone Prezzo Mese from input (column M) to output (column P)
                canone_value = get_cell_value(input_sheet.cell(row=row, column=13))  # Column M is 13
                if canone_value is not None and is_number(canone_value):
                    output_cell = output_sheet.cell(row=current_output_row, column=16)  # Column P is 16
                    output_cell.value = float(canone_value)
                    set_euro_format(output_cell)
                    logger.debug(f"Copied Canone Prezzo Mese: {canone_value} to row {current_output_row}")
                else:
                    # If no value, set to 0
                    output_cell = output_sheet.cell(row=current_output_row, column=16)
                    output_cell.value = 0
                    set_euro_format(output_cell)
            else:
                logger.debug(f"Processing row {row} -> {current_output_row}")
                
                # Copy element to column B
                element_cell = output_sheet.cell(row=current_output_row, column=2)
                element_cell.value = service_element
                
                # Check if this is an element catalog by looking for bold text in input
                cell = input_sheet.cell(row=row, column=2)
                is_element_catalog = cell.font and cell.font.b
                
                # Set formatting based on whether it's an element catalog or sub-element
                if is_element_catalog:
                    # Element catalog - set bold
                    if not element_cell.font:
                        element_cell.font = openpyxl.styles.Font(bold=True)
                    else:
                        element_cell.font = openpyxl.styles.Font(bold=True, name=element_cell.font.name, size=element_cell.font.size)
                else:
                    # Sub-element - set italic
                    if not element_cell.font:
                        element_cell.font = openpyxl.styles.Font(italic=True)
                    else:
                        element_cell.font = openpyxl.styles.Font(italic=True, name=element_cell.font.name, size=element_cell.font.size)
                
                # Determine and set cost type (column C)
                cost_type = determine_cost_type(row, input_sheet)
                output_sheet.cell(row=current_output_row, column=3).value = cost_type

                # Handle GG Startup (column D) for Fixed costs
                logger.info(f"Processing GG Startup for row {row} -> {current_output_row}")
                logger.info(f"Cost type: {cost_type}")
                
                if cost_type in ['Fixed Optional', 'Fixed Mandatory']:
                    # Get formula from column H using the formulas workbook
                    formula_cell = input_sheet_formulas.cell(row=row, column=8)
                    formula = formula_cell.value if formula_cell else None
                    logger.info(f"Formula from column H: {formula}")
                    
                    if formula and isinstance(formula, str):
                        try:
                            # Get PRIMITIVE sheet from both workbooks
                            if 'PRIMITIVE' not in input_wb_data.sheetnames:
                                logger.error("PRIMITIVE sheet not found in workbook")
                                continue
                            
                            primitive_sheet_data = input_wb_data['PRIMITIVE']
                            primitive_sheet_formulas = input_wb_formulas['PRIMITIVE']
                            startup_days = startup_analyzer.analyze_startup_days(formula, primitive_sheet_formulas, primitive_sheet_data, row, input_sheet)
                            
                            if startup_days is not None:
                                output_sheet.cell(row=current_output_row, column=4).value = startup_days
                                logger.info(f"Set GG Startup to {startup_days} days for row {current_output_row}")
                            else:
                                logger.warning(f"Could not extract startup days from formula on row {row}")
                        except Exception as e:
                            logger.error(f"Error analyzing startup days for row {row}: {str(e)}")
                    else:
                        logger.warning(f"No valid PRIMITIVE formula found in column H for row {row}")
                else:
                    logger.debug(f"Skipping GG Startup for non-Fixed cost type: {cost_type}")
                
                # Copy Canone Prezzo Mese (column M in input to column N in output)
                canone_value = get_cell_value(input_sheet.cell(row=row, column=13))  # Column M is 13
                if canone_value is not None and is_number(canone_value):
                    output_cell = output_sheet.cell(row=current_output_row, column=14)  # Column N is 14
                    output_cell.value = float(canone_value)
                    set_euro_format(output_cell)
                    logger.debug(f"Copied Canone Prezzo Mese: {canone_value} to row {current_output_row}")
                
                # Handle element catalogs and sub-elements
                if is_element_catalog:
                    # Set SUM formula in P2 (current row)
                    sum_cell = output_sheet.cell(row=current_output_row, column=16)  # Column P
                    sum_cell.value = f"=SUM(P{current_output_row + 1}:P{current_output_row + 9})"
                    set_euro_format(sum_cell)
                    
                    # Find and process existing sub-elements
                    sub_elements = []
                    for r in range(row + 1, input_sheet.max_row + 1):
                        sub_cell = input_sheet.cell(row=r, column=2)
                        sub_value = get_cell_value(sub_cell)
                        if sub_value and isinstance(sub_value, str) and "---" in sub_value:
                            break
                        if sub_value and not is_empty_or_dashes(sub_value):
                            sub_elements.append(r)
                            processed_rows.add(r)  # Mark this row as processed
                    
                    # Move to next row for sub-elements
                    current_output_row += 1
                    
                    # Process existing sub-elements first
                    for i, sub_row in enumerate(sub_elements):
                        # Copy sub-element name
                        sub_cell = output_sheet.cell(row=current_output_row + i, column=2)
                        sub_cell.value = get_cell_value(input_sheet.cell(row=sub_row, column=2))
                        if not sub_cell.font:
                            sub_cell.font = openpyxl.styles.Font(italic=True)
                        else:
                            sub_cell.font = openpyxl.styles.Font(italic=True, name=sub_cell.font.name, size=sub_cell.font.size)
                        
                        # Copy original formula from column H to column P
                        formula = input_sheet_formulas.cell(row=sub_row, column=8).value  # Column H
                        if formula and isinstance(formula, str):
                            output_cell = output_sheet.cell(row=current_output_row + i, column=16)  # Column P
                            output_cell.value = formula
                            set_euro_format(output_cell)
                    
                    # Add remaining empty rows with single dash
                    remaining_rows = 9 - len(sub_elements)
                    for i in range(remaining_rows):
                        sub_row = current_output_row + len(sub_elements) + i
                        dash_cell = output_sheet.cell(row=sub_row, column=2)
                        dash_cell.value = "-"
                        if not dash_cell.font:
                            dash_cell.font = openpyxl.styles.Font(italic=True)
                        else:
                            dash_cell.font = openpyxl.styles.Font(italic=True, name=dash_cell.font.name, size=dash_cell.font.size)
                        set_euro_format(output_sheet.cell(row=sub_row, column=16))
                    
                    current_output_row += 9  # Move past all sub-elements (exactly 9 total)
                else:
                    # This is a sub-element
                    # Copy value to column P
                    canone_value = get_cell_value(input_sheet.cell(row=row, column=13))  # Column M is 13
                    if canone_value is not None and is_number(canone_value):
                        output_cell = output_sheet.cell(row=current_output_row, column=16)  # Column P
                        output_cell.value = float(canone_value)
                        set_euro_format(output_cell)
                    current_output_row += 1
                
                rows_processed += 1
        
        logger.info(f"Successfully processed {rows_processed} rows")
        
        # Final cleanup and save
        logger.info("Performing final cleanup of external references...")
        clean_external_references(template_wb)
        
        # Save to temporary file first
        logger.info("Saving output workbook...")
        try:
            # Create temp file in same directory as output file
            temp_dir = os.path.dirname(os.path.abspath(output_file))
            temp_fd, temp_path = tempfile.mkstemp(dir=temp_dir, suffix='.xlsx')
            os.close(temp_fd)  # Close file descriptor
            
            # Save to temp file
            template_wb.save(temp_path)
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
