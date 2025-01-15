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
        # Initialize the StartupDaysAnalyzer if OpenAI key is provided
        startup_analyzer = None
        if openai_key:
            os.environ['OPENAI_API_KEY'] = openai_key
            startup_analyzer = StartupDaysAnalyzer()
            logger.info("Initialized StartupDaysAnalyzer for GG Startup analysis")
        else:
            logger.warning("No OpenAI API key provided. GG Startup analysis will be skipped.")

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
        logger.info("Loading input workbook...")
        input_wb = openpyxl.load_workbook(input_file, data_only=True, keep_links=False)
        
        logger.info("Loading template workbook...")
        template_wb = openpyxl.load_workbook(template_file, keep_links=False)
        
        # Skip EXPORT sheet for now as requested
        logger.info("Skipping EXPORT sheet handling for now...")

        # Clean any existing external references in template
        logger.info("Cleaning any existing external references in template...")
        clean_external_references(template_wb)
        
        # Get SCHEMA sheet from input workbook
        if "SCHEMA" not in input_wb.sheetnames:
            logger.error("Input file does not contain a 'SCHEMA' sheet")
            raise ValueError("Input file does not contain a 'SCHEMA' sheet")
        
        input_sheet = input_wb["SCHEMA"]
        output_sheet = template_wb.active
        
        logger.info("Found SCHEMA sheet, processing data...")
        
        # Find header row in input sheet
        header_row = find_header_row(input_sheet)
        logger.debug(f"Header row found at row {header_row}")
        
        # Process each row after headers
        current_output_row = 2  # Start after headers in template
        rows_processed = 0
        
        for row in range(header_row + 1, input_sheet.max_row + 1):
            # Get service element from column B
            service_element = get_cell_value(input_sheet.cell(row=row, column=2))
            
            # Skip completely empty rows
            if service_element is None:
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
                
                # Copy product element (column B)
                output_sheet.cell(row=current_output_row, column=2).value = service_element
                
                # Determine and set cost type (column C)
                cost_type = determine_cost_type(row, input_sheet)
                output_sheet.cell(row=current_output_row, column=3).value = cost_type

                # Handle GG Startup (column D) for Fixed costs
                if cost_type in ['Fixed Optional', 'Fixed Mandatory'] and startup_analyzer:
                    # Get description from column H
                    description = get_cell_value(input_sheet.cell(row=row, column=8))
                    if description:
                        startup_days = startup_analyzer.analyze_startup_days(description)
                        if startup_days is not None:
                            output_sheet.cell(row=current_output_row, column=4).value = startup_days
                            logger.debug(f"Set GG Startup to {startup_days} days for row {current_output_row}")
                
                # Copy Canone Prezzo Mese (column M in input to column N in output)
                canone_value = get_cell_value(input_sheet.cell(row=row, column=13))  # Column M is 13
                if canone_value is not None and is_number(canone_value):
                    output_cell = output_sheet.cell(row=current_output_row, column=14)  # Column N is 14
                    output_cell.value = float(canone_value)
                    set_euro_format(output_cell)
                    logger.debug(f"Copied Canone Prezzo Mese: {canone_value} to row {current_output_row}")
                
                # Set Euro format for column P
                set_euro_format(output_sheet.cell(row=current_output_row, column=16))
            
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
