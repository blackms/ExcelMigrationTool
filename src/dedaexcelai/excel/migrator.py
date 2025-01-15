from typing import Optional
from ..logger import get_logger
from ..llm import StartupDaysAnalyzer
from .workbook_handler import WorkbookHandler
from .sheet_processor_factory import SheetProcessorFactory

logger = get_logger()

def migrate_excel(input_file: str, output_file: str, template_file: str, openai_key: Optional[str] = None) -> bool:
    """
    Migrate Excel file using new architecture with separate handlers and processors.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path where the output Excel file will be saved
        template_file: Path to the template file (not used in new architecture)
        openai_key: Optional OpenAI API key for startup days analysis
        
    Returns:
        bool: True if migration was successful, False otherwise
    """
    try:
        # Initialize components
        startup_analyzer = StartupDaysAnalyzer(openai_key)
        workbook_handler = WorkbookHandler(input_file, output_file)
        
        # Load workbooks
        if not workbook_handler.load_workbooks():
            return False
            
        # Create output workbook
        if not workbook_handler.create_output_workbook():
            return False
            
        # Process PRIMITIVE sheet
        primitive_processor = SheetProcessorFactory.create_processor('PRIMITIVE')
        if not primitive_processor:
            logger.error("Failed to create PRIMITIVE processor")
            return False
            
        source_primitive = workbook_handler.get_sheet('PRIMITIVE')
        if not source_primitive:
            logger.error("PRIMITIVE sheet not found in input workbook")
            return False
            
        target_primitive = workbook_handler.create_sheet('PRIMITIVE')
        if not target_primitive:
            logger.error("Failed to create PRIMITIVE sheet in output workbook")
            return False
            
        if not primitive_processor.process(source_primitive, target_primitive):
            return False
            
        # Process SCHEMA sheet
        schema_processor = SheetProcessorFactory.create_processor('SCHEMA', startup_analyzer)
        if not schema_processor:
            logger.error("Failed to create SCHEMA processor")
            return False
            
        source_schema = workbook_handler.get_sheet('SCHEMA')
        source_schema_formulas = workbook_handler.get_sheet('SCHEMA', 'formulas')
        if not source_schema or not source_schema_formulas:
            logger.error("SCHEMA sheet not found in input workbook")
            return False
            
        target_schema = workbook_handler.create_sheet('SCHEMA', 0)  # Make it first sheet
        if not target_schema:
            logger.error("Failed to create SCHEMA sheet in output workbook")
            return False
            
        if not schema_processor.process(source_schema, target_schema, source_schema_formulas):
            return False
            
        # Save output workbook
        return workbook_handler.save_workbook()
        
    except Exception as e:
        logger.exception(f"Migration failed: {str(e)}")
        return False
