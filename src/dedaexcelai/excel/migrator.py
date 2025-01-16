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
        # Initialize startup analyzer
        logger.info("Initializing StartupDaysAnalyzer...")
        startup_analyzer = StartupDaysAnalyzer(openai_key)
        if startup_analyzer.client:
            logger.info("Successfully initialized StartupDaysAnalyzer with GPT-4")
        
        # Initialize workbook handler
        workbook_handler = WorkbookHandler(input_file, output_file)
        
        # Load workbooks and create output
        if not workbook_handler.load_workbooks():
            return False
            
        if not workbook_handler.create_output_workbook():
            return False
            
        # Create processor factory
        processor_factory = SheetProcessorFactory()
        
        # Process PRIMITIVE sheet
        logger.info("Creating PRIMITIVE processor...")
        primitive_processor = processor_factory.create_primitive_processor()
        if not primitive_processor:
            logger.error("Failed to create PRIMITIVE processor")
            return False
            
        source_primitive = workbook_handler.get_sheet('PRIMITIVE')
        source_primitive_formulas = workbook_handler.get_sheet('PRIMITIVE', data_only=False)  # Get formulas version
        if not source_primitive or not source_primitive_formulas:
            logger.error("PRIMITIVE sheet not found in input workbook")
            return False
            
        target_primitive = workbook_handler.create_sheet('PRIMITIVE')
        if not target_primitive:
            logger.error("Failed to create PRIMITIVE sheet in output workbook")
            return False
            
        if not primitive_processor.process(source_primitive, target_primitive):
            return False
            
        # Process SCHEMA sheet
        logger.info("Creating SCHEMA processor with startup analyzer...")
        schema_processor = processor_factory.create_schema_processor(
            startup_analyzer=startup_analyzer,
            filename=workbook_handler.filename
        )
        if not schema_processor:
            logger.error("Failed to create SCHEMA processor")
            return False
            
        # Set primitive sheets after creation
        if hasattr(schema_processor, 'primitive_data'):
            schema_processor.primitive_data = source_primitive
        if hasattr(schema_processor, 'primitive_formulas'):
            schema_processor.primitive_formulas = source_primitive_formulas
            
        source_schema = workbook_handler.get_sheet('SCHEMA')
        source_schema_formulas = workbook_handler.get_sheet('SCHEMA', 'formulas')
        if not source_schema or not source_schema_formulas:
            logger.error("SCHEMA sheet not found in input workbook")
            return False
            
        target_schema = workbook_handler.create_sheet('SCHEMA', 0)  # Make it first sheet
        if not target_schema:
            logger.error("Failed to create SCHEMA sheet in output workbook")
            return False
            
        logger.info("Processing SCHEMA sheet with startup analyzer...")
        if not schema_processor.process(source_schema, target_schema, source_schema_formulas):
            return False
            
        # Save output workbook
        return workbook_handler.save_workbook()
        
    except Exception as e:
        logger.exception(f"Migration failed: {str(e)}")
        return False
