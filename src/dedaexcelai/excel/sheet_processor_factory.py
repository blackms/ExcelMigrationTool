from typing import Optional
from .sheet_processor import SheetProcessor, PrimitiveSheetProcessor, SchemaSheetProcessor
from ..llm import StartupDaysAnalyzer

class SheetProcessorFactory:
    """Factory for creating sheet processors."""
    
    @staticmethod
    def create_processor(sheet_name: str, startup_analyzer: Optional[StartupDaysAnalyzer] = None) -> Optional[SheetProcessor]:
        """Create appropriate processor for given sheet name."""
        processors = {
            'PRIMITIVE': PrimitiveSheetProcessor(),
            'SCHEMA': SchemaSheetProcessor(startup_analyzer)
        }
        return processors.get(sheet_name)
