from typing import Optional
from .sheet_processor import SheetProcessor, PrimitiveSheetProcessor, SchemaSheetProcessor

class SheetProcessorFactory:
    """Factory for creating sheet processors."""
    
    def create_primitive_processor(self) -> SheetProcessor:
        """Create PRIMITIVE sheet processor."""
        return PrimitiveSheetProcessor()
    
    def create_schema_processor(self, startup_analyzer=None, filename: str = "") -> Optional[SchemaSheetProcessor]:
        """Create a processor for SCHEMA sheet."""
        return SchemaSheetProcessor(startup_analyzer, filename)
