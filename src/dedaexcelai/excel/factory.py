"""Factory for creating sheet processors."""
from typing import Optional
from .processors import SheetProcessor, PrimitiveSheetProcessor, SchemaSheetProcessor

def create_sheet_processor(sheet_name: str, startup_analyzer=None, filename: str = "") -> Optional[SheetProcessor]:
    """Create appropriate sheet processor based on sheet name."""
    if sheet_name == "PRIMITIVE":
        return PrimitiveSheetProcessor()
    elif sheet_name == "SCHEMA":
        return SchemaSheetProcessor(startup_analyzer, filename)
    return None
