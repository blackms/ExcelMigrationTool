import openpyxl
from typing import Optional
from pathlib import Path
import tempfile
import shutil
import os

from ..logger import get_logger

logger = get_logger()

class WorkbookHandler:
    """Handles workbook operations like loading, saving, and sheet management."""
    
    def __init__(self, input_file: str, output_file: str):
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.input_wb_data = None
        self.input_wb_formulas = None
        self.output_wb = None
        
    def load_workbooks(self) -> bool:
        """Load input workbooks with and without data_only."""
        try:
            logger.info("Loading input workbook...")
            self.input_wb_data = openpyxl.load_workbook(self.input_file, data_only=True, keep_links=False)
            self.input_wb_formulas = openpyxl.load_workbook(self.input_file, data_only=False, keep_links=False)
            return True
        except Exception as e:
            logger.error(f"Failed to load workbooks: {str(e)}")
            return False
            
    def create_output_workbook(self) -> bool:
        """Create new output workbook."""
        try:
            logger.info("Creating new workbook...")
            self.output_wb = openpyxl.Workbook()
            if 'Sheet' in self.output_wb.sheetnames:
                del self.output_wb['Sheet']
            return True
        except Exception as e:
            logger.error(f"Failed to create output workbook: {str(e)}")
            return False
            
    def save_workbook(self) -> bool:
        """Save output workbook with temporary file handling."""
        logger.info("Saving output workbook...")
        try:
            # Create temp file in same directory as output file
            temp_dir = self.output_file.parent
            temp_fd, temp_path = tempfile.mkstemp(dir=temp_dir, suffix='.xlsx')
            os.close(temp_fd)
            
            # Save to temp file
            self.output_wb.save(temp_path)
            logger.debug(f"Saved to temporary file: {temp_path}")
            
            # Replace output file
            if self.output_file.exists():
                self.output_file.unlink()
            shutil.move(temp_path, self.output_file)
            logger.info(f"Successfully saved to: {self.output_file}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to save workbook: {str(e)}")
            if temp_path:
                try:
                    os.remove(temp_path)
                except:
                    pass
            return False
            
    def get_sheet(self, name: str, workbook: str = 'data') -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
        """Get sheet from specified workbook."""
        wb = getattr(self, f'input_wb_{workbook}')
        if not wb or name not in wb.sheetnames:
            return None
        return wb[name]
        
    def create_sheet(self, name: str, index: Optional[int] = None) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
        """Create new sheet in output workbook."""
        try:
            if index is not None:
                return self.output_wb.create_sheet(name, index)
            return self.output_wb.create_sheet(name)
        except Exception as e:
            logger.error(f"Failed to create sheet {name}: {str(e)}")
            return None
