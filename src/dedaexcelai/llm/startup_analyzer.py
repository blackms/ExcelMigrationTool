import os
import openpyxl
from openai import OpenAI
from typing import Optional, Tuple, Dict, Any
from ..logger import get_logger
from ..excel.cell_operations import (
    get_cell_value, 
    is_empty_or_dashes, 
    extract_cell_references,
    get_cell_value_with_fallback
)

logger = get_logger()

class StartupDaysAnalyzer:
    def __init__(self, openai_key: Optional[str] = None):
        self.client = OpenAI(api_key=openai_key or os.getenv('OPENAI_API_KEY'))
        
    def get_primitive_context(self, primitive_sheet) -> Dict[str, Any]:
        """
        Extract relevant context from the PRIMITIVE sheet to help GPT understand its structure.
        """
        context = {
            'headers': [],
            'startup_related_cells': []
        }
        
        # Scan first few rows for headers
        for row in range(1, min(20, primitive_sheet.max_row + 1)):
            for col in range(1, primitive_sheet.max_column + 1):
                cell = primitive_sheet.cell(row=row, column=col)
                value = cell.value
                if value:
                    col_letter = openpyxl.utils.get_column_letter(col)
                    context['headers'].append({
                        'cell': f'{col_letter}{row}',
                        'value': str(value)
                    })
                    # Look for cells that might be related to startup/setup days
                    if any(keyword in str(value).lower() for keyword in ['gg', 'giorni', 'setup', 'startup']):
                        context['startup_related_cells'].append({
                            'cell': f'{col_letter}{row}',
                            'value': str(value)
                        })
        
        return context

    def analyze_startup_days(self, formula: str, primitive_sheet_formulas, primitive_sheet_data, current_row: int, sheet: openpyxl.worksheet.worksheet.Worksheet, input_sheet_formulas: openpyxl.worksheet.worksheet.Worksheet) -> Optional[float]:
        """
        Analyze a formula from column H to determine startup days by looking up values in PRIMITIVE sheet.
        For element catalog rows (in bold), looks at formulas from rows below until the separator.
        
        Args:
            formula: The formula from column H (e.g., "=PRIMITIVE!U25*PRIMITIVE!B12")
            primitive_sheet_formulas: The PRIMITIVE worksheet object with formulas
            primitive_sheet_data: The PRIMITIVE worksheet object with values
            current_row: The current row being analyzed
            sheet: The worksheet containing the element catalog
            
        Returns:
            Optional[float]: The startup days value if found, None otherwise
        """
        from ..excel.structure_analyzer import find_element_catalog_interval
        
        # Check if this is an element catalog row (in bold)
        cell = sheet.cell(row=current_row, column=2)  # Service Element column
        cell_value = get_cell_value(cell)
        logger.debug(f"Analyzing row {current_row}: '{cell_value}' (bold: {cell.font and cell.font.b})")
        
        if cell.font and cell.font.b:  # Bold font indicates element catalog row
            logger.info(f"Found element catalog: '{cell_value}' at row {current_row}")
            # For element catalogs, look at formulas in sub-elements
            for r in range(current_row + 1, sheet.max_row + 1):
                sub_cell = sheet.cell(row=r, column=2)
                sub_value = get_cell_value(sub_cell)
                logger.debug(f"Checking sub-element at row {r}: '{sub_value}'")
                
                if sub_value and isinstance(sub_value, str) and "---" in sub_value:
                    logger.debug(f"Found separator at row {r}")
                    break
                    
                if sub_value and not is_empty_or_dashes(sub_value):
                    logger.info(f"Found sub-element: '{sub_value}' at row {r}")
                    # Found a sub-element, get its formula from the formulas workbook
                    sub_formula = input_sheet_formulas.cell(row=r, column=8).value  # Column H
                    logger.debug(f"Sub-element formula: {sub_formula}")
                    
                    if sub_formula and isinstance(sub_formula, str) and sub_formula.startswith('='):
                        result = self._analyze_single_formula(sub_formula, primitive_sheet_formulas, primitive_sheet_data)
                        if result is not None:
                            logger.info(f"Found startup days {result} from sub-element '{sub_value}'")
                            return result
                        else:
                            logger.debug(f"Could not extract startup days from sub-element '{sub_value}'")
            
            logger.warning(f"No startup days found in any sub-elements of '{cell_value}'")
            return None
        else:
            # For non-element catalog rows (sub-elements), analyze the formula directly
            return self._analyze_single_formula(formula, primitive_sheet_formulas, primitive_sheet_data)
            
    def _analyze_single_formula(self, formula: str, primitive_sheet_formulas, primitive_sheet_data) -> Optional[float]:
        """
        Analyze a single formula to determine startup days.
        """
        try:
            logger.info(f"Analyzing formula for startup days: {formula}")
            
            # First try direct formula extraction
            logger.debug(f"Analyzing formula: {formula}")
            refs = extract_cell_references(formula, logger)
            if refs:
                sheet_name, cell_ref = refs[0]
                if sheet_name.upper() == 'PRIMITIVE':
                    logger.debug(f"Found PRIMITIVE sheet reference: {cell_ref}")
                    # Convert column letter and row number to cell coordinates
                    col = openpyxl.utils.column_index_from_string(cell_ref[0])
                    row = int(cell_ref[1:])
                    logger.debug(f"Converted to coordinates: row={row}, col={col}")
                    
                    # Try to get the value from both sheets
                    formula_cell = primitive_sheet_formulas.cell(row=row, column=col)
                    data_cell = primitive_sheet_data.cell(row=row, column=col)
                    value = get_cell_value_with_fallback(formula_cell, data_cell, logger)
                    if value is not None:
                        return value
                else:
                    logger.debug(f"Sheet name '{sheet_name}' is not PRIMITIVE")
            
            # If direct extraction fails, use GPT-4 to analyze the sheet
            logger.info("Direct reference extraction failed, using GPT-4 for analysis")
            
            # Get context from both PRIMITIVE sheets
            context_formulas = self.get_primitive_context(primitive_sheet_formulas)
            context_data = self.get_primitive_context(primitive_sheet_data)
            
            # Construct prompt for GPT-4
            prompt = f"""
            Analizza questa formula Excel e il contesto del foglio PRIMITIVE per determinare i giorni di startup.

            Formula analizzata: {formula}

            Contesto del foglio PRIMITIVE (Formule):
            Headers trovati:
            {[f"{h['cell']}: {h['value']}" for h in context_formulas['headers']]}
            
            Celle relative allo startup:
            {[f"{c['cell']}: {c['value']}" for c in context_formulas['startup_related_cells']]}

            Contesto del foglio PRIMITIVE (Valori):
            Headers trovati:
            {[f"{h['cell']}: {h['value']}" for h in context_data['headers']]}
            
            Celle relative allo startup:
            {[f"{c['cell']}: {c['value']}" for c in context_data['startup_related_cells']]}

            La formula fa riferimento al foglio PRIMITIVE. Analizza:
            1. Il primo riferimento nella formula (es. U25 in PRIMITIVE!U25*PRIMITIVE!B12) solitamente contiene i giorni
            2. Cerca valori numerici che rappresentano giorni di setup/startup
            3. Se trovi più valori, scegli quello che sembra più appropriato basandoti sul contesto

            Rispondi SOLO con il numero di giorni (come numero decimale) o 'None' se non determinabile.
            Non includere spiegazioni o altro testo.
            """
            
            completion = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Sei un esperto di analisi di fogli Excel e formule."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2
            )
            
            response = completion.choices[0].message.content.strip()
            logger.info(f"GPT-4 analysis response: {response}")
            
            try:
                if response.lower() == 'none':
                    return None
                return float(response)
            except ValueError:
                logger.warning(f"Could not convert GPT-4 response to float: {response}")
                return None
                
        except Exception as e:
            logger.error(f"Error analyzing startup days: {str(e)}")
            return None
            
    def calculate_sum_range(self, sheet: openpyxl.worksheet.worksheet.Worksheet, row: int, column: int) -> str:
        """
        Calculate the sum range for an element catalog row.
        
        Args:
            sheet: The worksheet
            row: The current row (element catalog row)
            column: The column to sum (H for Fixed, I for Fee)
            
        Returns:
            str: Excel formula for summing the range (e.g., "=SUM(H6:H11)")
        """
        from ..excel.structure_analyzer import find_element_catalog_interval
        
        # Find the interval for this element catalog
        start_row, end_row = find_element_catalog_interval(sheet, row)
        
        # Only create sum if this is an element catalog row (in bold)
        cell = sheet.cell(row=row, column=2)  # Service Element column
        if cell.font.b:
            # Get column letter
            col_letter = openpyxl.utils.get_column_letter(column)
            
            # Create sum formula for the range below this row until the separator
            # Add 1 to start_row to skip the element catalog row itself
            return f"=SUM({col_letter}{start_row + 1}:{col_letter}{end_row})"
        
        return ""
