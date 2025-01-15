import os
import openpyxl
from openai import OpenAI
from typing import Optional, Tuple, Dict, Any
from ..logger import get_logger

logger = get_logger()

class StartupDaysAnalyzer:
    def __init__(self, openai_key: Optional[str] = None):
        self.client = OpenAI(api_key=openai_key or os.getenv('OPENAI_API_KEY'))
        
    def extract_cell_references(self, formula: str) -> list[Tuple[str, str]]:
        """
        Extract sheet name and cell references from a formula.
        Example: "=PRIMITIVE!U25*PRIMITIVE!B12" -> [("PRIMITIVE", "U25"), ("PRIMITIVE", "B12")]
        """
        try:
            # Remove the leading = if present
            if formula.startswith('='):
                formula = formula[1:]
            
            refs = []
            parts = formula.split('*')  # Split by multiplication operator
            
            for part in parts:
                if '!' in part:  # Contains sheet reference
                    sheet, cell = part.split('!')
                    refs.append((sheet, cell))
                    logger.debug(f"Extracted reference: Sheet={sheet}, Cell={cell}")
            
            return refs
        except Exception as e:
            logger.error(f"Error extracting cell references: {str(e)}")
            return []

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

    def analyze_startup_days(self, formula: str, primitive_sheet_formulas, primitive_sheet_data, current_row: int, sheet: openpyxl.worksheet.worksheet.Worksheet) -> Optional[float]:
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
        if cell.font.b:  # Bold font indicates element catalog row
            # Skip analyzing element catalog rows (with SUM formulas)
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
            refs = self.extract_cell_references(formula)
            if refs:
                sheet_name, cell_ref = refs[0]
                if sheet_name.upper() == 'PRIMITIVE':
                    # Convert column letter and row number to cell coordinates
                    col = openpyxl.utils.column_index_from_string(cell_ref[0])
                    row = int(cell_ref[1:])
                    
                    # Try to get the value from both sheets
                    formula_value = primitive_sheet_formulas.cell(row=row, column=col).value
                    data_value = primitive_sheet_data.cell(row=row, column=col).value
                    
                    logger.info(f"Found values - Formula: {formula_value}, Data: {data_value}")
                    
                    # Prefer the data value if available
                    if data_value is not None and isinstance(data_value, (int, float)):
                        logger.info(f"Using data value: {data_value}")
                        return float(data_value)
                    elif formula_value is not None and isinstance(formula_value, (int, float)):
                        logger.info(f"Using formula value: {formula_value}")
                        return float(formula_value)
            
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
