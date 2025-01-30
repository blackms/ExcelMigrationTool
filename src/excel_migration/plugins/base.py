"""Base implementations for Excel migration plugins."""
import re
from datetime import datetime, date
from typing import Any, Dict, List, Optional
from .interfaces import FormulaExecutor, TransformationHandler

class DateDiffExecutor:
    """Execute DATEDIF formulas."""
    
    formula_type = "DATEDIF"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this is a DATEDIF formula."""
        return formula.startswith("DATEDIF(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Calculate difference between dates."""
        match = re.match(r"DATEDIF\(\[([^\]]+)\], TODAY\(\), '([^']+)'\)", formula)
        if not match:
            return 0
        
        field_name, unit = match.groups()
        date_value = values.get(field_name)
        if not date_value:
            return 0
        
        # Parse date
        if isinstance(date_value, str):
            try:
                date_value = datetime.strptime(date_value, "%Y-%m-%d").date()
            except ValueError:
                return 0
        elif isinstance(date_value, datetime):
            date_value = date_value.date()
        elif not isinstance(date_value, date):
            return 0
        
        # Calculate difference
        diff = (date.today() - date_value).days
        if unit.upper() == 'D':
            return diff
        elif unit.upper() == 'M':
            return diff // 30
        elif unit.upper() == 'Y':
            return diff // 365
        return diff

class CountExecutor:
    """Execute COUNT formulas."""
    
    formula_type = "COUNT"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this is a COUNT formula."""
        return formula.startswith("COUNT(") and not formula.startswith("COUNT_IF(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Calculate count of values."""
        match = re.match(r"COUNT\(\[([^\]]+)\]\)", formula)
        if not match:
            return 0
        
        field_name = match.group(1)
        value = values.get(field_name)
        
        if isinstance(value, list):
            return len(value)
        return 1 if value is not None else 0

class CountIfExecutor:
    """Execute COUNT_IF formulas."""
    
    formula_type = "COUNT_IF"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this is a COUNT_IF formula."""
        return formula.startswith("COUNT_IF(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Calculate count of values matching a condition."""
        match = re.match(r"COUNT_IF\(\[([^\]]+)\], '([^']+)'\)", formula)
        if not match:
            return 0
        
        field_name, condition = match.groups()
        value = values.get(field_name)
        
        if isinstance(value, list):
            return sum(1 for v in value if str(v) == condition)
        return 1 if str(value) == condition else 0

class SumExecutor:
    """Execute SUM formulas."""
    
    formula_type = "SUM"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this is a SUM formula."""
        return formula.startswith("SUM(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Calculate sum of values."""
        match = re.match(r"SUM\(\[([^\]]+)\]\)", formula)
        if not match:
            return 0.0
        
        field_name = match.group(1)
        value = values.get(field_name)
        
        if isinstance(value, list):
            return sum(float(v) for v in value if v is not None)
        return float(value) if value is not None else 0.0

class AverageExecutor:
    """Execute AVERAGE formulas."""
    
    formula_type = "AVERAGE"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this is an AVERAGE formula."""
        return formula.startswith("AVERAGE(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Calculate average of values."""
        match = re.match(r"AVERAGE\(\[([^\]]+)\]\)", formula)
        if not match:
            return 0.0
        
        field_name = match.group(1)
        value = values.get(field_name)
        
        if isinstance(value, list):
            valid_values = [float(v) for v in value if v is not None]
            return sum(valid_values) / len(valid_values) if valid_values else 0.0
        return float(value) if value is not None else 0.0

class DateTimeTransformer:
    """Transform datetime values."""
    
    transformation_type = "datetime_format"
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        return transformation.get("type") == self.transformation_type
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Format a datetime value."""
        format_str = params.get("format", "%Y-%m-%d")
        
        if isinstance(value, (datetime, date)):
            dt = value
        else:
            # Try parsing with provided formats
            formats = params.get("input_formats", ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y"])
            for fmt in formats:
                try:
                    dt = datetime.strptime(str(value), fmt)
                    break
                except ValueError:
                    continue
            else:
                return str(value)
        
        return dt.strftime(format_str)

class NumericTransformer:
    """Transform numeric values."""
    
    transformation_type = "numeric_format"
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        return transformation.get("type") == self.transformation_type
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Format a numeric value."""
        try:
            num = float(value)
            decimal_places = params.get("decimal_places", 2)
            thousands_separator = params.get("thousands_separator", True)
            
            if thousands_separator:
                return f"{num:,.{decimal_places}f}"
            return f"{num:.{decimal_places}f}"
        except (ValueError, TypeError):
            return str(value)

class BooleanTransformer:
    """Transform boolean values."""
    
    transformation_type = "boolean_transform"
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        return transformation.get("type") == self.transformation_type
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Transform a value to boolean."""
        str_value = str(value).lower()
        true_values = [v.lower() for v in params.get("true_values", [])]
        false_values = [v.lower() for v in params.get("false_values", [])]
        
        if str_value in true_values:
            return True
        if str_value in false_values:
            return False
        return str_value == "true"

class ConcatenateTransformer:
    """Transform multiple values by concatenation."""
    
    transformation_type = "concatenate"
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        return transformation.get("type") == self.transformation_type
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Concatenate multiple values."""
        if not isinstance(value, list):
            return str(value)
        
        separator = params.get("separator", " ")
        return separator.join(str(v) for v in value)