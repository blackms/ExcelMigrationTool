"""Rule execution engine."""
from typing import Dict, Any, List, Union
from loguru import logger
import openpyxl
from datetime import datetime, date
import re

class RuleExecutor:
    """Execute migration rules on Excel files."""
    
    def __init__(self):
        """Initialize the rule executor."""
        logger.debug("Initialized rule executor")

    async def execute(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a single rule."""
        try:
            rule_type = rule.get("type")
            if not rule_type:
                logger.error("Rule type not specified")
                return False
            
            if rule_type == "field_mapping":
                return await self._execute_field_mapping(rule, context)
            elif rule_type == "calculation":
                return await self._execute_calculation(rule, context)
            else:
                logger.error(f"Unknown rule type: {rule_type}")
                return False
            
        except Exception as e:
            logger.error(f"Rule execution failed: {str(e)}")
            return False

    async def _execute_field_mapping(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a field mapping rule."""
        try:
            source_field = rule.get("source_field")
            target_field = rule.get("target_field")
            transformation = rule.get("transformation", {})
            
            if not source_field or not target_field:
                logger.error("Source or target field not specified")
                return False
            
            # Handle multiple source fields
            if isinstance(source_field, list):
                source_values = []
                for field in source_field:
                    value = context.get("source_data", {}).get(field)
                    if value is None:
                        logger.error(f"Source field not found: {field}")
                        return False
                    source_values.append(value)
                transformed_value = self._apply_transformation(
                    source_values,
                    transformation
                )
            else:
                # Single source field
                source_value = context.get("source_data", {}).get(source_field)
                if source_value is None:
                    logger.error(f"Source field not found: {source_field}")
                    return False
                transformed_value = self._apply_transformation(
                    source_value,
                    transformation
                )
            
            # Update target
            context["target_data"][target_field] = transformed_value
            return True
            
        except Exception as e:
            logger.error(f"Field mapping failed: {str(e)}")
            return False

    async def _execute_calculation(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a calculation rule."""
        try:
            target_field = rule.get("target_field")
            formula = rule.get("formula")
            source_fields = rule.get("source_fields", [])
            
            if not target_field or not formula:
                logger.error("Target field or formula not specified")
                return False
            
            # Get source values
            values = {}
            for field in source_fields:
                value = context.get("source_data", {}).get(field)
                if value is None:
                    logger.error(f"Source field not found: {field}")
                    return False
                values[field] = value
            
            # Execute formula
            result = self._execute_formula(formula, values)
            
            # Update target
            context["target_data"][target_field] = result
            return True
            
        except Exception as e:
            logger.error(f"Calculation failed: {str(e)}")
            return False

    def _apply_transformation(self, value: Union[Any, List[Any]], transformation: Dict[str, Any]) -> Any:
        """Apply a transformation to a value."""
        try:
            trans_type = transformation.get("type", "direct")
            params = transformation.get("params", {})
            
            if trans_type == "direct":
                return value
            elif trans_type == "datetime_format":
                return self._format_datetime(value, params.get("format"))
            elif trans_type == "numeric_format":
                return self._format_numeric(
                    value,
                    params.get("decimal_places", 2),
                    params.get("thousands_separator", True)
                )
            elif trans_type == "concatenate":
                if not isinstance(value, list):
                    logger.error("Concatenation requires multiple values")
                    return str(value)
                return params.get("separator", " ").join(str(v) for v in value)
            elif trans_type == "boolean_transform":
                return self._transform_boolean(
                    value,
                    params.get("true_values", []),
                    params.get("false_values", [])
                )
            else:
                logger.warning(f"Unknown transformation type: {trans_type}")
                return value
            
        except Exception as e:
            logger.error(f"Transformation failed: {str(e)}")
            return value

    def _format_datetime(self, value: Any, format_str: str) -> str:
        """Format a datetime value."""
        try:
            if isinstance(value, (datetime, date)):
                dt = value
            else:
                # Try parsing common formats
                for fmt in ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y"]:
                    try:
                        dt = datetime.strptime(str(value), fmt)
                        break
                    except ValueError:
                        continue
                else:
                    return str(value)
            
            return dt.strftime(format_str or "%Y-%m-%d")
            
        except Exception as e:
            logger.error(f"Datetime formatting failed: {str(e)}")
            return str(value)

    def _format_numeric(self, value: Any, decimal_places: int, thousands_separator: bool) -> str:
        """Format a numeric value."""
        try:
            num = float(value)
            if thousands_separator:
                return f"{num:,.{decimal_places}f}"
            return f"{num:.{decimal_places}f}"
            
        except (ValueError, TypeError):
            return str(value)

    def _transform_boolean(self, value: Any, true_values: List[str], false_values: List[str]) -> bool:
        """Transform a value to boolean based on lists of true/false values."""
        str_value = str(value).lower()
        if str_value in [v.lower() for v in true_values]:
            return True
        if str_value in [v.lower() for v in false_values]:
            return False
        # Default to string comparison with "true"
        return str_value == "true"

    def _execute_formula(self, formula: str, values: Dict[str, Any]) -> Any:
        """Execute a formula with provided values."""
        try:
            # Handle special functions
            if formula.startswith("DATEDIF("):
                return self._calc_date_diff(formula, values)
            elif formula.startswith("COUNT("):
                return self._calc_count(formula, values)
            elif formula.startswith("COUNT_IF("):
                return self._calc_count_if(formula, values)
            elif formula.startswith("SUM("):
                return self._calc_sum(formula, values)
            elif formula.startswith("AVERAGE("):
                return self._calc_average(formula, values)
            
            # Replace field references with values
            for field, value in values.items():
                formula = formula.replace(f"[{field}]", str(value))
            
            # Evaluate the formula
            # Note: In a production environment, you would want to use a safer
            # evaluation method or a proper expression parser
            result = eval(formula)
            return result
            
        except Exception as e:
            logger.error(f"Formula execution failed: {str(e)}")
            return None

    def _calc_date_diff(self, formula: str, values: Dict[str, Any]) -> int:
        """Calculate difference between dates."""
        try:
            # Extract field and unit from DATEDIF([field], TODAY(), 'unit')
            match = re.match(r"DATEDIF\(\[([^\]]+)\], TODAY\(\), '([^']+)'\)", formula)
            if not match:
                return 0
            
            field_name, unit = match.groups()
            date_value = values.get(field_name)
            if not date_value:
                return 0
            
            # Parse date
            if isinstance(date_value, str):
                date_value = datetime.strptime(date_value, "%Y-%m-%d").date()
            elif isinstance(date_value, datetime):
                date_value = date_value.date()
            
            # Calculate difference
            diff = (date.today() - date_value).days
            if unit.upper() == 'D':
                return diff
            elif unit.upper() == 'M':
                return diff // 30
            elif unit.upper() == 'Y':
                return diff // 365
            return diff
            
        except Exception as e:
            logger.error(f"Date difference calculation failed: {str(e)}")
            return 0

    def _calc_count(self, formula: str, values: Dict[str, Any]) -> int:
        """Calculate count of values."""
        try:
            # Extract field from COUNT([field])
            match = re.match(r"COUNT\(\[([^\]]+)\]\)", formula)
            if not match:
                return 0
            
            field_name = match.group(1)
            if field_name == "TransactionID":
                transactions = values.get("Transactions", [])
                return len(transactions)
            
            value = values.get(field_name)
            if isinstance(value, list):
                return len(value)
            return 1 if value is not None else 0
            
        except Exception as e:
            logger.error(f"Count calculation failed: {str(e)}")
            return 0

    def _calc_count_if(self, formula: str, values: Dict[str, Any]) -> int:
        """Calculate count of values matching a condition."""
        try:
            # Extract field and condition from COUNT_IF([field], 'value')
            match = re.match(r"COUNT_IF\(\[([^\]]+)\], '([^']+)'\)", formula)
            if not match:
                return 0
            
            field_name, condition = match.groups()
            if "Transactions" in values:
                transactions = values["Transactions"]
                return sum(1 for t in transactions if str(t.get(field_name)) == condition)
            
            value = values.get(field_name)
            if isinstance(value, list):
                return sum(1 for v in value if str(v) == condition)
            return 1 if str(value) == condition else 0
            
        except Exception as e:
            logger.error(f"Conditional count calculation failed: {str(e)}")
            return 0

    def _calc_sum(self, formula: str, values: Dict[str, Any]) -> float:
        """Calculate sum of values."""
        try:
            # Extract field from SUM([field])
            match = re.match(r"SUM\(\[([^\]]+)\]\)", formula)
            if not match:
                return 0.0
            
            field_name = match.group(1)
            if "Transactions" in values:
                transactions = values["Transactions"]
                return sum(float(t.get(field_name, 0)) for t in transactions)
            
            value = values.get(field_name)
            if isinstance(value, list):
                return sum(float(v) for v in value if v is not None)
            return float(value) if value is not None else 0.0
            
        except Exception as e:
            logger.error(f"Sum calculation failed: {str(e)}")
            return 0.0

    def _calc_average(self, formula: str, values: Dict[str, Any]) -> float:
        """Calculate average of values."""
        try:
            # Extract field from AVERAGE([field])
            match = re.match(r"AVERAGE\(\[([^\]]+)\]\)", formula)
            if not match:
                return 0.0
            
            field_name = match.group(1)
            if "Transactions" in values:
                transactions = values["Transactions"]
                amounts = [float(t.get(field_name, 0)) for t in transactions]
                return sum(amounts) / len(amounts) if amounts else 0.0
            
            value = values.get(field_name)
            if isinstance(value, list):
                valid_values = [float(v) for v in value if v is not None]
                return sum(valid_values) / len(valid_values) if valid_values else 0.0
            return float(value) if value is not None else 0.0
            
        except Exception as e:
            logger.error(f"Average calculation failed: {str(e)}")
            return 0.0