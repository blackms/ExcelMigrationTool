"""Rule execution engine."""
from typing import Dict, Any
from loguru import logger
import openpyxl
from datetime import datetime

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
            
            # Get source value
            source_value = context.get("source_data", {}).get(source_field)
            if source_value is None:
                logger.error(f"Source field not found: {source_field}")
                return False
            
            # Apply transformation
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

    def _apply_transformation(self, value: Any, transformation: Dict[str, Any]) -> Any:
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
            else:
                logger.warning(f"Unknown transformation type: {trans_type}")
                return value
            
        except Exception as e:
            logger.error(f"Transformation failed: {str(e)}")
            return value

    def _format_datetime(self, value: Any, format_str: str) -> str:
        """Format a datetime value."""
        try:
            if isinstance(value, datetime):
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

    def _execute_formula(self, formula: str, values: Dict[str, Any]) -> Any:
        """Execute a formula with provided values."""
        try:
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