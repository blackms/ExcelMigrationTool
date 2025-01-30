"""Rule execution engine."""
from typing import Dict, Any, List, Union
from loguru import logger
from ..core.interfaces import RuleExecutor as RuleExecutorInterface
from ..plugins.interfaces import PluginRegistry, FormulaExecutor, TransformationHandler
from ..plugins.base import (
    DateDiffExecutor,
    CountExecutor,
    CountIfExecutor,
    SumExecutor,
    AverageExecutor,
    DateTimeTransformer,
    NumericTransformer,
    BooleanTransformer,
    ConcatenateTransformer
)

class RuleExecutor(RuleExecutorInterface):
    """Execute migration rules on Excel files."""
    
    def __init__(self):
        """Initialize the rule executor."""
        self.registry = PluginRegistry()
        self._register_default_plugins()
        logger.debug("Initialized rule executor with default plugins")
    
    def _register_default_plugins(self):
        """Register default formula executors and transformation handlers."""
        # Register formula executors
        formula_executors = [
            DateDiffExecutor(),
            CountExecutor(),
            CountIfExecutor(),
            SumExecutor(),
            AverageExecutor()
        ]
        for executor in formula_executors:
            self.registry.register_formula_executor(executor)
        
        # Register transformation handlers
        transformation_handlers = [
            DateTimeTransformer(),
            NumericTransformer(),
            BooleanTransformer(),
            ConcatenateTransformer()
        ]
        for handler in transformation_handlers:
            self.registry.register_transformation_handler(handler)
    
    async def validate_rule(self, rule: Dict[str, Any]) -> bool:
        """Validate a rule's structure and requirements."""
        if not isinstance(rule, dict):
            logger.error("Rule must be a dictionary")
            return False
        
        rule_type = rule.get("type")
        if not rule_type:
            logger.error("Rule type not specified")
            return False
        
        if rule_type == "field_mapping":
            return self._validate_field_mapping(rule)
        elif rule_type == "calculation":
            return self._validate_calculation(rule)
        else:
            logger.error(f"Unknown rule type: {rule_type}")
            return False
    
    def _validate_field_mapping(self, rule: Dict[str, Any]) -> bool:
        """Validate a field mapping rule."""
        required_fields = ["source_field", "target_field"]
        return all(field in rule for field in required_fields)
    
    def _validate_calculation(self, rule: Dict[str, Any]) -> bool:
        """Validate a calculation rule."""
        required_fields = ["target_field", "formula"]
        return all(field in rule for field in required_fields)
    
    async def execute(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a single rule."""
        try:
            if not await self.validate_rule(rule):
                return False
            
            rule_type = rule["type"]
            if rule_type == "field_mapping":
                return await self._execute_field_mapping(rule, context)
            elif rule_type == "calculation":
                return await self._execute_calculation(rule, context)
            
            return False
            
        except Exception as e:
            logger.error(f"Rule execution failed: {str(e)}")
            return False
    
    async def _execute_field_mapping(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a field mapping rule."""
        try:
            source_field = rule["source_field"]
            target_field = rule["target_field"]
            transformation = rule.get("transformation", {})
            
            # Handle multiple source fields
            if isinstance(source_field, list):
                source_values = []
                for field in source_field:
                    value = context.get("source_data", {}).get(field)
                    if value is None:
                        logger.error(f"Source field not found: {field}")
                        return False
                    source_values.append(value)
                value_to_transform = source_values
            else:
                # Single source field
                value_to_transform = context.get("source_data", {}).get(source_field)
                if value_to_transform is None:
                    logger.error(f"Source field not found: {source_field}")
                    return False
            
            # Apply transformation
            transformed_value = self._apply_transformation(value_to_transform, transformation)
            
            # Update target
            context["target_data"][target_field] = transformed_value
            return True
            
        except Exception as e:
            logger.error(f"Field mapping failed: {str(e)}")
            return False
    
    async def _execute_calculation(self, rule: Dict[str, Any], context: Dict[str, Any]) -> bool:
        """Execute a calculation rule."""
        try:
            target_field = rule["target_field"]
            formula = rule["formula"]
            source_fields = rule.get("source_fields", [])
            
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
            if result is None:
                return False
            
            # Update target
            context["target_data"][target_field] = result
            return True
            
        except Exception as e:
            logger.error(f"Calculation failed: {str(e)}")
            return False
    
    def _apply_transformation(self, value: Any, transformation: Dict[str, Any]) -> Any:
        """Apply a transformation to a value."""
        try:
            if not transformation:
                return value
            
            trans_type = transformation.get("type", "direct")
            if trans_type == "direct":
                return value
            
            handler = self.registry.get_transformation_handler(trans_type)
            if handler:
                return handler.transform(value, transformation.get("params", {}))
            
            logger.warning(f"No handler found for transformation type: {trans_type}")
            return value
            
        except Exception as e:
            logger.error(f"Transformation failed: {str(e)}")
            return value
    
    def _execute_formula(self, formula: str, values: Dict[str, Any]) -> Any:
        """Execute a formula with provided values."""
        try:
            executor = self.registry.get_formula_executor(formula)
            if executor:
                return executor.execute(formula, values)
            
            # For simple arithmetic formulas, replace field references with values
            for field, value in values.items():
                formula = formula.replace(f"[{field}]", str(value))
            
            # Evaluate the formula
            # Note: In a production environment, you would want to use a safer
            # evaluation method or a proper expression parser
            return eval(formula)
            
        except Exception as e:
            logger.error(f"Formula execution failed: {str(e)}")
            return None