"""Rule engine for Excel migrations."""
from typing import Any, Dict, List, Optional
import json
from pathlib import Path
import logging
import asyncio

from ..core.models import MigrationRule, RuleType, ValidationResult
from ..llm.chain import LLMProvider, ChainManager, ExcelProcessor, ProcessingCallback

logger = logging.getLogger(__name__)

class RuleEngine:
    """Engine for managing and executing migration rules."""
    
    def __init__(self, llm_provider: str = "openai", **llm_kwargs):
        """Initialize the rule engine."""
        callbacks = [ProcessingCallback()]
        llm_kwargs['callbacks'] = callbacks
        
        self.llm = LLMProvider.create_llm(llm_provider, **llm_kwargs)
        self.chain_manager = ChainManager(self.llm)
        self.processor = ExcelProcessor(self.chain_manager)
        
    def load_rules(self, rules_file: str) -> List[MigrationRule]:
        """Load rules from a JSON configuration file."""
        try:
            with open(rules_file, 'r') as f:
                rules_data = json.load(f)
            
            rules = []
            for rule_data in rules_data['rules']:
                rule = MigrationRule(
                    rule_type=RuleType[rule_data['type'].upper()],
                    source_columns=rule_data['source_columns'],
                    target_column=rule_data['target_column'],
                    conditions=rule_data.get('conditions'),
                    transformation=rule_data.get('transformation'),
                    llm_prompt=rule_data.get('llm_prompt')
                )
                rules.append(rule)
            
            return rules
            
        except Exception as e:
            logger.error(f"Failed to load rules from {rules_file}: {str(e)}")
            raise

    async def execute_rule(self, rule: MigrationRule, 
                         source_values: Dict[str, Any],
                         context: Optional[Dict[str, Any]] = None) -> Any:
        """Execute a single migration rule."""
        try:
            # Check conditions first
            if not self._check_conditions(rule.conditions, source_values, context):
                return None

            # Execute based on rule type
            if rule.rule_type == RuleType.COPY:
                return await self._execute_copy(rule, source_values)
            elif rule.rule_type == RuleType.TRANSFORM:
                return await self._execute_transform(rule, source_values, context)
            elif rule.rule_type == RuleType.COMPUTE:
                return await self._execute_compute(rule, source_values, context)
            elif rule.rule_type == RuleType.AGGREGATE:
                return await self._execute_aggregate(rule, source_values, context)
            elif rule.rule_type == RuleType.VALIDATE:
                return await self._execute_validate(rule, source_values, context)
            else:
                raise ValueError(f"Unsupported rule type: {rule.rule_type}")

        except Exception as e:
            logger.error(f"Failed to execute rule: {str(e)}")
            return None

    def _check_conditions(self, conditions: Optional[Dict[str, Any]],
                        source_values: Dict[str, Any],
                        context: Optional[Dict[str, Any]]) -> bool:
        """Check if conditions are met for rule execution."""
        if not conditions:
            return True

        try:
            for field, condition in conditions.items():
                if field not in source_values:
                    return False

                value = source_values[field]
                
                # Handle different condition types
                if isinstance(condition, dict):
                    operator = condition.get('operator', '==')
                    target = condition.get('value')
                    
                    if operator == '==':
                        if value != target:
                            return False
                    elif operator == '!=':
                        if value == target:
                            return False
                    elif operator == '>':
                        if not value > target:
                            return False
                    elif operator == '<':
                        if not value < target:
                            return False
                    elif operator == 'in':
                        if value not in target:
                            return False
                    elif operator == 'contains':
                        if target not in str(value):
                            return False
                else:
                    # Simple equality check
                    if value != condition:
                        return False

            return True

        except Exception as e:
            logger.error(f"Error checking conditions: {str(e)}")
            return False

    async def _execute_copy(self, rule: MigrationRule,
                          source_values: Dict[str, Any]) -> Any:
        """Execute a copy rule."""
        if len(rule.source_columns) != 1:
            raise ValueError("Copy rule must have exactly one source column")
        
        source_col = rule.source_columns[0]
        return source_values.get(source_col)

    async def _execute_transform(self, rule: MigrationRule,
                               source_values: Dict[str, Any],
                               context: Optional[Dict[str, Any]]) -> Any:
        """Execute a transform rule."""
        if not rule.transformation and not rule.llm_prompt:
            raise ValueError("Transform rule must have either transformation or llm_prompt")

        if rule.llm_prompt:
            # Use LLM for complex transformations
            transformation_rules = {
                "prompt": rule.llm_prompt,
                "steps": [{"type": "transform", "description": rule.llm_prompt}]
            }
            return await self.processor.process_transformation(
                source_values,
                transformation_rules,
                context
            )
        else:
            # Use predefined transformation
            # This could be extended to support different transformation types
            return eval(rule.transformation, {"__builtins__": {}}, source_values)

    async def _execute_compute(self, rule: MigrationRule,
                             source_values: Dict[str, Any],
                             context: Optional[Dict[str, Any]]) -> Any:
        """Execute a compute rule."""
        if not rule.transformation:
            raise ValueError("Compute rule must have a transformation")

        # For complex computations, use the formula analysis agent
        if self._is_complex_computation(rule.transformation):
            formula_agent = self.processor.get_or_create_agent("formula")
            result = await formula_agent.process_task(
                f"Compute result using formula: {rule.transformation}",
                {"values": source_values, **(context or {})}
            )
            return result

        # For simple computations, use direct evaluation
        compute_context = {
            **source_values,
            "sum": sum,
            "len": len,
            "min": min,
            "max": max,
            "round": round
        }
        
        return eval(rule.transformation, {"__builtins__": {}}, compute_context)

    async def _execute_aggregate(self, rule: MigrationRule,
                               source_values: Dict[str, Any],
                               context: Optional[Dict[str, Any]]) -> Any:
        """Execute an aggregate rule."""
        if not rule.transformation:
            raise ValueError("Aggregate rule must have a transformation")

        # For complex aggregations, use the transformation agent
        if self._is_complex_aggregation(rule.transformation):
            transformation_agent = self.processor.get_or_create_agent("transformation")
            result = await transformation_agent.process_task(
                f"Aggregate values using: {rule.transformation}",
                {"values": source_values, **(context or {})}
            )
            return result

        # For simple aggregations, use direct evaluation
        agg_context = {
            **source_values,
            "sum": sum,
            "avg": lambda x: sum(x) / len(x) if x else 0,
            "count": len,
            "min": min,
            "max": max
        }
        
        return eval(rule.transformation, {"__builtins__": {}}, agg_context)

    async def _execute_validate(self, rule: MigrationRule,
                              source_values: Dict[str, Any],
                              context: Optional[Dict[str, Any]]) -> ValidationResult:
        """Execute a validate rule."""
        if not rule.transformation and not rule.llm_prompt:
            raise ValueError("Validate rule must have either transformation or llm_prompt")

        try:
            if rule.llm_prompt:
                # Use validation agent for complex validations
                validation_agent = self.processor.get_or_create_agent("validation")
                result = await validation_agent.process_task(
                    f"Validate data using rules: {rule.llm_prompt}",
                    {"data": source_values, **(context or {})}
                )
                return ValidationResult(
                    is_valid=result.lower().startswith("valid"),
                    message=result
                )
            else:
                # Use direct validation for simple rules
                validation_context = {
                    **source_values,
                    "len": len,
                    "sum": sum,
                    "min": min,
                    "max": max,
                    "isinstance": isinstance,
                    "str": str,
                    "int": int,
                    "float": float
                }
                
                is_valid = eval(rule.transformation, {"__builtins__": {}}, validation_context)
                return ValidationResult(
                    is_valid=bool(is_valid),
                    message=None if is_valid else "Failed validation check"
                )

        except Exception as e:
            return ValidationResult(
                is_valid=False,
                message=f"Validation error: {str(e)}"
            )

    def _is_complex_computation(self, transformation: str) -> bool:
        """Determine if a computation requires the formula agent."""
        # Add logic to determine complexity
        return len(transformation.split()) > 5 or 'if' in transformation

    def _is_complex_aggregation(self, transformation: str) -> bool:
        """Determine if an aggregation requires the transformation agent."""
        # Add logic to determine complexity
        return len(transformation.split()) > 5 or any(
            keyword in transformation 
            for keyword in ['filter', 'map', 'lambda']
        )