"""Interfaces for Excel migration plugins."""
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Protocol
from datetime import datetime

class FormulaExecutor(Protocol):
    """Protocol for formula execution plugins."""
    
    @property
    def formula_type(self) -> str:
        """Get the type of formula this executor handles."""
        ...
    
    def can_execute(self, formula: str) -> bool:
        """Check if this executor can handle the given formula."""
        ...
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Execute the formula with the given values."""
        ...

class TransformationHandler(Protocol):
    """Protocol for transformation plugins."""
    
    @property
    def transformation_type(self) -> str:
        """Get the type of transformation this handler processes."""
        ...
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        ...
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Transform the value according to the parameters."""
        ...

class PluginRegistry:
    """Registry for formula and transformation plugins."""
    
    def __init__(self):
        """Initialize the plugin registry."""
        self._formula_executors: Dict[str, FormulaExecutor] = {}
        self._transformation_handlers: Dict[str, TransformationHandler] = {}
    
    def register_formula_executor(self, executor: FormulaExecutor) -> None:
        """Register a formula executor."""
        self._formula_executors[executor.formula_type] = executor
    
    def register_transformation_handler(self, handler: TransformationHandler) -> None:
        """Register a transformation handler."""
        self._transformation_handlers[handler.transformation_type] = handler
    
    def get_formula_executor(self, formula: str) -> Optional[FormulaExecutor]:
        """Get the appropriate formula executor for a formula."""
        for executor in self._formula_executors.values():
            if executor.can_execute(formula):
                return executor
        return None
    
    def get_transformation_handler(self, trans_type: str) -> Optional[TransformationHandler]:
        """Get a transformation handler by type."""
        return self._transformation_handlers.get(trans_type)