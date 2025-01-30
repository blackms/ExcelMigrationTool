"""Core domain models for Excel migration."""
from dataclasses import dataclass
from enum import Enum
from typing import Optional, Dict, Any

class CellType(Enum):
    """Generic cell type classification."""
    TEXT = "text"
    NUMBER = "number"
    FORMULA = "formula"
    DATE = "date"
    BOOLEAN = "boolean"

class RuleType(Enum):
    """Types of migration rules."""
    COPY = "copy"  # Direct copy
    TRANSFORM = "transform"  # Apply transformation
    COMPUTE = "compute"  # Compute new value
    AGGREGATE = "aggregate"  # Aggregate multiple values
    VALIDATE = "validate"  # Validation rule

@dataclass
class Cell:
    """Represents a cell in an Excel worksheet."""
    value: Any
    cell_type: CellType
    row: int
    column: int
    formula: Optional[str] = None
    style: Optional[Dict[str, Any]] = None

@dataclass
class MigrationRule:
    """Base class for migration rules."""
    rule_type: RuleType
    source_columns: list[str]  # Source column names/references
    target_column: str  # Target column name/reference
    conditions: Optional[Dict[str, Any]] = None  # Conditions for rule application
    transformation: Optional[str] = None  # Transformation logic/formula
    llm_prompt: Optional[str] = None  # LLM prompt for complex transformations

@dataclass
class ValidationResult:
    """Result of a validation rule."""
    is_valid: bool
    message: Optional[str] = None
    severity: str = "error"  # error, warning, info

@dataclass
class MigrationContext:
    """Context for migration execution."""
    source_file: str
    target_file: str
    rules: list[MigrationRule]
    sheet_mapping: Dict[str, str]  # Source to target sheet mapping
    variables: Dict[str, Any] = None  # Global variables for rule execution