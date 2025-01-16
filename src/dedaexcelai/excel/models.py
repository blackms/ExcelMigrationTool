"""Domain models for Excel processing."""
from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional

class ElementType(Enum):
    """Type of element in the schema."""
    ELEMENT = "Element"
    SUB_ELEMENT = "SubElement"

class CostType(Enum):
    """Type of cost."""
    FIXED_OPTIONAL = "Fixed Optional"
    FIXED_MANDATORY = "Fixed Mandatory"
    FEE_OPTIONAL = "Fee Optional"
    FEE_MANDATORY = "Fee Mandatory"
    
    @property
    def is_fixed(self) -> bool:
        """Check if this is a Fixed cost type."""
        return self in (CostType.FIXED_OPTIONAL, CostType.FIXED_MANDATORY)
    
    @property
    def is_fee(self) -> bool:
        """Check if this is a Fee cost type."""
        return self in (CostType.FEE_OPTIONAL, CostType.FEE_MANDATORY)

@dataclass
class Element:
    """Represents an element in the schema."""
    name: str
    element_type: ElementType
    cost_type: CostType
    row: int
    length: Optional[int] = None
    startup_days: Optional[int] = None

@dataclass
class CostMapping:
    """Column mappings for costs."""
    source_cost: int  # Source cost column
    target_cost: int  # Target cost column
    source_price: int  # Source price column
    target_price: int  # Target price column
    target_margin: int  # Target margin column
    margin_value: float = 0.3930  # Default 39.30%

class FixedCostMapping(CostMapping):
    """Column mappings for Fixed costs."""
    def __init__(self):
        super().__init__(
            source_cost=8,   # H
            target_cost=12,  # L
            source_price=12, # L
            target_price=14, # N
            target_margin=13 # M
        )

class FeeCostMapping(CostMapping):
    """Column mappings for Fee costs."""
    def __init__(self):
        super().__init__(
            source_cost=9,   # I
            target_cost=15,  # O
            source_price=13, # M
            target_price=17, # Q
            target_margin=16 # P
        )
