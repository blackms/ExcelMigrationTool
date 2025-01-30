# Excel Migration Framework

A flexible framework for migrating Excel data using configurable rules and LLM integration. This framework allows you to define complex migration rules and transformations, with support for multiple LLM providers through LangChain.

## Features

- Rule-based Excel data migration
- Support for multiple LLM providers (OpenAI, Anthropic, etc.)
- Flexible rule types:
  - Direct copy
  - Value transformation
  - Computed fields
  - Aggregations
  - Validation rules
- LLM-powered transformations for complex cases
- Configurable via JSON rules files
- Extensible architecture

## Installation

```bash
# Using pip
pip install excel-migration-framework

# Using poetry
poetry add excel-migration-framework
```

## Quick Start

1. Create a rules configuration file (e.g., `rules.json`):

```json
{
  "rules": [
    {
      "type": "copy",
      "source_columns": ["A"],
      "target_column": "B",
      "description": "Copy values from column A to B"
    }
  ],
  "sheet_mapping": {
    "Sheet1": "Output1"
  }
}
```

2. Use the framework in your code:

```python
from excel_migration.core.models import MigrationContext
from excel_migration.core.processor import ExcelMigrationProcessor
from excel_migration.rules.engine import RuleEngine

# Initialize the rule engine
rule_engine = RuleEngine(
    llm_provider="openai",
    api_key="your-api-key"  # For LLM-powered transformations
)

# Load rules
rules = rule_engine.load_rules("rules.json")

# Create migration context
context = MigrationContext(
    source_file="source.xlsx",
    target_file="target.xlsx",
    rules=rules,
    sheet_mapping={"Sheet1": "Output1"}
)

# Execute migration
processor = ExcelMigrationProcessor(context)
success = processor.process()
```

## Rule Types

### Copy Rule
Directly copies values from source to target columns.

```json
{
  "type": "copy",
  "source_columns": ["A"],
  "target_column": "B"
}
```

### Transform Rule
Applies transformations to source values.

```json
{
  "type": "transform",
  "source_columns": ["B", "C"],
  "target_column": "D",
  "transformation": "float(B) * float(C)"
}
```

### Compute Rule
Computes new values based on source data.

```json
{
  "type": "compute",
  "source_columns": ["E", "F", "G"],
  "target_column": "H",
  "transformation": "sum([float(E), float(F), float(G)])"
}
```

### Aggregate Rule
Performs aggregations on source data.

```json
{
  "type": "aggregate",
  "source_columns": ["M", "N", "O"],
  "target_column": "P",
  "transformation": "avg([float(x) for x in [M, N, O] if x])"
}
```

### Validate Rule
Validates data according to specified rules.

```json
{
  "type": "validate",
  "source_columns": ["K"],
  "target_column": "L",
  "transformation": "float(K) > 0 and float(K) < 1000000"
}
```

## LLM Integration

The framework supports LLM-powered transformations for complex cases:

```json
{
  "type": "transform",
  "source_columns": ["I"],
  "target_column": "J",
  "llm_prompt": "Convert the technical description to user-friendly format"
}
```

Configure LLM provider in your code:

```python
rule_engine = RuleEngine(
    llm_provider="openai",  # or "anthropic"
    api_key="your-api-key",
    model_name="gpt-4",  # optional
    temperature=0.7  # optional
)
```

## Conditions

Rules can include conditions for selective application:

```json
{
  "type": "transform",
  "source_columns": ["A"],
  "target_column": "B",
  "conditions": {
    "A": {
      "operator": ">",
      "value": 0
    }
  }
}
```

Available operators:
- `==`: Equal to
- `!=`: Not equal to
- `>`: Greater than
- `<`: Less than
- `in`: Value in list
- `contains`: String contains

## Variables

Global variables can be defined in the rules file:

```json
{
  "variables": {
    "margin_rate": 0.3,
    "tax_rate": 0.2,
    "currency": "EUR"
  }
}
```

These variables are available in transformations and computations.

## Error Handling

The framework provides detailed logging and error handling:

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
