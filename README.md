# Excel Migration Framework

A powerful framework for migrating Excel data using configurable rules, multimodal analysis, and LLM integration. This framework allows you to define complex migration rules, learn from examples, and leverage visual analysis of Excel sheets.

## Features

- Task-centric approach for Excel migrations
- Support for multiple LLM providers through LangChain
- Multimodal analysis capabilities:
  - Direct Excel file processing
  - Screenshot analysis and data extraction
  - Visual structure recognition
  - OCR for text extraction
- Rule generation from example files
- Flexible rule types:
  - Direct copy
  - Value transformation
  - Computed fields
  - Aggregations
  - Validation rules
- LLM-powered transformations
- Configurable via JSON rules
- Comprehensive logging with loguru
- SOLID principles and clean architecture

## Installation

```bash
# Using pip
pip install excel-migration-framework

# Using poetry
poetry add excel-migration-framework
```

## Quick Start

### Basic Usage

```bash
# Simple migration with rules
excel-migrate source.xlsx target.xlsx --rules rules.json

# Generate rules from example files
excel-migrate source.xlsx target.xlsx \
    --example-source example_source.xlsx \
    --example-target example_target.xlsx

# Include screenshots for visual analysis
excel-migrate source.xlsx target.xlsx \
    --screenshots sheet1.png sheet2.png
```

### Python API

```python
from excel_migration.tasks.base import MigrationTask
from excel_migration.core.processor import TaskBasedProcessor
from pathlib import Path

# Create a migration task
task = MigrationTask(
    source_file=Path("source.xlsx"),
    target_file=Path("target.xlsx"),
    task_type="migrate",
    description="Migrate customer data",
    context={},
    screenshots=[Path("sheet1.png")]
)

# Process the task
processor = TaskBasedProcessor(...)
success = await processor.process(task)
```

## Task Types

### Migration Task
Migrates data from source to target Excel files.

```bash
excel-migrate source.xlsx target.xlsx --task-type migrate
```

### Analysis Task
Analyzes Excel files and provides insights.

```bash
excel-migrate source.xlsx target.xlsx --task-type analyze
```

### Validation Task
Validates data against rules.

```bash
excel-migrate source.xlsx target.xlsx --task-type validate
```

## Multimodal Analysis

The framework can analyze Excel sheets through multiple approaches:

1. Direct File Analysis
   - Structure analysis
   - Formula parsing
   - Data type detection

2. Visual Analysis (from screenshots)
   - Table structure detection
   - Cell boundary recognition
   - Text extraction (OCR)
   - Layout analysis

3. LLM Integration
   - Natural language understanding
   - Complex pattern recognition
   - Context-aware transformations

## Rule Generation

Rules can be generated automatically by analyzing example files:

```bash
# Generate rules from examples
excel-migrate source.xlsx target.xlsx \
    --example-source example_source.xlsx \
    --example-target example_target.xlsx \
    --output-rules rules.json
```

The framework will:
1. Analyze source and target examples
2. Identify patterns and transformations
3. Generate appropriate rules
4. Save rules for future use

## Configuration

### LLM Providers

```bash
# Use OpenAI
excel-migrate source.xlsx target.xlsx --llm-provider openai --model gpt-4

# Use Anthropic
excel-migrate source.xlsx target.xlsx --llm-provider anthropic --model claude-2
```

### Logging

```bash
# Set log level
excel-migrate source.xlsx target.xlsx --log-level DEBUG

# Log to file
excel-migrate source.xlsx target.xlsx --log-file migration.log
```

## Advanced Features

### Custom Rule Types

Create custom rule types by implementing the Rule interface:

```python
from excel_migration.core.interfaces import Rule

class CustomRule(Rule):
    async def apply(self, data: Any, context: Dict[str, Any]) -> Any:
        # Implement custom logic
        pass
```

### Event Handling

Subscribe to migration events:

```python
from excel_migration.core.interfaces import EventEmitter

def on_cell_processed(data: Dict[str, Any]):
    print(f"Processed cell: {data}")

emitter = EventEmitter()
emitter.on("cell_processed", on_cell_processed)
```

### Caching

Enable caching for better performance:

```bash
excel-migrate source.xlsx target.xlsx --cache-dir ./cache
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

```bash
# Clone repository
git clone https://github.com/yourusername/excel-migration-framework.git

# Install dependencies
poetry install

# Run tests
poetry run pytest
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
