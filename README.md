# 📊 Excel Migration Framework

A powerful framework for migrating Excel data using configurable rules, multimodal analysis, and LLM integration. This framework allows you to define complex migration rules, learn from examples, and leverage visual analysis of Excel sheets.

## ✨ Features

- 🎯 Task-centric approach for Excel migrations
- 🤖 Support for multiple LLM providers through LangChain
- 👁️ Multimodal analysis capabilities:
  - 📑 Direct Excel file processing
  - 📸 Screenshot analysis and data extraction
  - 🔍 Visual structure recognition
  - 📝 OCR for text extraction
- 🧠 Rule generation from example files
- 🛠️ Flexible rule types:
  - 📋 Direct copy
  - 🔄 Value transformation
  - 🧮 Computed fields
  - 📊 Aggregations
  - ✅ Validation rules
- 🔌 Plugin-based rule execution:
  - 🧩 Extensible formula executors
  - 🔄 Custom transformations
  - 🎨 Modular design
  - 🛡️ SOLID principles
- 🤖 LLM-powered transformations
- ⚙️ Configurable via JSON rules
- 📝 Comprehensive logging with loguru
- 🏗️ SOLID principles and clean architecture

## 🚀 Installation

```bash
# Using pip
pip install excel-migration-framework

# Using poetry
poetry add excel-migration-framework
```

## 🏃‍♂️ Quick Start

### 📌 Basic Usage

```bash
# Simple migration with rules
excel-migrate source.xlsx target.xlsx --rules rules.json

# Process specific sheets
excel-migrate source.xlsx target.xlsx \
    --source-sheets "Sheet1" "Sheet2" \
    --target-sheets "Output1" "Output2"

# Generate rules from example files with sheet selection
excel-migrate source.xlsx target.xlsx \
    --example-source example_source.xlsx \
    --example-target example_target.xlsx \
    --example-source-sheets "Template1" \
    --example-target-sheets "Result1"

# Include screenshots with sheet mapping
excel-migrate source.xlsx target.xlsx \
    --screenshots sheet1.png sheet2.png \
    --screenshot-sheet-mapping "sheet1.png:Sheet1" "sheet2.png:Sheet2"
```

### 💻 Python API

```python
from excel_migration.tasks.base import MigrationTask
from excel_migration.core.processor import TaskBasedProcessor
from pathlib import Path

# Create a migration task with sheet selection
task = MigrationTask(
    source_file=Path("source.xlsx"),
    target_file=Path("target.xlsx"),
    task_type="migrate",
    description="Migrate customer data",
    context={
        "sheet_mapping": {
            "CustomerData": "Processed_Customers",
            "Transactions": "Processed_Transactions"
        }
    },
    screenshots=[Path("sheet1.png")]
)

# Process the task
processor = TaskBasedProcessor(...)
success = await processor.process(task)
```

## 🔌 Plugin System

The framework uses a flexible plugin system for formula execution and value transformations, following SOLID principles:

### 🧩 Formula Executors

```python
from excel_migration.plugins.interfaces import FormulaExecutor
from typing import Any, Dict

class CustomFormulaExecutor(FormulaExecutor):
    """Custom formula executor plugin."""
    
    formula_type = "CUSTOM"
    
    def can_execute(self, formula: str) -> bool:
        """Check if this executor can handle the formula."""
        return formula.startswith("CUSTOM(")
    
    def execute(self, formula: str, values: Dict[str, Any]) -> Any:
        """Execute the custom formula."""
        # Implement custom formula logic
        pass

# Register the plugin
registry = PluginRegistry()
registry.register_formula_executor(CustomFormulaExecutor())
```

### 🔄 Transformation Handlers

```python
from excel_migration.plugins.interfaces import TransformationHandler
from typing import Any, Dict

class CustomTransformer(TransformationHandler):
    """Custom transformation plugin."""
    
    transformation_type = "custom_format"
    
    def can_transform(self, transformation: Dict[str, Any]) -> bool:
        """Check if this handler can process the transformation."""
        return transformation.get("type") == self.transformation_type
    
    def transform(self, value: Any, params: Dict[str, Any]) -> Any:
        """Transform the value according to parameters."""
        # Implement custom transformation logic
        pass

# Register the plugin
registry.register_transformation_handler(CustomTransformer())
```

### 📦 Built-in Plugins

The framework includes several built-in plugins:

#### Formula Executors:
- 📅 `DateDiffExecutor`: Calculate date differences
- 🔢 `CountExecutor`: Count values or records
- 🎯 `CountIfExecutor`: Conditional counting
- ➕ `SumExecutor`: Sum numeric values
- 📊 `AverageExecutor`: Calculate averages

#### Transformation Handlers:
- 📅 `DateTimeTransformer`: Format dates and times
- 🔢 `NumericTransformer`: Format numbers
- ✅ `BooleanTransformer`: Convert to boolean values
- 🔤 `ConcatenateTransformer`: Join multiple values

## 🎯 Task Types

### 🔄 Migration Task
Migrates data from source to target Excel files.

```bash
excel-migrate source.xlsx target.xlsx \
    --task-type migrate \
    --source-sheets "Data" \
    --target-sheets "Processed"
```

### 🔍 Analysis Task
Analyzes Excel files and provides insights.

```bash
excel-migrate source.xlsx target.xlsx \
    --task-type analyze \
    --source-sheets "Financial" "Metrics"
```

### ✅ Validation Task
Validates data against rules.

```bash
excel-migrate source.xlsx target.xlsx \
    --task-type validate \
    --source-sheets "Input" \
    --rules validation_rules.json
```

## 🔮 Multimodal Analysis

The framework can analyze Excel sheets through multiple approaches:

1. 📊 Direct File Analysis
   - 🔍 Structure analysis
   - 📐 Formula parsing
   - 🏷️ Data type detection

2. 👁️ Visual Analysis (from screenshots)
   - 📏 Table structure detection
   - 🔲 Cell boundary recognition
   - 📝 Text extraction (OCR)
   - 🎨 Layout analysis

3. 🧠 LLM Integration
   - 💭 Natural language understanding
   - 🔄 Complex pattern recognition
   - 📚 Context-aware transformations

## ⚡ Rule Generation

Rules can be generated automatically by analyzing example files:

```bash
# Generate rules from specific sheets in examples
excel-migrate source.xlsx target.xlsx \
    --example-source example_source.xlsx \
    --example-target example_target.xlsx \
    --example-source-sheets "Template" \
    --example-target-sheets "Final" \
    --output-rules rules.json
```

The framework will:
1. 🔍 Analyze source and target examples
2. 🧮 Identify patterns and transformations
3. ✨ Generate appropriate rules
4. 💾 Save rules for future use

## ⚙️ Configuration

### 🤖 LLM Providers

```bash
# Use OpenAI
excel-migrate source.xlsx target.xlsx \
    --llm-provider openai \
    --model gpt-4

# Use Anthropic
excel-migrate source.xlsx target.xlsx \
    --llm-provider anthropic \
    --model claude-2
```

### 📝 Logging

```bash
# Set log level
excel-migrate source.xlsx target.xlsx --log-level DEBUG

# Log to file
excel-migrate source.xlsx target.xlsx --log-file migration.log
```

## 🔧 Advanced Features

### 🛠️ Custom Rule Types

Create custom rule types by implementing the Rule interface:

```python
from excel_migration.core.interfaces import Rule

class CustomRule(Rule):
    async def apply(self, data: Any, context: Dict[str, Any]) -> Any:
        # Implement custom logic
        pass
```

### 📡 Event Handling

Subscribe to migration events:

```python
from excel_migration.core.interfaces import EventEmitter

def on_cell_processed(data: Dict[str, Any]):
    print(f"Processed cell: {data}")

emitter = EventEmitter()
emitter.on("cell_processed", on_cell_processed)
```

### 💾 Caching

Enable caching for better performance:

```bash
excel-migrate source.xlsx target.xlsx --cache-dir ./cache
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### 🛠️ Development Setup

```bash
# Clone repository
git clone https://github.com/yourusername/excel-migration-framework.git

# Install dependencies
poetry install

# Run tests
poetry run pytest
```

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.
