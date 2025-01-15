# COaaS (Cost of Ownership as a Service) Calculator

A Python-based tool for calculating Cost of Ownership as a Service, including startup costs, yearly costs, and monthly fees across multiple service modules.

## Features

- Generates Excel spreadsheets with detailed cost breakdowns
- Calculates total costs including:
  - Startup costs
  - Year 2 and Year 3 costs
  - Total days of work
  - Total costs and revenues
  - Desired profit margins
  - Monthly fees
- Supports multiple service modules (M1, M2, M3)
- Automatic formula generation for Excel cells

## Requirements

- Python 3.x
- openpyxl
- openai (optional, for GG Startup analysis)

## Installation

1. Clone this repository
2. Install dependencies using Poetry:
```bash
poetry install
```

## Usage

Run the script with the following command:

```bash
python src/migrator.py -i <input_file> -o <output_file> [-t <template_file>] [--openai-key <key>]
```

Arguments:
- `-i, --input`: Input Excel file path (required)
- `-o, --output`: Output Excel file path (required)
- `-t, --template`: Template Excel file path (optional, defaults to template/template.xlsx)
- `-v, --verbose`: Enable verbose logging
- `--openai-key`: OpenAI API key for GG Startup analysis (optional)

You can also set the OpenAI API key via environment variable:
```bash
export OPENAI_API_KEY=your-key-here
python src/migrator.py -i input.xlsx -o output.xlsx
```

The script will create an Excel file with the following columns:
- Product Element (Column B)
- Cost Type (Column C)
- GG Startup (Column D) - Automatically analyzed for Fixed costs using OpenAI
- Canone Prezzo Mese (Column N)

## GG Startup Analysis

When the OpenAI API key is provided, the tool uses GPT-4 to analyze product descriptions and determine appropriate startup days for Fixed Optional and Fixed Mandatory costs. The analysis considers:

- Explicit references to setup/startup/installation days
- Product/service complexity
- Context-based estimation when explicit references are not available

This automated analysis helps maintain consistency in startup day calculations across the sheet.

## License

[MIT License](LICENSE)
