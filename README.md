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

## Installation

1. Clone this repository
2. Install dependencies using Poetry:
```bash
poetry install
```

## Usage

Run the script to generate an Excel file with cost calculations:

```bash
python src/COaaS.py
```

This will create an Excel file named `COaaS_12_mesi.xlsx` with the following columns:
- MODULO (Module)
- Giornate Startup (Startup Days)
- Giornate Anno 2 (Year 2 Days)
- Giornate Anno 3 (Year 3 Days)
- Giornate Totali (Total Days)
- Costo Totale (Total Cost)
- Margine desiderato (Desired Margin)
- Ricavo Totale (Total Revenue)
- Startup (Startup Cost)
- Resto (Remaining Revenue)
- Canone (12 mesi) (Monthly Fee)

## Configuration

The script includes default values for:
- Daily cost rate (€820)
- Desired margin (30%)
- Module configurations (M1, M2, M3) with predefined days and startup costs

These values can be modified directly in the script as needed.

## License

[MIT License](LICENSE)