# PyExcelerator

This project provides simple tools to convert Excel spreadsheets into a Python
project. Static data is written to JSON files while formulas are translated to
Python scripts with help from an LLM. The generated project can also be used to
recreate the original spreadsheet.

## Installation
```bash
pip install -r requirements.txt
```

## Usage
Convert an Excel file to a Python project:
```bash
python excel_to_python.py my_sheet.xlsx output_project
```

Rebuild the Excel file from the project:
```bash
python python_to_excel.py output_project restored.xlsx
```

An environment variable `OPENAI_API_KEY` may be set to enable LLM-based
translation of formulas. Without it, placeholder translations are written.
