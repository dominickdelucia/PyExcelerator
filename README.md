# PyExcelerator

This project provides tools to convert Excel spreadsheets into a Python
project. Static data is written to JSON files while formulas are translated to
Python scripts with help from an LLM. The converter groups contiguous cells into
blocks of data or formulas so that each block becomes a pandas DataFrame stored
in JSON. The generated project can also be used to recreate the original
spreadsheet.

## Installation
```bash
pip install -r requirements.txt
```

## Usage
Convert an Excel file to a Python project:
```bash
python excel_to_python.py my_sheet.xlsx output_project
```

Optionally verify the captured blocks with a multimodal LLM (requires an
OpenAI API key):
```bash
python -c "import openpyxl, verifier; wb=openpyxl.load_workbook('my_sheet.xlsx');\
verifier.sheet_to_png(wb.active, 'sheet.png'); print(verifier.verify_structure('sheet.png'))"
```

Rebuild the Excel file from the project:
```bash
python python_to_excel.py output_project restored.xlsx
```

An environment variable `OPENAI_API_KEY` may be set to enable LLM-based
translation of formulas. Without it, placeholder translations are written.

Generated data blocks are placed in `output_project/data` while formulas and
their Python translations are found in `output_project/formulas`.
