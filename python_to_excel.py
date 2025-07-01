import json
import os
import sys
from pathlib import Path

import openpyxl


def python_to_excel(project_dir: str = "output", excel_path: str = "reconstructed.xlsx") -> None:
    data_dir = Path(project_dir) / "data"
    formula_dir = Path(project_dir) / "formulas"
    wb = openpyxl.Workbook()
    if wb.worksheets:
        wb.remove(wb.active)

    for data_file in data_dir.glob("*.json"):
        sheet_name = data_file.stem
        ws = wb.create_sheet(title=sheet_name)
        with open(data_file, "r", encoding="utf-8") as f:
            sheet_data = json.load(f)
        for cell, value in sheet_data.items():
            ws[cell] = value
        formula_file = formula_dir / f"{sheet_name}.json"
        if formula_file.exists():
            with open(formula_file, "r", encoding="utf-8") as f:
                formulas = json.load(f)
            for cell, formula in formulas.items():
                ws[cell] = formula
    wb.save(excel_path)


if __name__ == "__main__":
    project = sys.argv[1] if len(sys.argv) > 1 else "output"
    out_excel = sys.argv[2] if len(sys.argv) > 2 else "reconstructed.xlsx"
    python_to_excel(project, out_excel)
