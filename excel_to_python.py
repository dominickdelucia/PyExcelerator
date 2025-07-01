import json
import os
import sys
from pathlib import Path

import openpyxl

from translator import excel_formula_to_python


def excel_to_python(excel_path: str, output_dir: str = "output") -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    data_dir = Path(output_dir) / "data"
    formula_dir = Path(output_dir) / "formulas"
    data_dir.mkdir(parents=True, exist_ok=True)
    formula_dir.mkdir(parents=True, exist_ok=True)

    for ws in wb.worksheets:
        sheet_name = ws.title
        sheet_data = {}
        sheet_formulas = {}
        py_translated = {}
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    sheet_formulas[cell.coordinate] = val
                    py_translated[cell.coordinate] = excel_formula_to_python(val)
                else:
                    sheet_data[cell.coordinate] = val
        with open(data_dir / f"{sheet_name}.json", "w", encoding="utf-8") as f:
            json.dump(sheet_data, f, indent=2, ensure_ascii=False)
        with open(formula_dir / f"{sheet_name}.json", "w", encoding="utf-8") as f:
            json.dump(sheet_formulas, f, indent=2, ensure_ascii=False)
        lines = [f"# Auto-generated formula translations for sheet {sheet_name}\n", "FORMULAS = {}\n"]
        for cell, code in py_translated.items():
            lines.append(f"FORMULAS['{cell}'] = {code!r}\n")
        with open(formula_dir / f"{sheet_name}.py", "w", encoding="utf-8") as f:
            f.writelines(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_python.py <excel_file> [output_dir]")
        sys.exit(1)
    excel_file = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    excel_to_python(excel_file, out_dir)
