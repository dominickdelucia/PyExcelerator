import json
import os
import sys
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

from translator import excel_formula_to_python


def _sheet_to_dataframe(ws: openpyxl.worksheet.worksheet.Worksheet) -> pd.DataFrame:
    """Return DataFrame with Excel-style column labels and 1-based row index."""
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    cols = [get_column_letter(i) for i in range(1, ws.max_column + 1)]
    df = pd.DataFrame(data, columns=cols, index=range(1, ws.max_row + 1))
    return df


def _find_blocks(mask: pd.DataFrame) -> list[list[tuple[int, int]]]:
    """Find contiguous blocks of True values in mask."""
    visited = set()
    blocks = []
    rows, cols = mask.shape
    for r in range(rows):
        for c in range(cols):
            if mask.iat[r, c] and (r, c) not in visited:
                stack = [(r, c)]
                block = []
                while stack:
                    x, y = stack.pop()
                    if (x, y) in visited or not mask.iat[x, y]:
                        continue
                    visited.add((x, y))
                    block.append((x, y))
                    for dx, dy in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
                        nx, ny = x + dx, y + dy
                        if 0 <= nx < rows and 0 <= ny < cols and (nx, ny) not in visited:
                            stack.append((nx, ny))
                if block:
                    blocks.append(block)
    return blocks


def excel_to_python(excel_path: str, output_dir: str = "output") -> None:
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    data_dir = Path(output_dir) / "data"
    formula_dir = Path(output_dir) / "formulas"
    data_dir.mkdir(parents=True, exist_ok=True)
    formula_dir.mkdir(parents=True, exist_ok=True)

    for ws in wb.worksheets:
        sheet_name = ws.title
        df = _sheet_to_dataframe(ws)

        is_formula = df.applymap(lambda v: isinstance(v, str) and str(v).startswith("="))
        is_value = df.notna() & (~is_formula)

        value_blocks = _find_blocks(is_value)
        formula_blocks = _find_blocks(is_formula)

        # Save value blocks
        for i, block in enumerate(value_blocks, start=1):
            rows = [r for r, _ in block]
            cols = [c for _, c in block]
            df_block = df.iloc[min(rows):max(rows)+1, min(cols):max(cols)+1]
            df_block.index = range(min(rows)+1, max(rows)+2)
            df_block.columns = [get_column_letter(c+1) for c in range(min(cols), max(cols)+1)]
            df_block.to_json(data_dir / f"{sheet_name}_data_block{i}.json", orient="split", indent=2)

        # Save formula blocks and Python translations
        for i, block in enumerate(formula_blocks, start=1):
            rows = [r for r, _ in block]
            cols = [c for _, c in block]
            df_block = df.iloc[min(rows):max(rows)+1, min(cols):max(cols)+1]
            df_block.index = range(min(rows)+1, max(rows)+2)
            df_block.columns = [get_column_letter(c+1) for c in range(min(cols), max(cols)+1)]

            formula_json = {}
            formula_py_lines = [f"# Auto-generated formulas for block {i} in sheet {sheet_name}\n", "FORMULAS = {}\n"]
            for r in range(df_block.shape[0]):
                for c in range(df_block.shape[1]):
                    cell_label = f"{df_block.columns[c]}{df_block.index[r]}"
                    formula = str(df_block.iat[r, c])
                    formula_json[cell_label] = formula
                    formula_py_lines.append(
                        f"FORMULAS['{cell_label}'] = {excel_formula_to_python(formula)!r}\n"
                    )

            with open(formula_dir / f"{sheet_name}_block{i}.json", "w", encoding="utf-8") as f:
                json.dump(formula_json, f, indent=2, ensure_ascii=False)
            with open(formula_dir / f"{sheet_name}_block{i}.py", "w", encoding="utf-8") as f:
                f.writelines(formula_py_lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_python.py <excel_file> [output_dir]")
        sys.exit(1)
    excel_file = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    excel_to_python(excel_file, out_dir)
