import os
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

try:
    import openai
except ImportError:
    openai = None


def sheet_to_png(ws: openpyxl.worksheet.worksheet.Worksheet, path: str) -> None:
    df = pd.DataFrame([[c.value for c in row] for row in ws.iter_rows()])
    fig, ax = plt.subplots()
    ax.axis("off")
    table = ax.table(cellText=df.values, loc="center", cellLoc="center")
    fig.tight_layout()
    plt.savefig(path)
    plt.close(fig)


def verify_structure(image_path: str, description: str = "") -> Optional[str]:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not (openai and api_key):
        return None
    openai.api_key = api_key
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": description or "Describe this spreadsheet"},
                {"type": "image_url", "image_url": image_path},
            ],
        }
    ]
    try:
        resp = openai.ChatCompletion.create(
            model="gpt-4-vision-preview",
            messages=messages,
        )
        return resp.choices[0].message.get("content")
    except Exception:
        return None

