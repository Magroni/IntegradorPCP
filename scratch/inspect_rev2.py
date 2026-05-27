import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")

if os.path.exists(FILE_REV2):
    try:
        xl = pd.ExcelFile(FILE_REV2)
        print("Sheets in REV 2:", xl.sheet_names)
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            print(f"Sheet {sheet} columns: {list(df.columns)}")
    except Exception as e:
        print("Error reading REV 2:", e)
else:
    print("REV 2 file does not exist at:", FILE_REV2)
