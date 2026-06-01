import pandas as pd
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import data_manager as dm

file_path = dm._get_apontamento_file()

def inspect_sheet(sheet_name):
    print(f"\n--- Sheet: {sheet_name} ---")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        print("Shape:", df.shape)
        print("Columns:", list(df.columns))
        print("Head:")
        print(df.head(5))
    except Exception as e:
        print("Error:", e)

inspect_sheet("DB")
inspect_sheet("PARADAS")
inspect_sheet("INSUMOS")
