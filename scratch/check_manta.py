import pandas as pd
import data_manager as dm

file_path = dm._get_apontamento_file()
sheet_name = dm._get_sheet("SHEET_AP_BD")

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
    print("Columns:", df.columns.tolist())
except Exception as e:
    print("Error:", e)
