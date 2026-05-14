import pandas as pd

file_path = 'Apontamento Produção (REV 2).xlsx'
try:
    xl = pd.ExcelFile(file_path)
    print("Sheets:", xl.sheet_names)
    for sheet in xl.sheet_names:
        if "ENDUR" in sheet.upper():
            print(f"\n--- {sheet} ---")
            df = pd.read_excel(file_path, sheet_name=sheet)
            print(df.head(20).to_string())
except Exception as e:
    print("Error:", e)
