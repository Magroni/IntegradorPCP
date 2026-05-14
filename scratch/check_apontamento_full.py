import pandas as pd
import openpyxl

file_path = 'Apontamento Produção (REV 2).xlsx'

print("--- EXAMINING APONTAMENTO PRODUÇÃO (REV 2).xlsx ---")

try:
    xl = pd.ExcelFile(file_path)
    print("Sheets:", xl.sheet_names)
    
    if 'DB' in xl.sheet_names:
        print("\n--- DB SHEET ---")
        # Read the first 20 rows to find the header
        df_db_top = pd.read_excel(file_path, sheet_name='DB', header=None, nrows=20)
        
        header_row = -1
        for i, row in df_db_top.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values]
            if "PROCESSO" in row_vals or "DATA REG" in row_vals or "ID" in row_vals:
                header_row = i
                print(f"Header found at row {i+1} (0-indexed: {i})")
                print("Header values:", row_vals)
                break
        
        if header_row != -1:
            df_db = pd.read_excel(file_path, sheet_name='DB', skiprows=header_row)
            print("Columns in DB:", df_db.columns.tolist())
            print("First row data:")
            print(df_db.head(1).to_dict(orient='records'))
        else:
            print("Could not find header row in DB sheet.")
            
    if 'PARADAS' in xl.sheet_names:
        print("\n--- PARADAS SHEET ---")
        df_paradas = pd.read_excel(file_path, sheet_name='PARADAS')
        print("Columns in PARADAS:", df_paradas.columns.tolist())
        print("First row data:")
        print(df_paradas.head(1).to_dict(orient='records'))
    else:
        print("\n'PARADAS' sheet NOT FOUND.")

except Exception as e:
    print("Error:", e)
