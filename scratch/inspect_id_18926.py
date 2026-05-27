# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")

try:
    df_db = pd.read_excel(file2, sheet_name="DB", engine="openpyxl")
    df_db.columns = [str(c).strip().upper() for c in df_db.columns]
    
    sub_db = df_db[df_db["ID"].isin([18926, 18944])]
    print("=== REV 2 DB records for ID 18926 and 18944 ===")
    print(sub_db.to_string())
    
    # Check if there are other paradas in REV 2 PARADAS sheet
    print("\n=== REV 2 PARADAS sheet (first 15 rows) ===")
    df_p = pd.read_excel(file2, sheet_name="PARADAS", engine="openpyxl")
    print(df_p.head(15).to_string())
except Exception as e:
    print("Error:", e)
