# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df = pd.read_excel(file1, sheet_name="BD", skiprows=6, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    date_col = "DATA REG" if "DATA REG" in df.columns else df.columns[0]
    df["DT"] = pd.to_datetime(df[date_col], errors="coerce")
    df_may = df[(df["DT"].dt.month == 5) & (df["DT"].dt.year == 2026)]
    
    df_parada = df_may[df_may["PROCESSO"] == "PARADA"]
    print(f"May 2026 PARADA rows: {len(df_parada)}")
    if not df_parada.empty:
        cols_with_data = [c for c in df_parada.columns if df_parada[c].notna().any()]
        print("Columns with data in PARADA rows:", cols_with_data)
        print("\nPARADA samples:")
        print(df_parada[cols_with_data].head(5).to_string())
except Exception as e:
    print("Error:", e)
