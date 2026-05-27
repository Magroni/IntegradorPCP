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
    df_may = df[(df["DT"].dt.month == 5) & (df["DT"].dt.year == 2026)].copy()
    
    print("=== UNIQUE MATERIAL+BLOCO VALUES IN MAY 2026 ===")
    print(df_may["MATERIAL+BLOCO"].unique())
    
    # Check if there are other columns that could contain stops
    # Such as OBSERVAÇÃO, OBS. REPROCESSOS, MOTIVO RETRABALHO/REPASSE
    print("\nColumns in df_may:")
    print(df_may.columns.tolist())
    
    # Check non-null values in OBS. REPROCESSOS in May 2026
    obs_cols = [c for c in ["OBS. REPROCESSOS", "MOTIVO RETRABALHO/REPASSE", "TIPO DE PARADA"] if c in df_may.columns]
    for c in obs_cols:
        non_null = df_may[df_may[c].notna()]
        print(f"\nNon-null in {c} ({len(non_null)} rows):")
        if len(non_null) > 0:
            print(non_null[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", c]].head(10).to_string())
            
except Exception as e:
    print("Error:", e)
