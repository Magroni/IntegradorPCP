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
    print(f"May 2026 rows found: {len(df_may)}")
    
    # Print the columns that are not entirely null in df_may
    non_null_cols = []
    for c in df_may.columns:
        if df_may[c].notna().any():
            non_null_cols.append(c)
            
    print(f"Non-null columns in May 2026: {non_null_cols}")
    print("\nFirst 3 rows of May 2026:")
    # Print a selection of interesting columns
    print(df_may.head(3)[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR", "QTD CH (SEM RET & REPASSE)", "QTD M² (SEM RET & REPASSE)", "COMP.", "ALT.", "TURNO"]].to_string())
    
    # Print details for resin columns
    resin_cols = [c for c in ["TPO RESINA", "QTD.KG", "TIPO.ENDUR", "QTD.KG.1", "24H"] if c in df_may.columns]
    print("\nResin columns for non-null rows:")
    print(df_may[df_may["TPO RESINA"].notna() | df_may["QTD.KG"].notna()].head(3)[["DATA REG", "MATERIAL+BLOCO", "PROCESSO"] + resin_cols].to_string())
    
except Exception as e:
    print("Error:", e)
