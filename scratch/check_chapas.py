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
    df_may = df[(df["DT"].dt.month == 5) & (df["DT"].dt.year == 2026) & (df["PROCESSO"] != "PARADA")]
    
    # Check the three QTD CH columns
    ch_cols = [c for c in df_may.columns if "QTD CH" in c or "CHAPAS" in c or "QTD. CH" in c]
    print("Available ch columns in May 2026:", ch_cols)
    
    print("\nCounts of non-null values in May 2026 production rows:")
    for col in ch_cols:
        print(f"  {col}: {df_may[col].notna().sum()} non-null, sum: {df_may[col].sum()}")
        
    print("\nSample rows where QTD CH (SEM RET & REPASSE) is null or 0:")
    null_ch = df_may[df_may["QTD CH (SEM RET & REPASSE)"].isna() | (df_may["QTD CH (SEM RET & REPASSE)"] == 0)]
    print(f"Total rows with null/0 chapas: {len(null_ch)} of {len(df_may)}")
    
    if not null_ch.empty:
        print(null_ch[["DATA REG", "MATERIAL+BLOCO", "PROCESSO"] + ch_cols + ["QTD M² (SEM RET & REPASSE)"]].head(10).to_string())
        
except Exception as e:
    print("Error:", e)
