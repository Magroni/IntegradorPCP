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
    
    print("=== Analyzing Chapas in May 2026 ===")
    print(f"Total production rows in May 2026: {len(df_may)}")
    
    # Rows with any chapa value recorded
    ch_cols = ['QTD CH (SEM RET & REPASSE)', 'QTD CH (RETRABALHO + REPASSE)', 'QTD CH (COM RET & REPASSE)']
    df_with_ch = df_may[df_may[ch_cols].notna().any(axis=1)]
    print(f"Rows with any chapa value: {len(df_with_ch)}")
    print(df_with_ch[["DATA REG", "MATERIAL+BLOCO", "PROCESSO"] + ch_cols].head(15).to_string())
    
    # Let's inspect other months to see if this is normal for REV 1
    print("\n=== Checking April 2026 to compare ===")
    df_apr = df[(df["DT"].dt.month == 4) & (df["DT"].dt.year == 2026) & (df["PROCESSO"] != "PARADA")]
    print(f"Total production rows in April 2026: {len(df_apr)}")
    df_apr_with_ch = df_apr[df_apr[ch_cols].notna().any(axis=1)]
    print(f"Rows with any chapa value in April: {len(df_apr_with_ch)}")
    
except Exception as e:
    print("Error:", e)
