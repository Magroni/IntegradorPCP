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
    df_2026 = df[df["DT"].dt.year == 2026].copy()
    
    print(f"Total rows in 2026 in REV 1: {len(df_2026)}")
    
    terms = ["LIXA", "LEVANTA", "CHECK"]
    for term in terms:
        mask_mat = df_2026["MATERIAL+BLOCO"].astype(str).str.upper().str.contains(term, na=False)
        mask_proc = df_2026["PROCESSO"].astype(str).str.upper().str.contains(term, na=False)
        matches = df_2026[mask_mat | mask_proc]
        print(f"\nTerm '{term}': found {len(matches)} matches in the year 2026:")
        if len(matches) > 0:
            print("Unique values in MATERIAL+BLOCO:")
            print(matches["MATERIAL+BLOCO"].unique())
            print("Unique values in PROCESSO:")
            print(matches["PROCESSO"].unique())
            print("Count by month in 2026:")
            print(df_2026.loc[matches.index, "DT"].dt.month.value_counts())
            
except Exception as e:
    print("Error:", e)
