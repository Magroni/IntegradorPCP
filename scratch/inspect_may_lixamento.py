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
    
    search_terms = ["CHECK", "LIXA", "LEVANTA"]
    print("=== SEARCHING IN MAY 2026 ROWS ===")
    
    for term in search_terms:
        mask_mat = df_may["MATERIAL+BLOCO"].astype(str).str.upper().str.contains(term, na=False)
        mask_proc = df_may["PROCESSO"].astype(str).str.upper().str.contains(term, na=False)
        
        matches = df_may[mask_mat | mask_proc]
        print(f"\nTerm '{term}': found {len(matches)} matches in May 2026:")
        if len(matches) > 0:
            print(matches[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR", "HORIM. INI", "HORIM. FIM"]].to_string())
            
except Exception as e:
    print("Error:", e)
