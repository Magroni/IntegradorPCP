# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df = pd.read_excel(file1, sheet_name="BD", skiprows=6, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    # Let's search for "LIXAMENTO" or "LEVANTAMENTO" or "LIXAR" in the entire sheet BD
    print("=== SEARCHING ENTIRE BD SHEET IN REV 1 FOR PARADAS ===")
    
    terms = ["LIXAMENTO", "LEVANTAMENTO", "LIXAR", "LEVANTA", "CHECK"]
    for term in terms:
        mask_mat = df["MATERIAL+BLOCO"].astype(str).str.upper().str.contains(term, na=False)
        mask_proc = df["PROCESSO"].astype(str).str.upper().str.contains(term, na=False)
        matches = df[mask_mat | mask_proc]
        print(f"\nKeyword '{term}': found {len(matches)} matches in the ENTIRE REV 1 sheet BD:")
        if len(matches) > 0:
            print("Unique values in MATERIAL+BLOCO:")
            print(matches["MATERIAL+BLOCO"].unique()[:10])
            print("Unique values in PROCESSO:")
            print(matches["PROCESSO"].unique()[:10])
            print("Count by month:")
            if "DATA REG" in df.columns:
                df["DT_TEMP"] = pd.to_datetime(df["DATA REG"], errors="coerce")
                print(df.loc[matches.index, "DT_TEMP"].dt.to_period("M").value_counts().head(10))
            
            # Print a few samples where PROCESSO is PARADA
            sub_parada = matches[matches["PROCESSO"] == "PARADA"]
            if len(sub_parada) > 0:
                print(f"Samples where PROCESSO is PARADA ({len(sub_parada)} rows):")
                print(sub_parada[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR"]].head(5).to_string())
            else:
                print("No samples found where PROCESSO is PARADA.")
                
except Exception as e:
    print("Error:", e)
