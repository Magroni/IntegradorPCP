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
    
    # Check rows that don't have block numbers or have parada-like text in MATERIAL+BLOCO
    parada_keywords = ['ALMO', 'ABRASI', 'PARADA', 'TROCA', 'AJUSTE', 'SETUP', 'ABASTECER', 'AGUARDANDO', 'INTERVALO', 'MANUTEN']
    
    def is_parada(row):
        proc = str(row.get("PROCESSO", "")).strip().upper()
        mat_blo = str(row.get("MATERIAL+BLOCO", "")).strip().upper()
        
        if proc == "PARADA":
            return True
        for kw in parada_keywords:
            if kw in mat_blo or kw in proc:
                return True
        return False
        
    df_may["IS_PARADA"] = df_may.apply(is_parada, axis=1)
    
    df_paradas = df_may[df_may["IS_PARADA"]]
    df_prod = df_may[~df_may["IS_PARADA"]]
    
    print(f"Total rows in May 2026: {len(df_may)}")
    print(f"Production rows: {len(df_prod)}")
    print(f"Parada rows: {len(df_paradas)}")
    
    print("\nParada rows uniqueMATERIAL+BLOCO:")
    print(df_paradas["MATERIAL+BLOCO"].unique())
    
    print("\nAll parada rows found:")
    print(df_paradas[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR", "HORIM. INI", "HORIM. FIM"]].to_string())
    
except Exception as e:
    print("Error:", e)
