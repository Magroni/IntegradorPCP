# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df = pd.read_excel(file1, sheet_name="BD", skiprows=6, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    # Extract bloco
    def extract_bloco(x):
        b = str(x).split("-")[-1].strip() if "-" in str(x) else str(x).strip()
        if b.endswith(".0"): b = b[:-2]
        return b
    df["BLOCO_NORM"] = df["MATERIAL+BLOCO"].apply(extract_bloco)
    
    # Check block 3706
    print("=== Block 3706 in REV 1 ===")
    sub_3706 = df[df["BLOCO_NORM"] == "3706"]
    print(sub_3706[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR", "QTD CH (SEM RET & REPASSE)", "HORIM. INI", "HORIM. FIM"]].to_string())
    
    # Check block 4450
    print("\n=== Block 4450 in REV 1 ===")
    sub_4450 = df[df["BLOCO_NORM"] == "4450"]
    print(sub_4450[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "SETOR", "QTD CH (SEM RET & REPASSE)", "HORIM. INI", "HORIM. FIM"]].to_string())
    
except Exception as e:
    print("Error:", e)
