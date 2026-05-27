# -*- coding: utf-8 -*-
import pandas as pd
import os
import openpyxl

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")
prog_file = os.path.join(BASE_DIR, "COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm")
blocks_file = os.path.join(BASE_DIR, "PLANILHA BLOCOS.xlsb")

try:
    print("Loading REV 1...")
    df_rev1 = pd.read_excel(file1, sheet_name="BD", skiprows=6, engine="openpyxl")
    df_rev1.columns = [str(c).strip().upper() for c in df_rev1.columns]
    
    date_col = "DATA REG" if "DATA REG" in df_rev1.columns else df_rev1.columns[0]
    df_rev1["DT"] = pd.to_datetime(df_rev1[date_col], errors="coerce")
    
    # Filter May 2026 production rows with null chapas
    df_may = df_rev1[(df_rev1["DT"].dt.month == 5) & (df_rev1["DT"].dt.year == 2026) & (df_rev1["PROCESSO"] != "PARADA")].copy()
    
    # Normalizar bloco em REV 1
    def normalize_bloco(b):
        if pd.isna(b): return ""
        b_str = str(b).strip().upper()
        if b_str.endswith(".0"): b_str = b_str[:-2]
        return b_str
        
    df_rev1["BLOCO_NORM"] = df_rev1["MATERIAL+BLOCO"].apply(lambda x: str(x).split("-")[-1].strip() if "-" in str(x) else "").apply(normalize_bloco)
    df_may["BLOCO_NORM"] = df_may["MATERIAL+BLOCO"].apply(lambda x: str(x).split("-")[-1].strip() if "-" in str(x) else "").apply(normalize_bloco)
    
    print("\n--- TEST: Looking up block 375 in REV 1 ---")
    sub_375 = df_rev1[df_rev1["BLOCO_NORM"] == "375"]
    print(f"Total rows for block 375 in REV 1: {len(sub_375)}")
    print(sub_375[["DATA REG", "PROCESSO", "QTD CH (SEM RET & REPASSE)", "QTD CH (COM RET & REPASSE)"]])
    
    print("\n--- TEST: Looking up block 6629 in REV 1 ---")
    sub_6629 = df_rev1[df_rev1["BLOCO_NORM"] == "6629"]
    print(f"Total rows for block 6629 in REV 1: {len(sub_6629)}")
    print(sub_6629[["DATA REG", "PROCESSO", "QTD CH (SEM RET & REPASSE)", "QTD CH (COM RET & REPASSE)"]])
    
except Exception as e:
    print("Error:", e)
