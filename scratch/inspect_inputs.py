# -*- coding: utf-8 -*-
import openpyxl
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")

print("=== INSPECTING REV 2 ADDITIONAL SHEETS ===")

for sheet in ["PARADAS", "INSUMOS"]:
    print(f"\nSheet: {sheet}")
    try:
        df = pd.read_excel(file2, sheet_name=sheet, nrows=5, engine="openpyxl")
        print("Columns:")
        for idx, col in enumerate(df.columns):
            print(f"  [{idx}] {col}")
        print("First row sample:")
        if not df.empty:
            print(df.iloc[0].to_dict())
        else:
            print("No rows found!")
    except Exception as e:
        print(f"Error reading {sheet}: {e}")
