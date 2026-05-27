# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df_hist = pd.read_excel(file1, sheet_name="HISTÓRICO PARADAS", skiprows=6, engine="openpyxl")
    print("Columns of HISTÓRICO PARADAS (REV 1):", df_hist.columns.tolist())
    print("\nFirst 20 rows of HISTÓRICO PARADAS:")
    print(df_hist.head(20).to_string())
except Exception as e:
    print("Error:", e)
