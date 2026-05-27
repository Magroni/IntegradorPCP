# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
prog_file = os.path.join(BASE_DIR, "COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm")

try:
    df = pd.read_excel(prog_file, sheet_name="DB", nrows=5, engine="openpyxl")
    print("Programming DB Columns:")
    for idx, col in enumerate(df.columns):
        print(f"  [{idx}] {col}")
    print("\nFirst row sample:")
    print(df.iloc[0].to_dict())
except Exception as e:
    print("Error:", e)
