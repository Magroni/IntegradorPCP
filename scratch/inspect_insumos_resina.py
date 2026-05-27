# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")

try:
    df = pd.read_excel(file2, sheet_name="INSUMOS", engine="openpyxl")
    print("Insumos unique types:", df["TIPO_INSUMO"].unique())
    print("\nInsumos samples by type:")
    for t in df["TIPO_INSUMO"].unique():
        sub = df[df["TIPO_INSUMO"] == t]
        print(f"\nType: {t}")
        print(sub.head(3).to_dict(orient='records'))
except Exception as e:
    print("Error:", e)
