# -*- coding: utf-8 -*-
import pandas as pd
import os
import openpyxl

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    wb = openpyxl.load_workbook(file1, read_only=True)
    print("Sheets in REV 1:", wb.sheetnames)
    wb.close()
    
    # Let's inspect "HISTÓRICO PARADAS"
    print("\n=== Inspecting 'HISTÓRICO PARADAS' sheet ===")
    df_hist = pd.read_excel(file1, sheet_name="HISTÓRICO PARADAS", nrows=10, engine="openpyxl")
    print(df_hist.head(10).to_string())
    
    # Check if there are columns related to date and description
    # Let's read more rows to search for May 2026
    print("\nReading more of 'HISTÓRICO PARADAS'...")
    df_hist_full = pd.read_excel(file1, sheet_name="HISTÓRICO PARADAS", engine="openpyxl")
    print("Dimensions of HISTÓRICO PARADAS:", df_hist_full.shape)
    
    # Check "PARADAS - PERDA PRODUÇÃO"
    print("\n=== Inspecting 'PARADAS - PERDA PRODUÇÃO' sheet ===")
    df_perda = pd.read_excel(file1, sheet_name="PARADAS - PERDA PRODUÇÃO", nrows=10, engine="openpyxl")
    print(df_perda.head(10).to_string())
    
except Exception as e:
    print("Error:", e)
