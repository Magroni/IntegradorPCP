# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
blocks_file = os.path.join(BASE_DIR, "PLANILHA BLOCOS.xlsb")
chapas_file = os.path.join(BASE_DIR, "Estoque Chapas 2026.xlsx")
prog_file = os.path.join(BASE_DIR, "COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm")

print("=== CHECKING BLOCKS FILE ===")
try:
    df_blocks = pd.read_excel(blocks_file, sheet_name="PLAN. BLOCOS", skiprows=8, nrows=3, engine="pyxlsb")
    print("Blocks columns:", df_blocks.columns.tolist())
except Exception as e:
    print("Error reading blocks:", e)

print("\n=== CHECKING CHAPAS FILE ===")
try:
    df_chapas = pd.read_excel(chapas_file, sheet_name="ENTRADAS", nrows=3, engine="openpyxl")
    print("Chapas columns:", df_chapas.columns.tolist())
    if not df_chapas.empty:
        print("Chapas first row sample:", df_chapas.iloc[0].to_dict())
except Exception as e:
    print("Error reading chapas:", e)

print("\n=== CHECKING PROG FILE ===")
try:
    df_prog = pd.read_excel(prog_file, sheet_name="PROGRAMAÇÃO", nrows=3, engine="openpyxl")
    print("Prog columns:", df_prog.columns.tolist())
    if not df_prog.empty:
        print("Prog first row sample:", df_prog.iloc[0].to_dict())
except Exception as e:
    print("Error reading prog:", e)
