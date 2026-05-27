# -*- coding: utf-8 -*-
import openpyxl
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
prog_file = os.path.join(BASE_DIR, "COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm")

try:
    wb = openpyxl.load_workbook(prog_file, read_only=True)
    print("Sheets in Programming DB:", wb.sheetnames)
    wb.close()
except Exception as e:
    print("Error:", e)
