# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df = pd.read_excel(file1, sheet_name="BD", skiprows=6, usecols=[0, 30, 31], engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    print("Columns read:", df.columns.tolist())
    
    # Check unique values in ANO or date column
    print("Unique years in ANO:")
    print(df["ANO"].unique())
    
    # Check unique month/years in MES/ANO
    print("Unique values in MÊS/ANO (first 20):")
    print(df["MÊS/ANO"].dropna().unique()[:20])
    
    # Parse date column
    date_col = "DATA REG" if "DATA REG" in df.columns else df.columns[0]
    df["DT"] = pd.to_datetime(df[date_col], errors="coerce")
    print("Min date:", df["DT"].min())
    print("Max date:", df["DT"].max())
    
    # Check specifically for May 2026 or any May
    df_may_2026 = df[(df["DT"].dt.month == 5) & (df["DT"].dt.year == 2026)]
    print(f"May 2026 rows: {len(df_may_2026)}")
    
    # Count rows by month for the latest year
    latest_year = df["DT"].dt.year.max()
    print(f"Months in the latest year ({latest_year}):")
    print(df[df["DT"].dt.year == latest_year]["DT"].dt.month.value_counts())
    
except Exception as e:
    print("Error:", e)
