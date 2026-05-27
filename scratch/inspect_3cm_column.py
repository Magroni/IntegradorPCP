# -*- coding: utf-8 -*-
import pandas as pd
import os

BASE_DIR = r"z:\PCP\PROJETOS MARLON\ProgramarProd"
file1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

try:
    df = pd.read_excel(file1, sheet_name="BD", skiprows=6, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    date_col = "DATA REG" if "DATA REG" in df.columns else df.columns[0]
    df["DT"] = pd.to_datetime(df[date_col], errors="coerce")
    df_may = df[(df["DT"].dt.month == 5) & (df["DT"].dt.year == 2026) & (df["PROCESSO"] != "PARADA")]
    
    print("=== INSPECTING '2CM' AND '3CM' COLUMNS IN MAY 2026 ===")
    print("Non-null count in '3CM':", df_may["3CM"].notna().sum())
    print("Non-null count in '2CM':", df_may["2CM"].notna().sum())
    
    print("\nSum of '3CM':", df_may["3CM"].sum())
    print("Sum of '2CM':", df_may["2CM"].sum())
    
    print("\nSample rows where '3CM' is not null and > 0:")
    sub_3cm = df_may[df_may["3CM"].notna() & (df_may["3CM"] > 0)]
    print(f"Count: {len(sub_3cm)}")
    if not sub_3cm.empty:
        print(sub_3cm[["DATA REG", "MATERIAL+BLOCO", "PROCESSO", "2CM", "3CM", "QTD CH (SEM RET & REPASSE)", "QTD M² (SEM RET & REPASSE)"]].head(15).to_string())
        
except Exception as e:
    print("Error:", e)
