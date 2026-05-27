import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

if os.path.exists(FILE_REV1):
    try:
        df_rev1 = pd.read_excel(FILE_REV1, sheet_name="BD", skiprows=6, engine="openpyxl")
        df_rev1.columns = [str(c).strip().upper() for c in df_rev1.columns]
        
        col_data_filt = "DATA REG" if "DATA REG" in df_rev1.columns else df_rev1.columns[0]
        df_rev1["DATETIME_CONVERTED"] = pd.to_datetime(df_rev1[col_data_filt], errors="coerce")
        df_may = df_rev1[
            (df_rev1["DATETIME_CONVERTED"].dt.month == 5) & 
            (df_rev1["DATETIME_CONVERTED"].dt.year == 2026)
        ].copy()
        
        print("Sample of TEMPO EFETIVO PRODUÇÃO/PARADA column in May:")
        s = df_may["TEMPO EFETIVO PRODUÇÃO/PARADA"].dropna()
        if not s.empty:
            print("First value:", s.iloc[0], "Type:", type(s.iloc[0]))
            print("Second value:", s.iloc[1], "Type:", type(s.iloc[1]))
            
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
