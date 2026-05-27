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
        
        # Filter May 2026
        df_may = df_rev1[
            (df_rev1["DATETIME_CONVERTED"].dt.month == 5) & 
            (df_rev1["DATETIME_CONVERTED"].dt.year == 2026)
        ].copy()
        
        print("All raw rows in May 2026 ordered chronologically:")
        print(f"{'IDX':<5} | {'DATA REG':<12} | {'PROCESSO':<20} | {'MAT+BLO':<30} | {'SETOR':<10} | {'HORA INI':<8} | {'HORA FIM':<8}")
        print("-" * 110)
        
        for idx, row in df_may.iterrows():
            print(f"{idx:<5} | {str(row.get('DATA REG'))[:10]:<12} | {str(row.get('PROCESSO'))[:20]:<20} | {str(row.get('MATERIAL+BLOCO'))[:30]:<30} | {str(row.get('SETOR'))[:10]:<10} | {str(row.get('HORIM. INI')):<8} | {str(row.get('HORIM. FIM')):<8}")
            
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
