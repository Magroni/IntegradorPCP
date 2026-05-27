import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_OUTPUT = os.path.join(BASE_DIR, "Apontamento_Maio_Transposto_Auditoria.xlsx")

if os.path.exists(FILE_OUTPUT):
    try:
        df_paradas = pd.read_excel(FILE_OUTPUT, sheet_name="PARADAS")
        df_db = pd.read_excel(FILE_OUTPUT, sheet_name="DB")
        
        print("--- PARADAS SHEET ---")
        print(df_paradas.head(45))
        
        print("\n--- STATS ---")
        print("Total rows in PARADAS:", len(df_paradas))
        print("Null or NA ID_APONTAMENTO:", df_paradas["ID_APONTAMENTO"].isna().sum())
        print("Non-null ID_APONTAMENTO:", df_paradas["ID_APONTAMENTO"].notna().sum())
        
        if df_paradas["ID_APONTAMENTO"].notna().sum() > 0:
            print("\nUnique ID_APONTAMENTO in PARADAS:", df_paradas["ID_APONTAMENTO"].dropna().unique())
            print("Min ID_APONTAMENTO in PARADAS:", df_paradas["ID_APONTAMENTO"].min())
            print("Max ID_APONTAMENTO in PARADAS:", df_paradas["ID_APONTAMENTO"].max())
            print("Range of IDs in DB sheet:", df_db["ID"].min(), "to", df_db["ID"].max())
            
    except Exception as e:
        print("Error reading output sheet:", e)
else:
    print("Output file does not exist at:", FILE_OUTPUT)
