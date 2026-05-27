import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")

if os.path.exists(FILE_REV2):
    try:
        df_paradas = pd.read_excel(FILE_REV2, sheet_name="PARADAS")
        print("Total rows in REV 2 PARADAS:", len(df_paradas))
        print("Null or NA ID_APONTAMENTO in REV 2 PARADAS:", df_paradas["ID_APONTAMENTO"].isna().sum())
        print("First 10 rows of REV 2 PARADAS:")
        print(df_paradas.head(10))
    except Exception as e:
        print("Error reading REV 2 PARADAS:", e)
else:
    print("REV 2 file does not exist")
