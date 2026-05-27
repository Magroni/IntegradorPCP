import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_OUTPUT = os.path.join(BASE_DIR, "Apontamento_Maio_Transposto_Auditoria.xlsx")

if os.path.exists(FILE_OUTPUT):
    try:
        df_paradas = pd.read_excel(FILE_OUTPUT, sheet_name="PARADAS")
        df_db = pd.read_excel(FILE_OUTPUT, sheet_name="DB")
        
        # Merge to see details of linked processes
        df_merged = df_paradas.merge(df_db, left_on="ID_APONTAMENTO", right_on="ID", how="left", suffixes=("_parada", "_processo"))
        
        print("Joined Stops and Processes:")
        print(df_merged[["ID_APONTAMENTO", "MOTIVO", "SETOR", "PROCESSO", "DATA_INICIO_parada", "HORA_INICIO_parada", "DATA_INICIO_processo"]])
        
        print("\nStops with missing or unlinked processes:")
        missing = df_merged[df_merged["ID"].isna()]
        print(missing)
        
    except Exception as e:
        print("Error:", e)
else:
    print("File does not exist")
