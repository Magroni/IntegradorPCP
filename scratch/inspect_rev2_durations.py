import pandas as pd
import os
import shutil

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")
TEMP_REV2 = os.path.join(BASE_DIR, "scratch", "temp_rev2.xlsx")

if os.path.exists(FILE_REV2):
    try:
        shutil.copy(FILE_REV2, TEMP_REV2)
        print("Successfully copied file to temporary location.")
        
        df_db = pd.read_excel(TEMP_REV2, sheet_name="DB")
        df_paradas = pd.read_excel(TEMP_REV2, sheet_name="PARADAS")
        
        print("\n--- TEMPLATE DB (Production) TEMPO_PROCESSO ---")
        print(df_db[["ID", "HORA_INICIO", "HORA_FIM", "TEMPO_PROCESSO"]].head(15))
        print("Datatype of TEMPO_PROCESSO in template DB:", df_db["TEMPO_PROCESSO"].dtype)
        if len(df_db["TEMPO_PROCESSO"]) > 0:
            print("First non-null value type:", type(df_db["TEMPO_PROCESSO"].dropna().iloc[0]), "Value:", df_db["TEMPO_PROCESSO"].dropna().iloc[0])
        
        print("\n--- TEMPLATE PARADAS TEMPO ---")
        print(df_paradas[["ID_APONTAMENTO", "MOTIVO", "HORA_INICIO", "HORA_FIM", "TEMPO"]].head(15))
        print("Datatype of TEMPO in template PARADAS:", df_paradas["TEMPO"].dtype)
        if len(df_paradas["TEMPO"]) > 0:
            print("First non-null value type:", type(df_paradas["TEMPO"].dropna().iloc[0]), "Value:", df_paradas["TEMPO"].dropna().iloc[0])
            
        # Clean up
        if os.path.exists(TEMP_REV2):
            os.remove(TEMP_REV2)
            
    except Exception as e:
        print("Error:", e)
else:
    print("REV 2 template file does not exist")
