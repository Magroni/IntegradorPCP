import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_OUTPUT = os.path.join(BASE_DIR, "Apontamento_Maio_Transposto_Auditoria.xlsx")

if os.path.exists(FILE_OUTPUT):
    try:
        df_db = pd.read_excel(FILE_OUTPUT, sheet_name="DB")
        df_paradas = pd.read_excel(FILE_OUTPUT, sheet_name="PARADAS")
        
        print("--- DB (Production) sheet samples ---")
        print(df_db[["ID", "PROCESSO", "HORA_INICIO", "HORA_FIM", "TEMPO_PROCESSO"]].head(15))
        
        print("\nNull counts in DB:")
        print(df_db[["HORA_INICIO", "HORA_FIM", "TEMPO_PROCESSO"]].isna().sum())
        
        print("\n--- PARADAS sheet samples ---")
        print(df_paradas[["ID_APONTAMENTO", "MOTIVO", "HORA_INICIO", "HORA_FIM", "TEMPO"]].head(15))
        
        print("\nNull counts in PARADAS:")
        print(df_paradas[["HORA_INICIO", "HORA_FIM", "TEMPO"]].isna().sum())
        
    except Exception as e:
        print("Error:", e)
else:
    print("Output file does not exist")
