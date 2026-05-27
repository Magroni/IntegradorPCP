import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_OUTPUT = os.path.join(BASE_DIR, "Apontamento_Maio_Transposto_Auditoria.xlsx")

if os.path.exists(FILE_OUTPUT):
    try:
        df_paradas = pd.read_excel(FILE_OUTPUT, sheet_name="PARADAS")
        df_db = pd.read_excel(FILE_OUTPUT, sheet_name="DB")
        
        df_merged = df_paradas.merge(df_db, left_on="ID_APONTAMENTO", right_on="ID", how="left", suffixes=("_parada", "_processo"))
        
        print(f"{'IDX':<4} | {'MOTIVO':<25} | {'SETOR':<10} | {'DATA STOP':<12} | {'DATA PROC':<12} | {'ID PROC':<8} | {'PROC NAME':<20}")
        print("-" * 100)
        for idx, row in df_merged.iterrows():
            print(f"{idx:<4} | {str(row['MOTIVO'])[:25]:<25} | {str(row['SETOR'])[:10]:<10} | {str(row['DATA_INICIO_parada']):<12} | {str(row['DATA_INICIO_processo']):<12} | {str(row['ID_APONTAMENTO']):<8} | {str(row['PROCESSO'])[:20]:<20}")
            
    except Exception as e:
        print("Error:", e)
else:
    print("File does not exist")
