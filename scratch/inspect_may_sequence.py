import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

if os.path.exists(FILE_REV1):
    try:
        df_rev1 = pd.read_excel(FILE_REV1, sheet_name="BD", skiprows=6, engine="openpyxl")
        df_rev1.columns = [str(c).strip().upper() for c in df_rev1.columns]
        
        # Sort or preserve order
        col_data_filt = "DATA REG" if "DATA REG" in df_rev1.columns else df_rev1.columns[0]
        df_rev1["DATETIME_CONVERTED"] = pd.to_datetime(df_rev1[col_data_filt], errors="coerce")
        
        # Filter May 2026
        df_may = df_rev1[
            (df_rev1["DATETIME_CONVERTED"].dt.month == 5) & 
            (df_rev1["DATETIME_CONVERTED"].dt.year == 2026)
        ].copy()
        
        def is_parada(row):
            proc = str(row.get("PROCESSO", "")).strip().upper()
            mat_blo = str(row.get("MATERIAL+BLOCO", "")).strip().upper()
            if proc == "PARADA":
                return True
            parada_keywords = [
                'ALMOÇO', 'ALMOCO', 'ABRASIVO', 'PARADA', 'TROCA', 'AJUSTE', 'SETUP', 
                'ABASTECER', 'AGUARDANDO', 'INTERVALO', 'MANUTENÇÃO', 'MANUTENCAO', 
                'FALTA', 'JANTA', 'ELÉTRICA', 'ELETRICA',
                'CHECK', 'CHECKLIST', 'LIXAMENTO', 'LEVANTAMENTO', 'LEVANTANDO'
            ]
            for kw in parada_keywords:
                if kw in mat_blo or kw in proc:
                    return True
            return False
            
        df_may["IS_PARADA"] = df_may.apply(is_parada, axis=1)
        
        print("First 15 rows of df_may, including their index, PROCESSO, MATERIAL+BLOCO, SETOR, IS_PARADA:")
        print(df_may[["PROCESSO", "MATERIAL+BLOCO", "SETOR", "IS_PARADA"]].head(25))
        
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
