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
        df_paradas = df_may[df_may["IS_PARADA"]].copy()
        
        print("Stop rows with all columns that might contain durations:")
        time_cols = [c for c in df_paradas.columns if any(kw in c for kw in ["TEMPO", "HORA", "MIN", "DUR", "HORIM"])]
        print("Time columns found in sheet:", time_cols)
        print("\nFirst 15 stops:")
        print(df_paradas[["PROCESSO", "MATERIAL+BLOCO", "HORIM. INI", "HORIM. FIM"] + [c for c in ["TEMPO EFETIVO PRODUÇÃO/PARADA", "TEMPO EFETIVO PRODUÇÃO/PARADA 2"] if c in df_paradas.columns]].head(15))
        
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
