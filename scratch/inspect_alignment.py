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
        
        df_prod_raw = df_may[~df_may["IS_PARADA"]].copy()
        df_paradas_raw = df_may[df_may["IS_PARADA"]].copy()
        
        start_id = 18999
        df_prod_raw["NEW_ID"] = list(range(start_id, start_id + len(df_prod_raw)))
        
        print("df_prod_raw columns and index sample:")
        print(df_prod_raw[["NEW_ID", "PROCESSO", "SETOR"]].head(15))
        
        # Trace row 18830
        idx = 18830
        row = df_paradas_raw.loc[idx]
        setor_parada = str(row.get("SETOR", "")).strip().upper()
        
        print(f"\nTracing Stop at index {idx}:")
        print("MOTIVO:", row["MATERIAL+BLOCO"])
        print("SETOR:", setor_parada)
        print("DATE:", row["DATETIME_CONVERTED"])
        
        df_same_sector = df_prod_raw[df_prod_raw["SETOR"].fillna("").astype(str).str.strip().str.upper() == setor_parada]
        print(f"\nAll processes in the same sector ({setor_parada}):")
        print(df_same_sector[["NEW_ID", "DATETIME_CONVERTED", "PROCESSO", "MATERIAL+BLOCO"]])
        
        distances = pd.Series((df_same_sector.index - idx), index=df_same_sector.index).abs()
        print("\nCalculated distances:")
        print(distances)
        
        idx_closest = distances.idxmin()
        print("\nidxmin:", idx_closest)
        print("Closest process NEW_ID:", df_same_sector.loc[idx_closest, "NEW_ID"])
        
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
