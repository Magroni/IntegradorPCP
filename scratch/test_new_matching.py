import pandas as pd
import os
import datetime as dt

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

def normalize_sector(s):
    s = str(s).strip().upper()
    if "5-POLITRIZ" in s:
        return "5-POLITRIZ"
    if "3-POLITRIZ" in s:
        return "3-POLITRIZ"
    if "2-RESINA" in s:
        return "2-RESINA"
    if "RETOQUE" in s:
        return "RETOQUE"
    return s

def make_datetime(date_val, time_val):
    if pd.isna(date_val) or pd.isna(time_val):
        return None
    if hasattr(date_val, "to_pydatetime"):
        d = date_val.to_pydatetime().date()
    elif isinstance(date_val, dt.datetime):
        d = date_val.date()
    else:
        d = date_val
    
    # In case time_val is not a datetime.time object
    if not isinstance(time_val, dt.time):
        try:
            # try to parse HH:MM:SS
            t_str = str(time_val).strip()
            parts = t_str.split(":")
            h = int(parts[0])
            m = int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            time_val = dt.time(h, m, s)
        except:
            return None
            
    return dt.datetime.combine(d, time_val)

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
        
        # Build datetime for production rows
        df_prod_raw["PROD_DATETIME"] = df_prod_raw.apply(
            lambda r: make_datetime(r["DATETIME_CONVERTED"], r["HORIM. INI"]), axis=1
        )
        
        # Build datetime for stops rows
        df_paradas_raw["STOP_DATETIME"] = df_paradas_raw.apply(
            lambda r: make_datetime(r["DATETIME_CONVERTED"], r["HORIM. INI"]), axis=1
        )
        
        matched_results = []
        for idx, row in df_paradas_raw.iterrows():
            motivo = str(row.get("MATERIAL+BLOCO", "PARADA")).strip().upper()
            setor_parada = str(row.get("SETOR", "")).strip().upper()
            stop_dt = row["STOP_DATETIME"]
            
            norm_setor = normalize_sector(setor_parada)
            
            # Filter production by same normalized sector
            df_same_sector = df_prod_raw[
                df_prod_raw["SETOR"].apply(normalize_sector) == norm_setor
            ].copy()
            
            id_apontamento = pd.NA
            matched_proc_name = "N/A"
            matched_proc_date = "N/A"
            matched_proc_time = "N/A"
            matched_proc_sector = "N/A"
            
            if not df_same_sector.empty and stop_dt is not None:
                # Calculate absolute time difference in seconds
                df_same_sector["TIME_DIFF"] = (df_same_sector["PROD_DATETIME"] - stop_dt).apply(lambda x: abs(x.total_seconds()))
                closest_idx = df_same_sector["TIME_DIFF"].idxmin()
                closest_row = df_same_sector.loc[closest_idx]
                id_apontamento = closest_row["NEW_ID"]
                matched_proc_name = closest_row["PROCESSO"]
                matched_proc_date = str(closest_row["DATETIME_CONVERTED"])[:10]
                matched_proc_time = str(closest_row["HORIM. INI"])
                matched_proc_sector = closest_row["SETOR"]
            elif stop_dt is not None:
                # Fallback to absolute closest time in the entire dataset
                df_temp = df_prod_raw.copy()
                df_temp["TIME_DIFF"] = (df_temp["PROD_DATETIME"] - stop_dt).apply(lambda x: abs(x.total_seconds()))
                closest_idx = df_temp["TIME_DIFF"].idxmin()
                closest_row = df_temp.loc[closest_idx]
                id_apontamento = closest_row["NEW_ID"]
                matched_proc_name = closest_row["PROCESSO"]
                matched_proc_date = str(closest_row["DATETIME_CONVERTED"])[:10]
                matched_proc_time = str(closest_row["HORIM. INI"])
                matched_proc_sector = closest_row["SETOR"]
                
            matched_results.append({
                "IDX": idx,
                "STOP_MOTIVO": motivo,
                "STOP_SETOR": setor_parada,
                "STOP_DATE": str(row["DATETIME_CONVERTED"])[:10],
                "STOP_TIME": str(row["HORIM. INI"]),
                "MATCHED_ID": id_apontamento,
                "MATCHED_PROC": matched_proc_name,
                "MATCHED_DATE": matched_proc_date,
                "MATCHED_TIME": matched_proc_time,
                "MATCHED_SECTOR": matched_proc_sector
            })
            
        df_matched = pd.DataFrame(matched_results)
        print(df_matched.head(45))
        
        # Check if there are any mismatches in sector or extremely large time gaps
        print("\nChecking any matched processes with more than 2 hours time difference on same sector:")
        # We can calculate the actual difference for auditing
        
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
