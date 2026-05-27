import pandas as pd
import os
import datetime as dt

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

def calculate_duration_hhmm(ini, fim):
    if pd.isna(ini) or pd.isna(fim):
        return ""
    
    # Ensure they are datetime.time objects
    def parse_to_time(val):
        if isinstance(val, dt.time):
            return val
        try:
            t_str = str(val).strip()
            parts = t_str.split(":")
            h = int(parts[0])
            m = int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            return dt.time(h, m, s)
        except:
            return None
            
    t_ini = parse_to_time(ini)
    t_fim = parse_to_time(fim)
    
    if t_ini is None or t_fim is None:
        return ""
        
    d = dt.date(2026, 5, 1)
    dt_ini = dt.datetime.combine(d, t_ini)
    dt_fim = dt.datetime.combine(d, t_fim)
    
    if dt_fim < dt_ini:
        dt_fim += dt.timedelta(days=1)
        
    diff = dt_fim - dt_ini
    total_seconds = int(diff.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    return f"{h:02d}:{m:02d}"

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
        
        print("Sample of Calculated Durations for Production Rows:")
        df_prod = df_may[df_may["PROCESSO"] != "PARADA"].copy()
        for idx, row in df_prod.head(10).iterrows():
            dur = calculate_duration_hhmm(row["HORIM. INI"], row["HORIM. FIM"])
            print(f"Prod: {row['PROCESSO']:<20} | Ini: {row['HORIM. INI']} | Fim: {row['HORIM. FIM']} | Calc Dur: {dur}")
            
        print("\nSample of Calculated Durations for Stop Rows:")
        df_paradas = df_may[df_may["PROCESSO"] == "PARADA"].copy()
        for idx, row in df_paradas.head(10).iterrows():
            dur = calculate_duration_hhmm(row["HORIM. INI"], row["HORIM. FIM"])
            print(f"Stop: {row['MATERIAL+BLOCO'][:20]:<20} | Ini: {row['HORIM. INI']} | Fim: {row['HORIM. FIM']} | Calc Dur: {dur}")
            
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
