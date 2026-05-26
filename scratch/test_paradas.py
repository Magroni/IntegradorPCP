import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import pandas as pd
from datetime import datetime, timedelta
import data_manager as dm

def format_to_hhmm(minutes):
    if pd.isna(minutes) or minutes <= 0:
        return "00:00"
    total_mins = int(round(minutes))
    h = total_mins // 60
    m = total_mins % 60
    return f"{h:02d}:{m:02d}"

def run_test():
    try:
        df_raw_an = dm.get_all_apontamentos()
        mapping_an = dm._get_system_mapping()
        
        def find_col_an(key):
            for alias in mapping_an.get(key, []):
                if alias.upper() in df_raw_an.columns: return alias.upper()
            return None

        c_dt = find_col_an("DIA_FIM") or find_col_an("DIA_INICIO")
        c_hr = find_col_an("HORA_FIM") or find_col_an("HORA_INICIO")
        c_dt_ini = find_col_an("DIA_INICIO")
        c_hr_ini = find_col_an("HORA_INICIO")
        c_st = find_col_an("SETOR_AP")
        col_id_ap = find_col_an("ID") or "ID"

        c_st_effective = c_st or "SETOR"
        cols_to_use = [col_id_ap]
        if c_st_effective and c_st_effective in df_raw_an.columns:
            cols_to_use.append(c_st_effective)
        
        for col_name in [c_dt, c_dt_ini, c_hr, c_hr_ini]:
            if col_name and col_name in df_raw_an.columns and col_name not in cols_to_use:
                cols_to_use.append(col_name)
                
        df_raw_subset = df_raw_an[cols_to_use].copy()
        
        if c_dt and c_dt in df_raw_subset.columns:
            df_raw_subset[c_dt] = pd.to_datetime(df_raw_subset[c_dt], errors='coerce', dayfirst=True)
        if c_dt_ini and c_dt_ini in df_raw_subset.columns and c_dt_ini != c_dt:
            df_raw_subset[c_dt_ini] = pd.to_datetime(df_raw_subset[c_dt_ini], errors='coerce', dayfirst=True)
            
        def get_dia_producao(row):
            dt = row[c_dt]
            val_hr = row.get(c_hr)
            if pd.isna(dt) and c_dt_ini:
                dt = row[c_dt_ini]
                val_hr = row.get(c_hr_ini)
            if pd.isna(dt): return None
            try:
                hour = val_hr.hour if hasattr(val_hr, 'hour') else int(str(val_hr).split(":")[0])
                if hour < 7: return (dt - timedelta(days=1)).date()
                return dt.date()
            except: return dt.date()

        df_raw_subset["DIA_PROD"] = df_raw_subset.apply(get_dia_producao, axis=1)
        df_raw_an_subset = df_raw_subset[[col_id_ap, c_st_effective, "DIA_PROD"]].copy()
        
        # Load paradas
        ap_file = dm._get_apontamento_file()
        sheet_p = dm._get_sheet("SHEET_AP_PARADAS")
        df_paradas = pd.read_excel(ap_file, sheet_name=sheet_p)
        df_paradas.columns = [str(c).strip().upper() for c in df_paradas.columns]
        
        def parse_downtime_minutes(row):
            t_val = row.get("TEMPO")
            if pd.notna(t_val) and str(t_val).strip() != "":
                t_str = str(t_val).strip()
                if ":" in t_str:
                    try:
                        parts = t_str.split(":")
                        return int(parts[0]) * 60 + int(parts[1])
                    except:
                        pass
                try:
                    num = float(t_str.replace(",", "."))
                    if num > 0:
                        if num < 1.0:
                            return int(round(num * 24 * 60))
                        return int(round(num))
                except:
                    pass
            return 0

        df_paradas["TEMPO"] = df_paradas.apply(parse_downtime_minutes, axis=1)
        df_paradas_m = df_paradas.merge(df_raw_an_subset, left_on="ID_APONTAMENTO", right_on=col_id_ap, how="inner")
        
        df_paradas_m["TEMPO_HHMM"] = df_paradas_m["TEMPO"].apply(format_to_hhmm)
        print("First 5 rows of df_paradas_m with TEMPO_HHMM:\n", df_paradas_m[["ID_APONTAMENTO", "MOTIVO", "TEMPO", "TEMPO_HHMM"]].head())

    except Exception as e:
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_test()
