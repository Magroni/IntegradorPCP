import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import pandas as pd
from datetime import datetime, timedelta
import data_manager as dm

def check():
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
    c_m2 = find_col_an("QTD_M2")
    c_ch = find_col_an("QTD_CH")
    c_pr = find_col_an("PROCESSO_APONTADO")

    df_an = df_raw_an.copy()
    for col_num in [c_m2, c_ch]:
        df_an[col_num] = pd.to_numeric(df_an[col_num], errors='coerce').fillna(0)

    if c_dt:
        df_an[c_dt] = pd.to_datetime(df_an[c_dt], errors='coerce', dayfirst=True)
    if c_dt_ini and c_dt_ini != c_dt:
        df_an[c_dt_ini] = pd.to_datetime(df_an[c_dt_ini], errors='coerce', dayfirst=True)
    
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
        
    df_an["DIA_PROD"] = df_an.apply(get_dia_producao, axis=1)
    
    # Vamos buscar especificamente o registro de ID 18995
    row_119 = df_an[df_an["ID"] == 18995]
    if row_119.empty:
        print("Erro: ID 18995 nao encontrado!")
        return
        
    row = row_119.iloc[0]
    print("Detalhes do registro 18995 apos processamento:")
    print("  ID:", row["ID"])
    print("  BLOCO:", row.get("NUMERO_BLOCO"))
    print("  PROCESSO:", row.get("PROCESSO"))
    print("  SETOR:", row.get("SETOR"))
    print("  DATA_FIM (c_dt):", row[c_dt])
    print("  HORA_FIM (c_hr):", row[c_hr])
    print("  DIA_PROD:", row["DIA_PROD"])
    
    # Agora vamos ver se ele sobrevive aos filtros dos indicadores
    hoje = datetime(2026, 5, 26).date() # Simula hoje
    df_an_filtered = df_an.dropna(subset=["DIA_PROD"])
    
    # Filtro de periodo
    df_an_7 = df_an_filtered[df_an_filtered["DIA_PROD"] >= hoje - timedelta(days=7)]
    print("\nFiltro 'Ultimos 7 dias' (hoje = 2026-05-26):")
    print("  Sobreviveu?", 18995 in df_an_7["ID"].values)
    
    # Filtro de processo RETOQUE
    df_an_ret = df_an_7[~df_an_7[c_pr].astype(str).str.upper().str.contains("RETOQUE")]
    print("Filtro RETOQUE:")
    print("  Sobreviveu?", 18995 in df_an_ret["ID"].values)

if __name__ == "__main__":
    check()
