import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import pandas as pd
from datetime import datetime, timedelta
import data_manager as dm

def test_refeito_calculation():
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

    if df_raw_an.empty or not c_dt or not c_m2 or not c_ch:
        print("Erro: Colunas ou dados não encontrados.")
        return

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
    df_an = df_an.dropna(subset=["DIA_PROD"])

    # Simular os períodos
    print(f"Total de linhas no apontamento com DIA_PROD válido: {len(df_an)}")
    
    df_an = df_an[~df_an[c_pr].astype(str).str.upper().str.contains("RETOQUE")]
    df_an["REFEITO"] = df_an[c_pr].astype(str).str.upper().str.contains("REPASSE|REFEITO|REPROCESSO")
    
    m_tot_m2 = df_an[c_m2].sum()
    m_tot_ch = df_an[c_ch].sum()
    m_refeito_m2 = df_an[df_an["REFEITO"]][c_m2].sum()
    m_refeito_ch = df_an[df_an["REFEITO"]][c_ch].sum()
    
    print("\n--- TOTAIS GERAIS ---")
    print(f"M² Total: {m_tot_m2:.2f}")
    print(f"M² Refeito: {m_refeito_m2:.2f}")
    print(f"M² Normal: {(m_tot_m2 - m_refeito_m2):.2f}")
    p_refeito_m2_total = (m_refeito_m2 / m_tot_m2 * 100) if m_tot_m2 > 0 else 0
    p_refeito_m2_normal = (m_refeito_m2 / (m_tot_m2 - m_refeito_m2) * 100) if (m_tot_m2 - m_refeito_m2) > 0 else 0
    print(f"%% Refeito em M² (base Total): {p_refeito_m2_total:.2f}%%")
    print(f"%% Refeito em M² (base Normal): {p_refeito_m2_normal:.2f}%%")
    
    print("\nChapas Totais: ", m_tot_ch)
    print("Chapas Refeito: ", m_refeito_ch)
    print("Chapas Normal: ", m_tot_ch - m_refeito_ch)
    p_refeito_ch_total = (m_refeito_ch / m_tot_ch * 100) if m_tot_ch > 0 else 0
    p_refeito_ch_normal = (m_refeito_ch / (m_tot_ch - m_refeito_ch) * 100) if (m_tot_ch - m_refeito_ch) > 0 else 0
    print(f"%% Refeito em Chapas (base Total): {p_refeito_ch_total:.2f}%%")
    print(f"%% Refeito em Chapas (base Normal): {p_refeito_ch_normal:.2f}%%")

if __name__ == "__main__":
    test_refeito_calculation()
