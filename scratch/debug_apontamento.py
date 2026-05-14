import pandas as pd
import data_manager as dm
from datetime import datetime

print("--- DEBUG APONTAMENTO ---")
file_path = dm._get_apontamento_file()
print(f"Arquivo: {file_path}")

try:
    # Tenta ler as primeiras 15 linhas para ver a cara do arquivo
    df_raw = pd.read_excel(file_path, sheet_name=dm._get_sheet("SHEET_AP_BD"), header=None, nrows=15)
    print("\nPrimeiras 15 linhas (brutas):")
    print(df_raw.to_string())
    
    # Tenta rodar a função oficial para o dia 12/05/2026
    data_alvo = datetime(2026, 5, 12).date()
    df_dia = dm.get_apontamentos_do_dia(data_alvo)
    print(f"\nResultado get_apontamentos_do_dia para {data_alvo}:")
    if df_dia.empty:
        print("VAZIO")
        # Se vazio, vamos ver todas as datas que existem no arquivo
        df_all = pd.read_excel(file_path, sheet_name=dm._get_sheet("SHEET_AP_BD"), skiprows=6)
        if not df_all.empty:
            # Normalizar colunas como no dm
            df_all.columns = [str(c).strip().upper() for c in df_all.columns]
            if "DATA REG" in df_all.columns:
                datas = pd.to_datetime(df_all["DATA REG"], errors="coerce").dropna().dt.date.unique()
                print(f"Datas encontradas no arquivo: {datas}")
            else:
                print(f"Coluna 'DATA REG' não encontrada. Colunas disponíveis: {df_all.columns.tolist()}")
    else:
        print(df_dia.head(10).to_string())

except Exception as e:
    print(f"ERRO: {e}")
