import pandas as pd
import os

def check():
    file_path = 'Apontamento Produção (REV 2).xlsx'
    df = pd.read_excel(file_path, sheet_name='DB', engine='openpyxl')
    
    print("Mapeando linhas recentes a partir do índice 100:")
    for idx in range(100, len(df)):
        row = df.iloc[idx]
        print(f"Row {idx}: ID={row.get('ID')} BLOCO={row.get('NUMERO_BLOCO')} SETOR={row.get('SETOR')} PROCESSO={row.get('PROCESSO')} INI={row.get('DATA_INICIO')} FIM={row.get('DATA_FIM')} H_INI={row.get('HORA_INICIO')} H_FIM={row.get('HORA_FIM')}")

if __name__ == "__main__":
    check()
