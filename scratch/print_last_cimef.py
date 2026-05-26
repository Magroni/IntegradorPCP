import pandas as pd
import os

def check():
    file_path = 'Apontamento Produção (REV 2).xlsx'
    df = pd.read_excel(file_path, sheet_name='DB', engine='openpyxl')
    
    # Filtra por CIMEF
    cimef = df[df['SETOR'].astype(str).str.upper().str.strip() == 'CIMEF']
    print(f"Total de registros para CIMEF: {len(cimef)}")
    print("\nUltimos 10 registros de CIMEF na planilha:")
    
    last_10 = cimef.tail(10)
    for idx, row in last_10.iterrows():
        print(f"\nRow {idx}:")
        print(f"  ID: {row.get('ID')}")
        print(f"  MATERIAL: {row.get('MATERIAL')}")
        print(f"  BLOCO: {row.get('NUMERO_BLOCO')}")
        print(f"  PROCESSO: {row.get('PROCESSO')}")
        print(f"  DATA_INICIO: {row.get('DATA_INICIO')}")
        print(f"  DATA_FIM: {row.get('DATA_FIM')}")
        print(f"  HORA_INICIO: {row.get('HORA_INICIO')}")
        print(f"  HORA_FIM: {row.get('HORA_FIM')}")
        print(f"  TURNO: {row.get('TURNO')}")

if __name__ == "__main__":
    check()
