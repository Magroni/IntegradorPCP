import pandas as pd
import os

def check():
    file_path = 'Apontamento Produção (REV 2).xlsx'
    df = pd.read_excel(file_path, sheet_name='DB', engine='openpyxl')
    
    # Busca por CIMEF
    cimef = df[df['SETOR'].astype(str).str.upper().str.strip() == 'CIMEF']
    print(f"Total de registros para CIMEF: {len(cimef)}")
    
    # Filtra por datas de inicio/fim contendo 25 ou 26
    cimef_target = cimef[
        cimef['DATA_INICIO'].astype(str).str.contains('25/05/2026|25-05-2026|2026-05-25') |
        cimef['DATA_FIM'].astype(str).str.contains('26/05/2026|26-05-2026|2026-05-26')
    ]
    print(f"Registros CIMEF em 25/05 ou 26/05: {len(cimef_target)}")
    
    for idx, row in cimef_target.iterrows():
        print(f"\nRow {idx}:")
        for col in df.columns:
            print(f"  {col}: {row[col]} ({type(row[col])})")

if __name__ == "__main__":
    check()
