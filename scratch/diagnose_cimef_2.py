import pandas as pd
import os

def diagnose():
    file_path = 'Apontamento Produção (REV 2).xlsx'
    if not os.path.exists(file_path):
        print(f"Erro: Arquivo {file_path} nao existe.")
        return
        
    df = pd.read_excel(file_path, sheet_name='DB', engine='openpyxl')
    print("Colunas encontradas:", list(df.columns))
    
    # Vamos buscar linhas onde o setor contem CIMEF
    cimef_rows = df[df['SETOR'].astype(str).str.upper().str.contains('CIMEF')]
    print(f"\nEncontradas {len(cimef_rows)} linhas para o setor CIMEF.")
    
    for idx, row in cimef_rows.iterrows():
        print(f"\n--- Linha {idx} ---")
        print(f"NUMERO_BLOCO: {row.get('NUMERO_BLOCO')} (tipo: {type(row.get('NUMERO_BLOCO'))})")
        print(f"PROCESSO: {row.get('PROCESSO')} (tipo: {type(row.get('PROCESSO'))})")
        print(f"DATA_INICIO: {row.get('DATA_INICIO')} (tipo: {type(row.get('DATA_INICIO'))})")
        print(f"DATA_FIM: {row.get('DATA_FIM')} (tipo: {type(row.get('DATA_FIM'))})")
        print(f"HORA_INICIO: {row.get('HORA_INICIO')} (tipo: {type(row.get('HORA_INICIO'))})")
        print(f"HORA_FIM: {row.get('HORA_FIM')} (tipo: {type(row.get('HORA_FIM'))})")
        print(f"TURNO: {row.get('TURNO')} (tipo: {type(row.get('TURNO'))})")

if __name__ == "__main__":
    diagnose()
