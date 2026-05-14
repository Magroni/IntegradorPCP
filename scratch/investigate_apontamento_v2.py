import pandas as pd
try:
    xl = pd.ExcelFile('Apontamento Produção (REV 2).xlsx')
    print('Sheets:', xl.sheet_names)
    df_bd = pd.read_excel('Apontamento Produção (REV 2).xlsx', sheet_name='DB', header=6) # Assuming header in row 7
    print('DB Columns:', df_bd.columns.tolist())
    if 'PARADAS' in xl.sheet_names:
        df_paradas = pd.read_excel('Apontamento Produção (REV 2).xlsx', sheet_name='PARADAS')
        print('PARADAS Columns:', df_paradas.columns.tolist())
except Exception as e:
    print('Error:', e)
