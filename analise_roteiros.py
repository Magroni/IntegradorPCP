import pandas as pd
df = pd.read_excel('COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm', sheet_name='PROGRAMAÇÃO', engine='openpyxl')
df['BLOCO'] = df['BLOCO'].astype(str)
df = df.dropna(subset=['PROCESSO'])
routes = df.groupby('BLOCO')['PROCESSO'].apply(list)
# Para contar listas no pandas de forma fácil, converte pra tuple
print(routes.apply(tuple).value_counts().head(15).to_string())
