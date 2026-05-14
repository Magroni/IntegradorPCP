import pandas as pd
df = pd.read_excel('z:/PCP/PROJETOS MARLON/ProgramarProd/COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm', sheet_name='BASE DE DADOS', engine='openpyxl')
print("Columns:", df.columns.tolist())
print(df.head(20).to_string())
