import pandas as pd
try:
    df = pd.read_excel(r"z:\PCP\PROJETOS MARLON\ProgramarProd\Estoque Chapas 2026.xlsx", sheet_name="ENTRADAS", nrows=10)
    print("Columns:", df.columns.tolist())
    print("Head:")
    print(df.head())
except Exception as e:
    print("Error:", e)
