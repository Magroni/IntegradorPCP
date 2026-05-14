import pandas as pd
try:
    # Try reading with pyxlsb if available
    df = pd.read_excel(r"z:\PCP\PROJETOS MARLON\ProgramarProd\PLANILHA BLOCOS.xlsb", engine="pyxlsb")
    print("Columns:", df.columns.tolist())
    print("First 5 rows:")
    print(df.head())
except Exception as e:
    print("Error reading with pyxlsb:", e)
    try:
        # Fallback to default engine just in case it's actually not xlsb or pandas handles it
        df = pd.read_excel(r"z:\PCP\PROJETOS MARLON\ProgramarProd\PLANILHA BLOCOS.xlsb")
        print("Columns (fallback):", df.columns.tolist())
    except Exception as e2:
        print("Error with default engine:", e2)
