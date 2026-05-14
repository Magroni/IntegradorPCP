import pandas as pd
try:
    df = pd.read_excel('Estoque Chapas 2026.xlsx', sheet_name='ENTRADAS')
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_bloco = next((c for c in ["BLOCO", "Nº BLOCO", "N_BLOCO", "NUMERO DO BLOCO", "MAT+BLO"] if c in df.columns), None)
    if col_bloco:
        df["N_BLOCO_STR"] = df[col_bloco].astype(str).str.strip().str.split(".").str[0].str.upper()
        match = df[df["N_BLOCO_STR"] == "1654"]
        if not match.empty:
            print("Block 1654 Found:")
            print(match.iloc[0].to_dict())
        else:
            print("Block 1654 NOT FOUND in the whole sheet.")
except Exception as e:
    print("Error:", e)
