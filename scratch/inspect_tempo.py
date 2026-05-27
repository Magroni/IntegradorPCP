import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")

if os.path.exists(FILE_REV1):
    try:
        df_rev1 = pd.read_excel(FILE_REV1, sheet_name="BD", skiprows=6, engine="openpyxl")
        df_rev1.columns = [str(c).strip().upper() for c in df_rev1.columns]
        
        print("Columns containing 'TEMPO' or 'HORA' or 'MIN' or 'DUR':")
        matching_cols = [c for c in df_rev1.columns if any(kw in c for kw in ["TEMPO", "HORA", "MIN", "DUR", "HORIM"])]
        print(matching_cols)
        
        # Filter May 2026
        col_data_filt = "DATA REG" if "DATA REG" in df_rev1.columns else df_rev1.columns[0]
        df_rev1["DATETIME_CONVERTED"] = pd.to_datetime(df_rev1[col_data_filt], errors="coerce")
        df_may = df_rev1[
            (df_rev1["DATETIME_CONVERTED"].dt.month == 5) & 
            (df_rev1["DATETIME_CONVERTED"].dt.year == 2026)
        ].copy()
        
        print("\nFirst 10 rows of May 2026 for time columns:")
        cols_to_show = ["PROCESSO", "SETOR", "HORIM. INI", "HORIM. FIM"]
        for c in ["TEMPO EFETIVO PRODUÇÃO/PARADA", "TEMPO", "DURAÇÃO", "DURACAO"]:
            if c in df_may.columns:
                cols_to_show.append(c)
        print(df_may[cols_to_show].head(15))
        
        print("\nTypes of 'HORIM. INI' and 'HORIM. FIM':")
        print("HORIM. INI first value:", df_may["HORIM. INI"].iloc[0], type(df_may["HORIM. INI"].iloc[0]))
        print("HORIM. FIM first value:", df_may["HORIM. FIM"].iloc[0], type(df_may["HORIM. FIM"].iloc[0]))
        
        if "TEMPO EFETIVO PRODUÇÃO/PARADA" in df_may.columns:
            print("\nNon-null counts for 'TEMPO EFETIVO PRODUÇÃO/PARADA' in May:")
            print(df_may["TEMPO EFETIVO PRODUÇÃO/PARADA"].notna().sum())
            print("Unique values:")
            print(df_may["TEMPO EFETIVO PRODUÇÃO/PARADA"].value_counts().head(10))
            
    except Exception as e:
        print("Error:", e)
else:
    print("REV 1 file does not exist")
