import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import data_manager as dm
import pandas as pd

df_base = dm.get_base_dados()
if not df_base.empty:
    if "SETOR" in df_base.columns:
        print("\nSetores Únicos:")
        print(df_base["SETOR"].unique().tolist())
    
    if "PROCESSO" in df_base.columns:
        print("\nProcessos de Resinagem (Baseados no nome):")
        resin_procs = [p for p in df_base["PROCESSO"].unique() if any(k in str(p).upper() for k in ["RESINA", "TELA", "MANTA", "ESTUQUE"])]
        print(resin_procs)
else:
    print("Base de dados vazia.")
