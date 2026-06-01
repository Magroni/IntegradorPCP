import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import data_manager as dm

# Let's see some pointing IDs that have paradas
p_df = dm.get_all_apontamentos()
print("All Pointings count:", len(p_df))

import pandas as pd
file_path = dm._get_apontamento_file()
paradas_df = pd.read_excel(file_path, sheet_name="PARADAS", engine="openpyxl")
print("Total paradas rows:", len(paradas_df))
if not paradas_df.empty:
    print("Sample paradas linked IDs:", paradas_df["ID_APONTAMENTO"].head(5).tolist())
    sample_id = paradas_df["ID_APONTAMENTO"].iloc[0]
    print(f"\nFetching paradas for sample ID {sample_id}...")
    res = dm.get_apontamento_paradas(sample_id)
    print("Found paradas count:", len(res))
    print(res)

insumos_df = pd.read_excel(file_path, sheet_name="INSUMOS", engine="openpyxl")
print("\nTotal insumos rows:", len(insumos_df))
if not insumos_df.empty:
    print("Sample insumos linked IDs:", insumos_df["ID_APONTAMENTO"].head(5).tolist())
    sample_id = insumos_df["ID_APONTAMENTO"].iloc[0]
    print(f"\nFetching insumos for sample ID {sample_id}...")
    res = dm.get_apontamento_insumos(sample_id)
    print("Found insumos count:", len(res))
    print(res)
