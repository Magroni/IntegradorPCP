import pandas as pd
import json
from datetime import time

with open("config.json", "r", encoding="utf-8") as f:
    cfg = json.load(f)

ap_file = cfg["APONTAMENTO_FILE"]
sheet_name = cfg["SHEET_AP_PARADAS"]

try:
    df = pd.read_excel(ap_file, sheet_name=sheet_name, engine="openpyxl")
    print("Types in HORA_INICIO:")
    print(df["HORA_INICIO"].apply(lambda x: type(x).__name__).value_counts())
    print("\nTypes in DATA_INICIO:")
    print(df["DATA_INICIO"].apply(lambda x: type(x).__name__).value_counts())
    print("\nSample HORA_INICIO values and their types:")
    for v in df["HORA_INICIO"].dropna().unique()[:10]:
        print(f"Value: {v!r}, Type: {type(v).__name__}")
except Exception as e:
    print("Error:", e)
