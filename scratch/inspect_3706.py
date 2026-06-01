import pandas as pd
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import data_manager as dm

print("Running dm.get_apontamentos_por_bloco('3706')...")
try:
    res = dm.get_apontamentos_por_bloco("3706")
    print(f"Success! Number of records returned: {len(res)}")
    print("\nRecords:")
    print(res.head(20))
except Exception as e:
    import traceback
    print("Error calling get_apontamentos_por_bloco:")
    traceback.print_exc()
