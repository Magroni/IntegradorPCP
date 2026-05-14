import data_manager as dm
import pandas as pd

def test_lookup(bloco_id):
    print(f"Testing lookup for block: {bloco_id}")
    info = dm.get_bloco_info(bloco_id)
    if info:
        print(f"Found: {info}")
    else:
        print("Not found.")

# Test known blocks from earlier analysis
test_lookup(470)
test_lookup("46")
test_lookup(213)
test_lookup(99999) # Non-existent
