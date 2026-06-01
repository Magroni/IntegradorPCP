import inspect
import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import data_manager as dm

print("Functions in data_manager:")
for name, obj in inspect.getmembers(dm, inspect.isfunction):
    print(f"- {name}{inspect.signature(obj)}")
