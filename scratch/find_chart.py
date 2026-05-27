import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_FILE = os.path.join(BASE_DIR, "app.py")

if os.path.exists(APP_FILE):
    with open(APP_FILE, "r", encoding="utf-8") as f:
        lines = f.readlines()
        
    print("Found chart patterns:")
    for idx, line in enumerate(lines):
        line_num = idx + 1
        if any(kw in line for kw in ["alt.Chart", "altair_chart", "px.bar", "bar_chart", "chart"]):
            print(f"Line {line_num}: {line.strip()}")
else:
    print("app.py does not exist")
