import json
import os

config_path = "config.json"
try:
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    print("DB_FILE from JSON:", cfg.get("DB_FILE"))
    print("Exists?", os.path.exists(cfg.get("DB_FILE")))
except Exception as e:
    print("Error:", e)
