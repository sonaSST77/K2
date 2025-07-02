import oracledb
import os

def get_db_connection(config_file=os.path.join("parametry", "db_config.txt")):
    params = {}
    with open(config_file, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, value = line.strip().split("=", 1)
                params[key.strip()] = value.strip()
    return oracledb.connect(
        user=params["username"],
        password=params["password"],
        dsn=params["dsn"]
    )
