from __future__ import annotations

import json
from os import getenv
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
from dotenv import load_dotenv


# Loading the credentials to get access to the DB
# from the .env file
load_dotenv()
USERNAME = getenv("USERNAME")
PASSWORD = getenv("PASSWORD")
HOST = getenv("HOST")
PORT = getenv("PORT")
DATABASE = getenv("DATABASE")

# Loading the maps for departments and columns replacements, Excel styles
# config and request line from the JSON and TXT files
with open("./config/department_map.json", "r", encoding="utf-8") as deps, \
     open("./config/column_names.json", "r", encoding="utf-8") as col_names, \
     open("./config/request_to_db.txt", "r", encoding="utf-8") as req, \
     open("./config/styles_config.json", "r", encoding="utf-8") as style:
    department_map = json.load(deps)
    columns_rename = json.load(col_names)
    request = req.read()
    styles = json.load(style)


# Make an Engine with successful connection to the DB
def connect_to_database() -> Engine | None:
    try:
        return create_engine(
            f"mysql+pymysql://{USERNAME}:{PASSWORD}@{HOST}:{PORT}/{DATABASE}"
        )
    except ValueError:
        raise(
            "Return value must be of 'Engine' type. "
            "Connection to the database not established."
        )
