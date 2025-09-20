import pyodbc
import configparser
import os

# Load configuration
config = configparser.ConfigParser()
config.read(os.path.join(os.path.dirname(__file__), "config.ini"))

def get_connection() -> pyodbc.Connection:
    """
    Create and return a connection to the MSSQL server.
    Reads settings from config.ini.
    """
    db = config["database"]

    conn_str = (
        f"DRIVER={{{db['DRIVER']}}};"
        f"SERVER={db['SERVER']};"
        f"DATABASE={db['DATABASE']};"
        f"UID={db['UID']};"
        f"PWD={db['PWD']};"
        f"TrustServerCertificate={db.get('TrustServerCertificate', 'yes')};"
    )
    return pyodbc.connect(conn_str)

def get_base_folder() -> str:
    """
    Return the base folder path from config.ini.
    """
    return config["paths"]["base_folder"]