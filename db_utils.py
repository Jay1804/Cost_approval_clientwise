"""
Direct MySQL connection to the checkpoint_live database, replacing the old
Selenium/MIS-export-portal flow. Credentials come from .env.
"""

import os
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

load_dotenv()


def get_engine() -> Engine:
    host = os.getenv("DB_HOST")
    port = os.getenv("DB_PORT", "3306")
    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")
    name = os.getenv("DB_NAME")

    missing = [k for k, v in {
        "DB_HOST": host, "DB_USER": user, "DB_PASSWORD": password, "DB_NAME": name
    }.items() if not v]
    if missing:
        raise ValueError(f"Missing DB credentials in .env: {', '.join(missing)}")

    url = f"mysql+pymysql://{user}:{password}@{host}:{port}/{name}"
    return create_engine(url, pool_pre_ping=True)


def test_connection(engine: Engine):
    try:
        with engine.connect() as conn:
            conn.exec_driver_sql("SELECT 1")
        return True, "Connected successfully"
    except Exception as e:
        return False, str(e)


def fetch_client_list(engine: Engine) -> pd.DataFrame:
    from sql_queries import CLIENT_LIST_QUERY
    df = pd.read_sql(text(CLIENT_LIST_QUERY), engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def fetch_case_data(engine: Engine, client_ids=None) -> pd.DataFrame:
    from sql_queries import get_case_data_query
    df = pd.read_sql(text(get_case_data_query(client_ids)), engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def fetch_flexi_fields(engine: Engine, client_ids=None) -> pd.DataFrame:
    from sql_queries import get_flexi_field_query
    df = pd.read_sql(text(get_flexi_field_query(client_ids)), engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df
