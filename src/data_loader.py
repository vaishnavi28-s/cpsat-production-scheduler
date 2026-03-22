"""
Data loader — supports both Snowflake (production) and CSV (demo/local) modes.

In production: connects to Snowflake using environment variables.
In demo mode:  loads from data/sample_jobs.csv — no credentials needed.
"""

import os
import re
import numpy as np
import pandas as pd
from dotenv import load_dotenv

from .models import PrintJob

load_dotenv()

SEP = " | "
NA_VALS = ["", "NULL", "null", "None", "none"]


def _fmt(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s not in {"", "NaT", "nan"} else None


def _concat(series):
    seen, out = set(), []
    for v in series:
        fv = _fmt(v)
        if fv is not None and fv not in seen:
            seen.add(fv)
            out.append(fv)
    return SEP.join(out)


def _to_int(x):
    s = str(x or "").replace(",", "").strip()
    try:
        return int(float(s))
    except ValueError:
        return 0


def _merge_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Merge duplicate JOB rows by concatenating unique values."""
    df = df.copy()
    df["JOB"] = df["JOB"].astype(str)

    obj_cols = df.select_dtypes(include=["object", "string"]).columns
    for c in obj_cols:
        df[c] = df[c].replace(NA_VALS, np.nan)

    cols = [c for c in df.columns if c != "JOB"]
    merged = df.groupby("JOB", dropna=False).agg({c: _concat for c in cols}).reset_index()
    member_count = df.groupby("JOB", dropna=False).size().rename("member_count")
    merged = merged.merge(member_count, on="JOB")

    ordered = ["JOB", "member_count"] + [c for c in merged.columns if c not in {"JOB", "member_count"}]
    return merged[ordered].sort_values("JOB").reset_index(drop=True)


def _df_to_jobs(df: pd.DataFrame) -> list:
    """Convert a DataFrame to a list of PrintJob objects."""
    df = df.fillna("")

    qty_map = {}
    if "QUANTITYORDERED" in df.columns:
        df_qty = df[["JOB", "QUANTITYORDERED"]].copy()
        df_qty["QUANTITYORDERED"] = (
            df_qty["QUANTITYORDERED"].astype(str)
            .str.replace(",", "", regex=False)
            .str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)
        )
        df_qty = df_qty.dropna(subset=["QUANTITYORDERED"])
        df_qty["QUANTITYORDERED"] = df_qty["QUANTITYORDERED"].astype(float).round().astype(int)
        df_qty = df_qty.groupby("JOB", as_index=False)["QUANTITYORDERED"].first()
        qty_map = dict(zip(df_qty["JOB"].astype(str), df_qty["QUANTITYORDERED"]))

    pages_map = {}
    if "PAGES" in df.columns:
        df_pages = df[["JOB", "PAGES"]].copy()
        df_pages["PAGES"] = (
            df_pages["PAGES"].astype(str)
            .str.replace(",", "", regex=False)
            .str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)
        )
        df_pages = df_pages.dropna(subset=["PAGES"])
        df_pages["PAGES"] = df_pages["PAGES"].astype(float).round().astype(int)
        df_pages = df_pages.groupby("JOB", as_index=False)["PAGES"].first()
        pages_map = dict(zip(df_pages["JOB"].astype(str), df_pages["PAGES"]))

    merged = _merge_duplicates(df)

    return [
        PrintJob(
            JOB=str(row.get("JOB", "")),
            PRESS_LOCATION=str(row.get("PRESS_LOCATION", "")),
            SEND_TO_LOCATION=str(row.get("SEND_TO_LOCATION", "")),
            PRODUCTTYPE=str(row.get("PRODUCTTYPE", "")),
            PAPER=str(row.get("PAPER", "")),
            FINISHTYPE=str(row.get("FINISHTYPE", "")),
            FINISHINGOP=str(row.get("FINISHINGOP", "")),
            DELIVERYDATE=str(row.get("DELIVERYDATE", "")),
            INKSS1=str(row.get("INKSS1", "")),
            INKSS2=str(row.get("INKSS2", "")),
            QUANTITYORDERED=qty_map.get(str(row.get("JOB", "")), 0),
            PAGES=pages_map.get(str(row.get("JOB", "")), 0),
        )
        for _, row in merged.iterrows()
    ]


def load_from_csv(path: str) -> list:
    """Load print jobs from a CSV file (demo/local mode)."""
    df = pd.read_csv(path, dtype=str)
    print(f"Loaded {len(df)} rows from {path}")
    return _df_to_jobs(df)


def load_from_snowflake(limit: int = 5000) -> list:
    """Load print jobs from Snowflake (production mode)."""
    try:
        from snowflake.snowpark import Session
        from getpass import getpass
    except ImportError:
        raise ImportError("Install snowflake-snowpark-python to use Snowflake mode.")

    pwd = os.getenv("SF_PASSWORD") or getpass("Snowflake password: ")
    otp = os.getenv("SF_OTP") or input("Authenticator 6-digit code: ").strip()

    params = {
        "account": os.getenv("SF_ACCOUNT"),
        "user": os.getenv("SF_USER"),
        "password": f"{pwd}{otp}",
        "passcode_in_password": True,
        "role": os.getenv("SF_ROLE"),
        "warehouse": os.getenv("SF_WAREHOUSE"),
        "database": os.getenv("SF_DATABASE"),
        "schema": os.getenv("SF_SCHEMA"),
    }

    session = Session.builder.configs(params).create()
    print("Connected to Snowflake")

    session.sql("USE DATABASE BI_PROD").collect()
    session.sql("USE SCHEMA BPG_USA_CDW").collect()

    df = session.sql(f"""
        SELECT *
        FROM BI_PROD.BPG_USA_CDW.AI_COMBINED_RUN
        WHERE PRODUCTTYPE IN ('Jacket', 'Cover')
            AND SEND_TO_LOCATION IN (
                'Martinsburg - BVG', 'BERRYVILLE GRAPHICS',
                'BERRYVILLE', 'Offset Paperback MFG, Inc.', 'Fairfield - BVG'
            )
        ORDER BY JOB
        LIMIT {limit}
    """).to_pandas()

    print(f"Fetched {len(df)} rows from Snowflake")
    return _df_to_jobs(df)
