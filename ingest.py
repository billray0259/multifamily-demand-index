"""
CoStar spreadsheet ingestion module.

Reads one or more CoStar multifamily .xlsx files, standardizes column names,
coerces types, parses periods, and returns a single clean DataFrame.
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import List, Tuple, Union
import io

from config import NUMERIC_COERCE_COLUMNS


def _extract_market_name(filename: str) -> str:
    """
    Derive the market identifier from a filename.

    Handles patterns like:
        Austin_TX.xlsx          → Austin_TX
        Dallas_Fort_Worth_TX.xlsx → Dallas_Fort_Worth_TX
        Some Market (1).xlsx    → Some_Market
    """
    stem = Path(filename).stem
    # Remove trailing copy markers like " (1)"
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    # Replace spaces with underscores for consistency
    stem = stem.replace(" ", "_")
    return stem


def _parse_period(period_str: str) -> dict:
    """
    Parse a CoStar period string into components.

    Examples:
        '2025 Q4'       → {'Year': 2025, 'Quarter': 4, 'Is_QTD': False}
        '2026 Q1 QTD'   → {'Year': 2026, 'Quarter': 1, 'Is_QTD': True}
    """
    parts = str(period_str).strip().split()
    result = {"Year": None, "Quarter": None, "Is_QTD": False}
    try:
        result["Year"] = int(parts[0])
        result["Quarter"] = int(parts[1].replace("Q", ""))
        result["Is_QTD"] = len(parts) > 2 and "QTD" in parts[2].upper()
    except (IndexError, ValueError):
        pass
    return result


def read_costar_file(
    file_obj: Union[str, Path, io.BytesIO],
    filename: str = None,
) -> Tuple[pd.DataFrame, dict]:
    """
    Read a single CoStar .xlsx file and return a cleaned DataFrame.

    Parameters
    ----------
    file_obj : path-like or file-like
        The Excel file to read.  Accepts a filesystem path *or* an in-memory
        BytesIO object (from a Streamlit upload).
    filename : str, optional
        Original filename — used to derive the Market name when ``file_obj``
        is a BytesIO and has no path.

    Returns
    -------
    df : pd.DataFrame
        Cleaned data with a ``Market`` column added.
    report : dict
        Parse diagnostics (issues found, row count, etc.).
    """
    report: dict = {"issues": [], "rows": 0, "market": ""}

    # Derive market name
    if filename:
        market = _extract_market_name(filename)
    elif isinstance(file_obj, (str, Path)):
        market = _extract_market_name(str(file_obj))
    else:
        market = "Unknown_Market"
    report["market"] = market

    # Read the Excel file
    try:
        df = pd.read_excel(file_obj, engine="openpyxl")
    except Exception as exc:
        report["issues"].append(f"Failed to read file: {exc}")
        return pd.DataFrame(), report

    # ── Strip leading/trailing whitespace from column names ──────────────
    original_cols = df.columns.tolist()
    df.columns = df.columns.str.strip()
    if original_cols != df.columns.tolist():
        report["issues"].append("Stripped whitespace from column names")

    # ── Add Market column ────────────────────────────────────────────────
    df["Market"] = market

    # ── Coerce mixed-type columns to numeric ─────────────────────────────
    for col in NUMERIC_COERCE_COLUMNS:
        if col in df.columns and df[col].dtype == object:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            report["issues"].append(f"Coerced '{col}' to numeric")

    # ── Parse Period → Year / Quarter / Is_QTD ───────────────────────────
    if "Period" in df.columns:
        parsed = df["Period"].apply(_parse_period).apply(pd.Series)
        df["Year"] = parsed["Year"]
        df["Quarter"] = parsed["Quarter"]
        df["Is_QTD"] = parsed["Is_QTD"]
    else:
        report["issues"].append("No 'Period' column found")

    report["rows"] = len(df)
    return df, report


def ingest_files(
    file_objects: List[Tuple[Union[str, Path, io.BytesIO], str]],
) -> Tuple[pd.DataFrame, List[dict]]:
    """
    Ingest multiple CoStar .xlsx files and concatenate.

    Parameters
    ----------
    file_objects : list of (file_obj, filename) tuples
        Each element is either a (path, filename) or (BytesIO, filename).

    Returns
    -------
    combined : pd.DataFrame
        All markets combined, sorted by Market → Year → Quarter.
    reports : list of dict
        Per-file ingestion diagnostics.
    """
    frames: list = []
    reports: list = []

    for file_obj, filename in file_objects:
        df, report = read_costar_file(file_obj, filename)
        if not df.empty:
            frames.append(df)
        reports.append(report)

    if not frames:
        return pd.DataFrame(), reports

    combined = pd.concat(frames, ignore_index=True)

    # Sort consistently
    sort_cols = [c for c in ["Market", "Year", "Quarter"] if c in combined.columns]
    if sort_cols:
        combined = combined.sort_values(sort_cols).reset_index(drop=True)

    return combined, reports


def get_latest_quarter(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return one row per market representing the latest *complete* quarter
    (i.e. excluding QTD rows).
    """
    mask = df["Is_QTD"] == False  # noqa: E712
    complete = df[mask].copy()

    if complete.empty:
        return complete

    # Build a sortable period key
    complete["_period_key"] = complete["Year"] * 10 + complete["Quarter"]

    # Keep only the latest quarter per market
    idx = complete.groupby("Market")["_period_key"].idxmax()
    latest = complete.loc[idx].drop(columns=["_period_key"]).reset_index(drop=True)

    return latest


def compute_lagged_features(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute prior-year vacancy per market from the full historical dataset.

    The prior-year vacancy rate is the single largest suppressive predictor of
    rent growth in NMHC's panel regression (−24.5 to −27.7 bps per ppt; Bruen
    Dec 2024).  Using the same quarter one year earlier removes seasonal noise.

    Parameters
    ----------
    df : pd.DataFrame
        Full CoStar dataset (all periods, all markets).

    Returns
    -------
    pd.DataFrame
        One row per market with columns ``Market`` and ``Prior_Year_Vacancy``.
    """
    if "Vacancy Percent" not in df.columns:
        return pd.DataFrame(columns=["Market", "Prior_Year_Vacancy"])

    complete = df[df["Is_QTD"] == False].copy()  # noqa: E712

    records = []
    for market, grp in complete.groupby("Market"):
        grp = grp.sort_values(["Year", "Quarter"])
        latest = grp.iloc[-1]
        latest_year = int(latest["Year"])
        latest_quarter = int(latest["Quarter"])

        # Same quarter one year ago
        prior = grp[
            (grp["Year"] == latest_year - 1) &
            (grp["Quarter"] == latest_quarter)
        ]
        if len(prior):
            prior_vacancy = float(prior["Vacancy Percent"].values[0])
        else:
            # Fall back to current-quarter vacancy when prior year isn't available.
            # This is clearly less ideal than a true lag, but preserves the signal
            # direction (high vacancy = market slack) and avoids empty cells in export.
            current_vacancy = float(latest["Vacancy Percent"]) if pd.notna(latest["Vacancy Percent"]) else np.nan
            prior_vacancy = current_vacancy

        records.append({"Market": market, "Prior_Year_Vacancy": prior_vacancy})

    return pd.DataFrame(records)
