"""
Optional Census ACS data enhancement.

When a Census API key is provided, fetches population, income, and employment
data for each market's CBSA, computes growth rates, and merges onto the CoStar
data.  Falls back gracefully when no key is available.
"""

import pandas as pd
import numpy as np
import requests
import time
from typing import Optional, Dict, Tuple

from config import CBSA_SEARCH_TERMS, CENSUS_VARIABLES


# ── Public API ───────────────────────────────────────────────────────────────

def enhance_with_census(
    df: pd.DataFrame,
    api_key: str,
    progress_callback=None,
    cbsa_overrides: Dict[str, str] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, list]:
    """
    Fetch Census ACS data and merge onto the CoStar DataFrame.

    Parameters
    ----------
    df : pd.DataFrame
        CoStar data with ``Market`` and ``Year`` columns.
    api_key : str
        Census Bureau API key.
    progress_callback : callable, optional
        Called with (message: str) for UI progress updates.
    cbsa_overrides : dict, optional
        Manually supplied {market_name: cbsa_code} mappings that override
        or supplement the automatic CBSA search.

    Returns
    -------
    df_enhanced : pd.DataFrame
        Input DataFrame with Census columns appended (matched markets only
        receive Census data; unmatched rows are left as-is).
    census_df : pd.DataFrame
        Standalone Census data (one row per matched market) for the report.
    unmatched : list
        Market names for which no CBSA code could be found.
    """
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    markets = df["Market"].unique().tolist()
    log(f"Resolving CBSA codes for {len(markets)} markets...")

    # Step 1: Resolve CBSA codes (overrides applied inside)
    cbsa_map = _resolve_cbsa_codes(markets, api_key, log,
                                   cbsa_overrides=cbsa_overrides or {})
    if not cbsa_map:
        log("⚠️  Could not resolve any CBSA codes. Skipping Census enhancement.")
        return df, pd.DataFrame(), markets

    unmatched = [m for m in markets if m not in cbsa_map]
    log(f"Matched {len(cbsa_map)}/{len(markets)} markets to CBSA codes.")
    if unmatched:
        log(f"⚠️  No CBSA found for: {', '.join(unmatched)}")

    # Step 2: Determine which years we need
    years_needed = sorted(df["Year"].dropna().unique().astype(int))
    # We need the prior year too for growth-rate computation
    min_year = max(min(years_needed) - 1, 2013)
    fetch_years = list(range(min_year, max(years_needed) + 1))

    # Step 3: Fetch bulk Census data year by year
    log(f"Fetching Census ACS data for {len(fetch_years)} years...")
    all_records = []
    for year in fetch_years:
        log(f"  Fetching {year}...")
        year_data = _fetch_acs_data(year, api_key)
        if year_data is None:
            continue
        # Extract rows for our CBSAs
        for market, cbsa_code in cbsa_map.items():
            row = _extract_cbsa_row(year_data, cbsa_code)
            if row is not None:
                record = {"Market": market, "Year": year}
                variables = list(CENSUS_VARIABLES.keys())
                for i, var_code in enumerate(variables):
                    var_name = CENSUS_VARIABLES[var_code]
                    try:
                        val = float(row[i + 1]) if row[i + 1] not in [None, "", "-666666666", "-666666666.0"] else np.nan
                        record[var_name] = val
                    except (ValueError, IndexError):
                        record[var_name] = np.nan
                all_records.append(record)
        time.sleep(0.3)  # rate-limit courtesy

    if not all_records:
        log("⚠️  No Census data retrieved.")
        return df, pd.DataFrame(), unmatched

    census_raw = pd.DataFrame(all_records)
    log(f"Retrieved Census data: {len(census_raw)} market-year observations.")

    # Step 4: Compute derived metrics
    census_raw = _compute_derived(census_raw)

    # Step 5: Compute growth rates
    census_raw = _compute_growth_rates(census_raw)

    # Step 6: Build a latest-year snapshot for the report sheet
    latest_year = census_raw.groupby("Market")["Year"].max().reset_index()
    latest_year.columns = ["Market", "_max_year"]
    census_snapshot = census_raw.merge(latest_year, on="Market")
    census_snapshot = census_snapshot[census_snapshot["Year"] == census_snapshot["_max_year"]]
    census_snapshot = census_snapshot.drop(columns=["_max_year"])

    # Step 7: Merge onto CoStar data (join on Market + Year)
    census_cols_to_merge = [
        "Market", "Year",
        "Population", "Median_Household_Income", "Employment_Rate",
        "Employment_Growth",
        "In_Migration", "In_Migration_Rate",
        "Population_Growth", "Median_Household_Income_Growth",
    ]
    available_cols = [c for c in census_cols_to_merge if c in census_raw.columns]
    merge_df = census_raw[available_cols].copy()

    df_enhanced = df.merge(merge_df, on=["Market", "Year"], how="left")

    # Forward-fill Census data within each market (annual → quarterly)
    census_fill_cols = [c for c in merge_df.columns if c not in ("Market", "Year")]
    df_enhanced = df_enhanced.sort_values(["Market", "Year", "Quarter"])
    for col in census_fill_cols:
        df_enhanced[col] = df_enhanced.groupby("Market")[col].transform(
            lambda s: s.ffill().bfill()
        )

    log("✓ Census data merged successfully.")
    return df_enhanced, census_snapshot, unmatched


# ── Internal helpers ─────────────────────────────────────────────────────────

def _resolve_cbsa_codes(
    markets: list, api_key: str, log=None,
    cbsa_overrides: Dict[str, str] = None,
) -> Dict[str, str]:
    """Map market names to CBSA codes via the Census API.
    
    ``cbsa_overrides`` values are applied first; auto-search fills in the rest.
    """
    if log is None:
        log = lambda msg: None  # noqa: E731
    if cbsa_overrides is None:
        cbsa_overrides = {}

    # Fetch full CBSA list from Census
    url = (
        f"https://api.census.gov/data/2023/acs/acs5"
        f"?get=NAME"
        f"&for=metropolitan%20statistical%20area/micropolitan%20statistical%20area:*"
        f"&key={api_key}"
    )
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        log(f"⚠️  Failed to fetch CBSA list: {exc}")
        return {}

    # Build lookup: name → code
    cbsa_lookup = {row[0]: row[1] for row in data[1:]}

    # Match our markets (apply overrides first, auto-search for the rest)
    cbsa_map: Dict[str, str] = {}
    for market in markets:
        # Manual override wins unconditionally
        if market in cbsa_overrides:
            cbsa_map[market] = cbsa_overrides[market]
            continue

        search_term = CBSA_SEARCH_TERMS.get(market)
        if not search_term:
            search_term = market.replace("_", " ").rsplit(" ", 1)[0]

        for cbsa_name, cbsa_code in cbsa_lookup.items():
            if search_term.lower() in cbsa_name.lower():
                cbsa_map[market] = cbsa_code
                break

    return cbsa_map


def _fetch_acs_data(year: int, api_key: str) -> Optional[list]:
    """Fetch ACS 5-year data for ALL CBSAs for a given year."""
    var_list = ",".join(CENSUS_VARIABLES.keys())
    url = (
        f"https://api.census.gov/data/{year}/acs/acs5"
        f"?get=NAME,{var_list}"
        f"&for=metropolitan%20statistical%20area/micropolitan%20statistical%20area:*"
        f"&key={api_key}"
    )
    try:
        resp = requests.get(url, timeout=60)
        if resp.status_code == 200:
            return resp.json()
    except Exception:
        pass
    return None


def _extract_cbsa_row(data: list, cbsa_code: str) -> Optional[list]:
    """Find a specific CBSA's row in the bulk Census response."""
    for row in data[1:]:  # skip header
        if row[-1] == cbsa_code:
            return row
    return None


def _compute_derived(df: pd.DataFrame) -> pd.DataFrame:
    """Compute derived Census metrics."""
    df = df.copy()

    # Employment rate
    df["Employment_Rate"] = np.where(
        df["Labor_Force"] > 0,
        df["Employed"] / df["Labor_Force"],
        np.nan,
    )

    # In-migration rate (as % of population)
    df["In_Migration_Rate"] = np.where(
        df["Population"] > 0,
        df["In_Migration"] / df["Population"],
        np.nan,
    )

    return df


def _compute_growth_rates(df: pd.DataFrame) -> pd.DataFrame:
    """Compute YoY growth rates for key Census variables.

    Employment_Growth is the year-over-year change in Employment_Rate.
    NMHC Dec 2024 found employment *growth* (+19.8 bps) to be the largest
    single demand predictor — the level (employment rate) is not significant;
    it is the change that matters.
    """
    df = df.sort_values(["Market", "Year"]).copy()

    for var in ["Population", "Median_Household_Income"]:
        if var in df.columns:
            df[f"{var}_Growth"] = df.groupby("Market")[var].pct_change()

    # Employment growth: YoY change in the employment rate (not its level)
    if "Employment_Rate" in df.columns:
        df["Employment_Growth"] = df.groupby("Market")["Employment_Rate"].pct_change()

    return df
