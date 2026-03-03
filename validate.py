"""
Demand Index backtest / validation engine.

For each historical quarter in the uploaded data, computes a cross-sectional
Demand Index across all markets and correlates it with *forward* rent growth
derived from effective rent levels — not CoStar's YoY column, which would
introduce information overlap (75% of the YoY value at t+1 has already been
realised at the time the index is computed at t).

Rent growth is intentionally excluded from the index inputs, making this a
genuine out-of-sample test.
"""

import pandas as pd
import numpy as np
from scipy import stats as sp_stats
from typing import Tuple

from config import (
    INDEX_WEIGHTS_FULL,
    INDEX_WEIGHTS_COSTAR_ONLY,
    TIER_THRESHOLDS,
)


# ── Internal helpers ─────────────────────────────────────────────────────────

def _extract_components(row: pd.Series, use_census: bool) -> dict:
    """Map a single market-quarter row into named index components.

    Mirrors demand_index._extract_components exactly so the backtest
    uses the same logic the production tool uses.
    """
    components = {}

    inv = row.get("Inventory Units", np.nan)
    abs_units = row.get("Absorption Units", np.nan)
    if pd.notna(inv) and inv > 0 and pd.notna(abs_units):
        components["absorption_pct"] = abs_units / inv
    else:
        components["absorption_pct"] = row.get("Absorption Percent", np.nan)

    components["prior_vacancy"] = row.get("Prior_Year_Vacancy", np.nan)
    components["deliveries_pct"] = row.get("Deliveries Percent", np.nan)
    components["occupancy"] = row.get("Occupancy Percent", np.nan)

    if use_census:
        components["population_growth"] = row.get("Population_Growth", np.nan)
        components["income_growth"] = row.get("Median_Household_Income_Growth", np.nan)
        components["employment_growth"] = row.get("Employment_Growth", np.nan)

    return components


def _compute_index_for_quarter(
    quarter_df: pd.DataFrame,
    weights: dict,
    use_census: bool,
) -> pd.DataFrame:
    """Compute the Demand Index for one cross-section (one quarter, many markets).

    Returns a DataFrame with Market, Demand_Index, Tier columns.
    """
    records = []
    for _, row in quarter_df.iterrows():
        rec = {"Market": row["Market"]}
        rec.update(_extract_components(row, use_census))
        records.append(rec)

    comp_df = pd.DataFrame(records)
    comp_keys = list(weights.keys())

    # Z-score normalisation
    z_df = comp_df[["Market"]].copy()
    for key in comp_keys:
        col = comp_df[key]
        mean = col.mean()
        std = col.std()
        if std == 0 or pd.isna(std):
            z_df[f"{key}_z"] = 0.0
        else:
            z_df[f"{key}_z"] = (col - mean) / std

    # Weighted sum
    z_df["raw_score"] = 0.0
    for key, weight in weights.items():
        z_df["raw_score"] += z_df[f"{key}_z"].fillna(0) * weight

    # Rescale 0–100
    raw_min = z_df["raw_score"].min()
    raw_max = z_df["raw_score"].max()
    if raw_max == raw_min:
        z_df["Demand_Index"] = 50.0
    else:
        z_df["Demand_Index"] = (z_df["raw_score"] - raw_min) / (raw_max - raw_min) * 100

    # Tier
    def _tier(s):
        if s >= TIER_THRESHOLDS["High Demand"]:
            return "High Demand"
        elif s >= TIER_THRESHOLDS["Moderate Demand"]:
            return "Moderate Demand"
        return "Low Demand"

    z_df["Tier"] = z_df["Demand_Index"].apply(_tier)
    return z_df[["Market", "Demand_Index", "Tier"]]


def _compute_prior_year_vacancy(df: pd.DataFrame) -> pd.DataFrame:
    """Add Prior_Year_Vacancy to every row (same quarter, year − 1)."""
    df = df.copy()
    df["Prior_Year_Vacancy"] = np.nan

    for market, grp in df.groupby("Market"):
        for idx, row in grp.iterrows():
            prior = grp[
                (grp["Year"] == row["Year"] - 1) &
                (grp["Quarter"] == row["Quarter"])
            ]
            if len(prior):
                df.loc[idx, "Prior_Year_Vacancy"] = float(
                    prior["Vacancy Percent"].values[0]
                )
            else:
                # Fall back to current vacancy (same as production code)
                df.loc[idx, "Prior_Year_Vacancy"] = row.get("Vacancy Percent", np.nan)

    return df


# ── Public API ───────────────────────────────────────────────────────────────

def run_backtest(
    df_all: pd.DataFrame,
    use_census: bool,
    min_markets: int = 5,
) -> Tuple[pd.DataFrame, dict]:
    """Run the historical backtest.

    Parameters
    ----------
    df_all : pd.DataFrame
        Full multi-quarter, multi-market DataFrame from ingest (all rows,
        including Census columns if available).
    use_census : bool
        Whether Census growth columns are present and should be used.
    min_markets : int
        Minimum number of markets required in a quarter to compute the index.

    Returns
    -------
    results : pd.DataFrame
        One row per market-quarter with columns:
        Market, Year, Quarter, Demand_Index, Tier, Fwd_1Q_Growth, Fwd_4Q_Growth,
        Effective_Rent
    stats : dict
        Correlation stats for each horizon (pearson_r, spearman_rho, p-values,
        r_squared, slope, intercept, n, per-tier medians/counts).
    """
    weights = INDEX_WEIGHTS_FULL if use_census else INDEX_WEIGHTS_COSTAR_ONLY

    # ── Prepare data ─────────────────────────────────────────────────────
    df = df_all.copy()
    df = df[df["Is_QTD"] == False].reset_index(drop=True)  # noqa: E712
    df = df.sort_values(["Market", "Year", "Quarter"]).reset_index(drop=True)

    # Prior-year vacancy for every row
    df = _compute_prior_year_vacancy(df)

    # ── Compute forward rent growth from rent levels ─────────────────────
    # Forward 1Q:  rent_{t+1} / rent_t  − 1
    # Forward 4Q:  rent_{t+4} / rent_t  − 1
    rent_col = "Effective Rent Per Unit"
    if rent_col not in df.columns:
        raise ValueError(f"Column '{rent_col}' not found in data — cannot compute forward rent growth.")

    df["_period_key"] = df["Year"] * 10 + df["Quarter"]

    for market, grp in df.groupby("Market"):
        grp_sorted = grp.sort_values("_period_key")
        rent = grp_sorted[rent_col].values
        idxs = grp_sorted.index.values
        n = len(rent)
        for i in range(n):
            # 1Q forward
            if i + 1 < n:
                if pd.notna(rent[i]) and rent[i] > 0 and pd.notna(rent[i + 1]):
                    df.loc[idxs[i], "Fwd_1Q_Growth"] = rent[i + 1] / rent[i] - 1
            # 4Q forward
            if i + 4 < n:
                if pd.notna(rent[i]) and rent[i] > 0 and pd.notna(rent[i + 4]):
                    df.loc[idxs[i], "Fwd_4Q_Growth"] = rent[i + 4] / rent[i] - 1

    # ── Compute index per quarter ────────────────────────────────────────
    quarter_groups = df.groupby(["Year", "Quarter"])
    all_results = []

    for (year, qtr), qdf in quarter_groups:
        if len(qdf) < min_markets:
            continue

        # If using census, skip quarters where census columns are all NaN
        if use_census:
            census_cols = ["Population_Growth", "Median_Household_Income_Growth", "Employment_Growth"]
            present = [c for c in census_cols if c in qdf.columns]
            if not present or qdf[present].notna().sum().sum() == 0:
                continue

        idx_df = _compute_index_for_quarter(qdf, weights, use_census)
        idx_df["Year"] = year
        idx_df["Quarter"] = qtr

        # Merge forward growth targets back
        merge_cols = ["Market", "Fwd_1Q_Growth", "Fwd_4Q_Growth", rent_col]
        merge_cols = [c for c in merge_cols if c in qdf.columns]
        idx_df = idx_df.merge(
            qdf[merge_cols].rename(columns={rent_col: "Effective_Rent"}),
            on="Market",
            how="left",
        )

        all_results.append(idx_df)

    if not all_results:
        return pd.DataFrame(), {"error": "No quarters with enough data to compute index."}

    results = pd.concat(all_results, ignore_index=True)

    # ── Compute statistics ───────────────────────────────────────────────
    stats = {}
    for horizon in ["Fwd_1Q_Growth", "Fwd_4Q_Growth"]:
        label = "1Q" if "1Q" in horizon else "4Q"
        mask = results["Demand_Index"].notna() & results[horizon].notna()
        sub = results[mask]

        if len(sub) < 10:
            stats[label] = {"n": len(sub), "error": "Insufficient observations"}
            continue

        x = sub["Demand_Index"].values
        y = sub[horizon].values

        # Pearson
        pearson_r, pearson_p = sp_stats.pearsonr(x, y)
        # Spearman
        spearman_rho, spearman_p = sp_stats.spearmanr(x, y)
        # OLS
        slope, intercept, r_val, p_val, std_err = sp_stats.linregress(x, y)
        r_squared = r_val ** 2

        # Per-tier stats
        tier_stats = {}
        for tier in ["High Demand", "Moderate Demand", "Low Demand"]:
            t_sub = sub[sub["Tier"] == tier]
            tier_stats[tier] = {
                "n": len(t_sub),
                "median": float(t_sub[horizon].median()) if len(t_sub) else None,
                "mean": float(t_sub[horizon].mean()) if len(t_sub) else None,
                "q25": float(t_sub[horizon].quantile(0.25)) if len(t_sub) else None,
                "q75": float(t_sub[horizon].quantile(0.75)) if len(t_sub) else None,
            }

        stats[label] = {
            "n": len(sub),
            "pearson_r": pearson_r,
            "pearson_p": pearson_p,
            "spearman_rho": spearman_rho,
            "spearman_p": spearman_p,
            "r_squared": r_squared,
            "slope": slope,
            "intercept": intercept,
            "std_err": std_err,
            "tiers": tier_stats,
        }

    stats["model"] = "Full (Census + CoStar)" if use_census else "CoStar Only"
    stats["total_quarters"] = results[["Year", "Quarter"]].drop_duplicates().shape[0]
    stats["total_markets"] = results["Market"].nunique()

    return results, stats
