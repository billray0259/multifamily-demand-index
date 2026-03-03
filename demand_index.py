"""
Core Demand Index computation engine.

Computes a composite 0–100 score per market using z-score normalization
and research-backed component weights.  See config.METHODOLOGY_TEXT for
the full academic grounding.
"""

import pandas as pd
import numpy as np
from typing import Tuple

from config import (
    INDEX_WEIGHTS_FULL,
    INDEX_WEIGHTS_COSTAR_ONLY,
    TIER_THRESHOLDS,
)

MODEL_WEIGHTED_Z    = "weighted_z"
MODEL_ABS_SUPPLY    = "abs_supply"


def _has_census_data(df: pd.DataFrame) -> bool:
    """Check whether Census columns are present and populated."""
    census_cols = ["Population_Growth", "Median_Household_Income_Growth", "Employment_Growth"]
    for col in census_cols:
        if col in df.columns and df[col].notna().any():
            return True
    return False


def _extract_components(row: pd.Series, use_census: bool) -> dict:
    """
    Map a single market's row into the named index components.

    Returns a dict whose keys match the weight-dict keys.
    """
    components = {}

    # ── CoStar demand fundamentals ───────────────────────────────────────
    inv = row.get("Inventory Units", np.nan)
    abs_units = row.get("Absorption Units", np.nan)

    # Absorption as % of inventory
    if pd.notna(inv) and inv > 0 and pd.notna(abs_units):
        components["absorption_pct"] = abs_units / inv
    else:
        components["absorption_pct"] = row.get("Absorption Percent", np.nan)

    # ── Supply pressure & lagged vacancy ────────────────────────────────
    # prior_vacancy is always extracted (CoStar-derived lag, computed in ingest).
    # Higher prior-year vacancy signals existing slack → negative weight in index.
    # Rent growth is NOT included: it is the dependent variable in both NMHC
    # regressions and would create circularity if used as an input here.
    components["prior_vacancy"] = row.get("Prior_Year_Vacancy", np.nan)
    components["deliveries_pct"] = row.get("Deliveries Percent", np.nan)

    # occupancy is only used in the CoStar-only model weight dict
    components["occupancy"] = row.get("Occupancy Percent", np.nan)

    # ── Census demand drivers (optional) ──────────────────────────────────
    if use_census:
        components["population_growth"] = row.get("Population_Growth", np.nan)
        components["income_growth"] = row.get("Median_Household_Income_Growth", np.nan)
        # Use employment *growth* (YoY change in rate), not the rate level.
        # NMHC Dec 2024: employment growth +19.8 bps vs rate level which is not significant.
        components["employment_growth"] = row.get("Employment_Growth", np.nan)

    return components


def compute_demand_index(df_latest: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compute the Demand Index for every market in ``df_latest``.

    Parameters
    ----------
    df_latest : pd.DataFrame
        One row per market (the latest complete quarter), with CoStar columns
        and optionally Census columns already merged.

    Returns
    -------
    rankings : pd.DataFrame
        Markets ranked by Demand Index (descending).  Columns include
        ``Demand_Index``, ``Tier``, and all input metrics.
    components_df : pd.DataFrame
        Per-market breakdown: raw values, z-scores, and weighted contributions
        for each index component.
    """
    use_census = _has_census_data(df_latest)
    weights = INDEX_WEIGHTS_FULL if use_census else INDEX_WEIGHTS_COSTAR_ONLY

    # ── Build component matrix ───────────────────────────────────────────
    records = []
    for _, row in df_latest.iterrows():
        rec = {"Market": row["Market"]}
        rec.update(_extract_components(row, use_census))
        records.append(rec)

    comp_df = pd.DataFrame(records)
    comp_keys = [k for k in weights.keys()]

    # ── Z-score normalisation ────────────────────────────────────────────
    z_df = comp_df[["Market"]].copy()
    for key in comp_keys:
        col = comp_df[key]
        mean = col.mean()
        std = col.std()
        if std == 0 or pd.isna(std):
            z_df[f"{key}_z"] = 0.0
        else:
            z_df[f"{key}_z"] = (col - mean) / std
        # Store raw values for the breakdown sheet
        z_df[f"{key}_raw"] = col

    # ── Weighted sum ─────────────────────────────────────────────────────
    z_df["raw_score"] = 0.0
    for key, weight in weights.items():
        z_col = f"{key}_z"
        contribution = z_df[z_col].fillna(0) * weight
        z_df[f"{key}_contribution"] = contribution
        z_df["raw_score"] += contribution

    # ── Rescale to 0–100 ─────────────────────────────────────────────────
    raw_min = z_df["raw_score"].min()
    raw_max = z_df["raw_score"].max()
    if raw_max == raw_min:
        z_df["Demand_Index"] = 50.0
    else:
        z_df["Demand_Index"] = (
            (z_df["raw_score"] - raw_min) / (raw_max - raw_min) * 100
        )

    # ── Tier classification ──────────────────────────────────────────────
    def _tier(score):
        if score >= TIER_THRESHOLDS["High Demand"]:
            return "High Demand"
        elif score >= TIER_THRESHOLDS["Moderate Demand"]:
            return "Moderate Demand"
        else:
            return "Low Demand"

    z_df["Tier"] = z_df["Demand_Index"].apply(_tier)

    # ── Assemble rankings table ──────────────────────────────────────────
    rankings = df_latest.copy()
    rankings = rankings.merge(
        z_df[["Market", "Demand_Index", "Tier"]], on="Market", how="left"
    )
    rankings = rankings.sort_values("Demand_Index", ascending=False).reset_index(drop=True)
    rankings.insert(0, "Rank", range(1, len(rankings) + 1))

    # ── Assemble components breakdown ────────────────────────────────────
    components_out = z_df.sort_values("Demand_Index", ascending=False).reset_index(drop=True)

    return rankings, components_out


def compute_absorption_supply_index(
    df_latest: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Alternative model: score = absorption_units / (inventory * vacancy_pct/100 + uc_units)

    Conceptually this is absorption relative to the total "contested" unit pool:
    the units already vacant today plus the under-construction pipeline that will
    hit the market soon.  A high ratio means demand is outrunning available and
    incoming supply — analogous to a low months-of-supply figure.

    Entirely CoStar-derived; no Census data required.
    """
    records = []
    for _, row in df_latest.iterrows():
        inv       = row.get("Inventory Units", np.nan)
        abs_units = row.get("Absorption Units", np.nan)
        vac_pct   = row.get("Vacancy Percent", np.nan)   # e.g. 7.2 (not 0.072)
        uc_units  = row.get("Under Construction Units", np.nan)

        vacant_units = (inv * vac_pct / 100.0
                        if pd.notna(inv) and pd.notna(vac_pct) else np.nan)
        uc = float(uc_units) if pd.notna(uc_units) else 0.0
        vac = float(vacant_units) if pd.notna(vacant_units) else 0.0
        denom = vac + uc

        if denom > 0 and pd.notna(abs_units):
            raw_score = abs_units / denom
        else:
            raw_score = np.nan

        records.append({
            "Market":            row["Market"],
            "absorption_units_raw": abs_units,
            "vacant_units_raw":     vacant_units,
            "uc_units_raw":         uc_units,
            "denominator_raw":      denom if denom > 0 else np.nan,
            "raw_score":            raw_score,
        })

    comp_df = pd.DataFrame(records)

    # ── Rescale to 0–100 ────────────────────────────────────────────────
    valid = comp_df["raw_score"].dropna()
    raw_min, raw_max = (valid.min(), valid.max()) if len(valid) > 1 else (0.0, 1.0)

    if raw_max == raw_min:
        comp_df["Demand_Index"] = 50.0
    else:
        comp_df["Demand_Index"] = (
            (comp_df["raw_score"] - raw_min) / (raw_max - raw_min) * 100
        ).fillna(50.0)

    # ── Tier classification ──────────────────────────────────────────────
    def _tier(score):
        if score >= TIER_THRESHOLDS["High Demand"]:
            return "High Demand"
        elif score >= TIER_THRESHOLDS["Moderate Demand"]:
            return "Moderate Demand"
        else:
            return "Low Demand"

    comp_df["Tier"] = comp_df["Demand_Index"].apply(_tier)

    # ── Assemble rankings table ──────────────────────────────────────────
    rankings = df_latest.copy()
    rankings = rankings.merge(
        comp_df[["Market", "Demand_Index", "Tier"]], on="Market", how="left"
    )
    rankings = rankings.sort_values("Demand_Index", ascending=False).reset_index(drop=True)
    rankings.insert(0, "Rank", range(1, len(rankings) + 1))

    components_out = comp_df.sort_values("Demand_Index", ascending=False).reset_index(drop=True)
    return rankings, components_out
