"""
Chart rendering for Demand Index backtest validation.

Produces publication-quality matplotlib figures using the app's color palette
(from .streamlit/config.toml and export.py).
"""

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from scipy import stats as sp_stats


# ── App-consistent palette ───────────────────────────────────────────────────
PRIMARY     = "#1F4E79"
ACCENT_DARK = "#1F3864"
TEXT_COLOR  = "#1A1A2E"
NOTE_GRAY   = "#595959"
SECONDARY_BG = "#F0F4F8"

TIER_COLORS = {
    "High Demand":     "#C6EFCE",
    "Moderate Demand": "#FFEB9C",
    "Low Demand":      "#FFC7CE",
}
TIER_EDGE = {
    "High Demand":     "#2E7D32",
    "Moderate Demand": "#F57F17",
    "Low Demand":      "#C62828",
}
TIER_ORDER = ["Low Demand", "Moderate Demand", "High Demand"]


def scatter_chart(
    results_df: pd.DataFrame,
    stats: dict,
    horizon: str = "4Q",
) -> plt.Figure:
    """Scatter plot: Demand Index vs forward rent growth, with OLS line.

    Parameters
    ----------
    results_df : pd.DataFrame
        Output of validate.run_backtest — must contain Demand_Index, Tier,
        and the forward growth column.
    stats : dict
        Stats dict from run_backtest, keyed by horizon label ("1Q" or "4Q").
    horizon : str
        "1Q" or "4Q".

    Returns
    -------
    fig : matplotlib.figure.Figure
    """
    growth_col = f"Fwd_{horizon}_Growth"
    mask = results_df["Demand_Index"].notna() & results_df[growth_col].notna()
    df = results_df[mask].copy()

    fig, ax = plt.subplots(figsize=(8, 5.5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Plot points by tier
    for tier in TIER_ORDER:
        sub = df[df["Tier"] == tier]
        if sub.empty:
            continue
        ax.scatter(
            sub["Demand_Index"],
            sub[growth_col] * 100,  # convert to percentage
            c=TIER_COLORS[tier],
            edgecolors=TIER_EDGE[tier],
            linewidths=0.6,
            s=30,
            alpha=0.75,
            label=tier,
            zorder=3,
        )

    # OLS trend line
    x = df["Demand_Index"].values
    y = df[growth_col].values * 100
    slope, intercept, _, _, _ = sp_stats.linregress(x, y)
    x_line = np.linspace(0, 100, 200)
    y_line = slope * x_line + intercept
    ax.plot(x_line, y_line, color=PRIMARY, linewidth=2, zorder=4, label="OLS trend")

    # Stats annotation
    h_stats = stats.get(horizon, {})
    pearson_r = h_stats.get("pearson_r", 0)
    r_squared = h_stats.get("r_squared", 0)
    p_val = h_stats.get("pearson_p", 1)
    n = h_stats.get("n", 0)

    p_str = f"{p_val:.2e}" if p_val < 0.001 else f"{p_val:.4f}"
    sig_label = "***" if p_val < 0.001 else ("**" if p_val < 0.01 else ("*" if p_val < 0.05 else "n.s."))
    annotation = (
        f"Pearson r = {pearson_r:.3f} {sig_label}\n"
        f"R² = {r_squared:.3f}\n"
        f"p = {p_str}\n"
        f"n = {n:,}"
    )
    ax.text(
        0.03, 0.97, annotation,
        transform=ax.transAxes,
        fontsize=10,
        verticalalignment="top",
        fontfamily="sans-serif",
        bbox=dict(boxstyle="round,pad=0.4", facecolor=SECONDARY_BG, edgecolor=PRIMARY, alpha=0.9),
    )

    # Formatting
    horizon_label = "1-Quarter" if horizon == "1Q" else "4-Quarter (1-Year)"
    ax.set_xlabel("Demand Index (0–100)", fontsize=11, color=TEXT_COLOR)
    ax.set_ylabel(f"Forward {horizon_label} Rent Growth (%)", fontsize=11, color=TEXT_COLOR)
    ax.set_title(
        f"Demand Index vs. Forward {horizon_label} Rent Growth",
        fontsize=13, fontweight="bold", color=PRIMARY, pad=12,
    )
    ax.set_xlim(-2, 102)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.tick_params(colors=NOTE_GRAY, labelsize=9)
    ax.grid(axis="y", alpha=0.3, color="#CCCCCC")
    ax.legend(loc="lower right", fontsize=9, framealpha=0.9, edgecolor="#CCCCCC")

    fig.tight_layout()
    return fig


def tier_boxplot(
    results_df: pd.DataFrame,
    horizon: str = "4Q",
) -> plt.Figure:
    """Box-and-whisker of forward rent growth grouped by Demand tier.

    Parameters
    ----------
    results_df, horizon : same as scatter_chart.

    Returns
    -------
    fig : matplotlib.figure.Figure
    """
    growth_col = f"Fwd_{horizon}_Growth"
    mask = results_df["Demand_Index"].notna() & results_df[growth_col].notna()
    df = results_df[mask].copy()

    # Convert to percentage
    df["_growth_pct"] = df[growth_col] * 100

    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Prepare data in tier order
    data_by_tier = []
    labels = []
    colors = []
    edge_colors = []
    for tier in TIER_ORDER:
        sub = df[df["Tier"] == tier]["_growth_pct"]
        if sub.empty:
            continue
        data_by_tier.append(sub.values)
        n = len(sub)
        med = sub.median()
        labels.append(f"{tier}\n(n={n:,}, med={med:.2f}%)")
        colors.append(TIER_COLORS[tier])
        edge_colors.append(TIER_EDGE[tier])

    if not data_by_tier:
        ax.text(0.5, 0.5, "No data available", ha="center", va="center",
                transform=ax.transAxes, fontsize=14, color=NOTE_GRAY)
        return fig

    bp = ax.boxplot(
        data_by_tier,
        labels=labels,
        patch_artist=True,
        widths=0.5,
        showfliers=True,
        flierprops=dict(marker="o", markersize=3, alpha=0.4, color=NOTE_GRAY),
        medianprops=dict(color=PRIMARY, linewidth=2),
        whiskerprops=dict(color=NOTE_GRAY),
        capprops=dict(color=NOTE_GRAY),
    )

    for patch, fc, ec in zip(bp["boxes"], colors, edge_colors):
        patch.set_facecolor(fc)
        patch.set_edgecolor(ec)
        patch.set_linewidth(1.2)

    # Formatting
    horizon_label = "1-Quarter" if horizon == "1Q" else "4-Quarter (1-Year)"
    ax.set_ylabel(f"Forward {horizon_label} Rent Growth (%)", fontsize=11, color=TEXT_COLOR)
    ax.set_title(
        f"Rent Growth Distribution by Demand Tier ({horizon_label} Forward)",
        fontsize=13, fontweight="bold", color=PRIMARY, pad=12,
    )
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.tick_params(colors=NOTE_GRAY, labelsize=9)
    ax.grid(axis="y", alpha=0.3, color="#CCCCCC")

    # Zero line for reference
    ax.axhline(y=0, color=NOTE_GRAY, linestyle="--", linewidth=0.8, alpha=0.5)

    fig.tight_layout()
    return fig
