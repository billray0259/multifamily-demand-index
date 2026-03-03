"""
Multifamily Demand Index — Streamlit Application

Upload CoStar multifamily .xlsx spreadsheets → compute a research-backed
Demand Index → download a formatted Excel workbook with market rankings.
"""

import streamlit as st
import pandas as pd
import io
from pathlib import Path
from streamlit_js_eval import streamlit_js_eval

from ingest import ingest_files, get_latest_quarter, compute_lagged_features
from census_enhance import enhance_with_census
from demand_index import (
    compute_demand_index,
    compute_absorption_supply_index,
    MODEL_WEIGHTED_Z,
    MODEL_ABS_SUPPLY,
)
from export import generate_workbook
from validate import run_backtest
from validate_charts import scatter_chart, tier_boxplot
from config import (
    APP_TITLE,
    APP_VERSION,
    APP_DESCRIPTION,
    METHODOLOGY_TEXT,
    INDEX_WEIGHTS_FULL,
    INDEX_WEIGHTS_COSTAR_ONLY,
)


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)


def main():
    # ── Header ───────────────────────────────────────────────────────────
    st.title(f"🏢 {APP_TITLE}")
    st.markdown(APP_DESCRIPTION)

    st.divider()

    # ── Sidebar — Census API key ─────────────────────────────────────────
    with st.sidebar:
        st.header("Settings")

        # Read saved key from browser localStorage on first load
        _LS_KEY = "mf_demand_index_census_key"
        if "census_key_loaded" not in st.session_state:
            saved_key = streamlit_js_eval(
                js_expressions=f"localStorage.getItem('{_LS_KEY}')",
                key="_read_census_key",
            )
            st.session_state["census_key_init"] = saved_key or ""
            st.session_state["census_key_loaded"] = True

        census_key = st.text_input(
            "Census API Key (optional)",
            value=st.session_state.get("census_key_init", ""),
            type="password",
            help=(
                "Providing a free Census API key enriches the index with "
                "population growth, income growth, and employment data. "
                "[Get a free key →](https://api.census.gov/data/key_signup.html)"
            ),
            key="census_key_input",
        )

        # Persist any change back to localStorage immediately
        if census_key:
            streamlit_js_eval(
                js_expressions=f"localStorage.setItem('{_LS_KEY}', {repr(census_key)})",
                key="_save_census_key",
            )
        else:
            streamlit_js_eval(
                js_expressions=f"localStorage.removeItem('{_LS_KEY}')",
                key="_clear_census_key",
            )

        st.caption(
            "Without a Census key, the index uses CoStar data only "
            "(absorption, occupancy, rent growth, construction, deliveries)."
        )

        st.divider()
        st.header("About")
        st.markdown(
            f"**Version** {APP_VERSION}  \n"
            "Built on research from NMHC, Harvard JCHS, "
            "and real estate economics literature."
        )

        with st.expander("View full methodology"):
            st.markdown(METHODOLOGY_TEXT)

    # ── Step 0: How to export from CoStar ───────────────────────────────
    with st.expander("0 · How to export data from CoStar", expanded=False):
        st.markdown(
            """
**You will need one `.xlsx` file per market, downloaded from CoStar's Market Analytics module.**

---

#### Export steps (CoStar Market Analytics)

1. **Log in** at [costar.com](https://www.costar.com) and click **Research** in the top navigation bar.
2. Click **Market Analytics** from the Research dropdown.
3. In the **Property Type** filter (left panel), select **Multifamily** (or **Apartment**, depending on your subscription tier).
4. Use the **Geography** search box to select the target metro market (e.g. *Austin, TX*).
5. Click the **Statistics** tab to confirm the quarterly time-series table is visible.  
   You should see columns including *Period*, *Vacancy %*, *Absorption Units*, *Effective Rent % Growth/Yr*, *Under Construction %*, and *Deliveries %*.
6. Click the **Export** button ( ⬇ icon, top-right of the table) → choose **Excel (.xlsx)**.
7. CoStar will download a file — **rename it** using the convention below before uploading here.

> **Note:** CoStar updates its UI periodically. If the navigation looks different, look for  
> *Research → Market Analytics → Multifamily → [Select Market] → Statistics → Export*.

---

#### File naming convention

The market name shown in rankings is derived **directly from the filename** — rename each file before uploading:

| Market | Correct filename |
|---|---|
| Austin, TX | `Austin_TX.xlsx` |
| Dallas / Fort Worth, TX | `Dallas_Fort_Worth_TX.xlsx` |
| Northwest Arkansas | `Northwest_Arkansas_AK.xlsx` |
| College Station–Bryan, TX | `College_Station_Bryan_TX.xlsx` |

**Rules:**
- Replace spaces and punctuation with underscores `_`
- Append the two-letter state abbreviation: `CityName_ST.xlsx`
- Multi-word cities: `Kansas_City_MO.xlsx`, `Salt_Lake_City_UT.xlsx`
- Multi-city metros: `Dallas_Fort_Worth_TX.xlsx`, `Saint_Louis_MO.xlsx`
- Do **not** include dates or version numbers (e.g. avoid `Austin_TX (1).xlsx`)

The app will clean up trailing copy markers like ` (1)` automatically, but clean names are clearest.
"""
        )

    # ── File upload ──────────────────────────────────────────────────────
    st.subheader("1 · Upload CoStar Spreadsheets")
    st.caption(
        "Upload one or more `.xlsx` files exported from CoStar.  "
        "Each file should represent one market (e.g. `Austin_TX.xlsx`).  "
        "The market name is derived from the filename."
    )

    uploaded_files = st.file_uploader(
        "Drag and drop CoStar .xlsx files",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not uploaded_files:
        st.info("👆 Upload one or more CoStar .xlsx files to get started.")
        st.stop()

    # ── Process button ───────────────────────────────────────────────────
    st.subheader("2 · Process & Compute Index")

    model_type = st.radio(
        "Scoring model",
        options=[
            MODEL_WEIGHTED_Z,
            MODEL_ABS_SUPPLY,
        ],
        format_func=lambda m: {
            MODEL_WEIGHTED_Z: "📊 Weighted Z-Score  (research-backed composite; uses Census if available)",
            MODEL_ABS_SUPPLY: "⚗️ Absorption / Supply Pressure  (absorption ÷ vacant + pipeline; experimental)",
        }[m],
        horizontal=False,
        help=(
            "**Weighted Z-Score** normalises each component to a z-score, applies "
            "research-derived weights, and rescales to 0–100.\n\n"
            "**Absorption / Supply Pressure** computes "
            "`absorption_units / (inventory × vacancy% + under_construction_units)` — "
            "a ratio inversely proportional to months-of-supply at current absorption. "
            "CoStar-only; no Census data used."
        ),
    )

    col1, col2 = st.columns([1, 3])
    with col1:
        process_btn = st.button("🚀 Compute Demand Index", type="primary", use_container_width=True)

    # Auto-trigger when returning from a CBSA re-run (overrides stored, no results yet)
    cbsa_rerun = bool(st.session_state.get("cbsa_overrides")) and "rankings" not in st.session_state

    if not process_btn and not cbsa_rerun and "rankings" not in st.session_state:
        st.info(f"📁 {len(uploaded_files)} file(s) selected. Click **Compute Demand Index** to proceed.")
        st.stop()

    def _run_pipeline(uploaded_files, census_key, cbsa_overrides, model_type=MODEL_WEIGHTED_Z):
        """Run the full pipeline; returns a dict of results or raises."""
        with st.status("Processing…", expanded=True) as status:
            # Step 1: Ingest
            st.write("📥 Reading CoStar spreadsheets…")
            file_tuples = [(f, f.name) for f in uploaded_files]
            combined, reports = ingest_files(file_tuples)

            if combined.empty:
                st.error("No data could be parsed from the uploaded files.")
                st.stop()

            markets = combined["Market"].unique()
            st.write(f"✅ Parsed **{len(markets)} markets** ({len(combined):,} rows)")

            issues = [r for r in reports if r["issues"]]
            if issues:
                with st.expander(f"⚠️ {len(issues)} file(s) had minor issues (auto-fixed)"):
                    for r in issues:
                        st.caption(f"**{r['market']}**: {', '.join(r['issues'])}")

            # Step 2: Census enhancement (optional)
            use_census = False
            census_snapshot = pd.DataFrame()
            unmatched = []

            if census_key:
                st.write("🏛️ Fetching Census data…")
                try:
                    combined, census_snapshot, unmatched = enhance_with_census(
                        combined,
                        census_key,
                        progress_callback=lambda msg: st.write(msg),
                        cbsa_overrides=cbsa_overrides,
                    )
                    matched_count = len(markets) - len(unmatched)
                    if not census_snapshot.empty and matched_count > 0:
                        use_census = True
                        st.write(f"✅ Census data merged for **{matched_count}** market(s)")
                    else:
                        st.write("⚠️ Census data unavailable — using CoStar-only index")
                except Exception as exc:
                    st.write(f"⚠️ Census API error: {exc} — using CoStar-only index")
            else:
                st.write("ℹ️ No Census API key — using CoStar-only index")

            # Step 3: Filter to matched markets only when using Census
            st.write("📊 Extracting latest quarter per market…")
            latest = get_latest_quarter(combined)

            if latest.empty:
                st.error("No complete quarterly data found.")
                st.stop()

            # Merge prior-year vacancy (computed from full history before filtering)
            lagged = compute_lagged_features(combined)
            latest = latest.merge(lagged, on="Market", how="left")

            if use_census and unmatched:
                # Exclude markets without Census data from the full-census ranking
                latest = latest[~latest["Market"].isin(unmatched)].copy()
                st.write(
                    f"✅ Using **{len(latest)} markets** with full Census data  "
                    f"*(excluded {len(unmatched)} without CBSA match)*"
                )
            else:
                periods = latest["Period"].unique() if "Period" in latest.columns else []
                st.write(f"✅ Latest periods: {', '.join(str(p) for p in periods)}")

            # Step 4: Compute index
            st.write("🔢 Computing Demand Index…")
            if model_type == MODEL_ABS_SUPPLY:
                rankings, components = compute_absorption_supply_index(latest)
            else:
                rankings, components = compute_demand_index(latest)
            st.write(f"✅ Ranked **{len(rankings)} markets**")

            # Step 5: Generate Excel
            st.write("📄 Generating Excel workbook…")
            excel_bytes = generate_workbook(
                rankings, components,
                census_snapshot=census_snapshot if use_census else None,
                use_census=use_census,
            )
            st.write("✅ Workbook ready")

            status.update(label="✅ Processing complete!", state="complete")

        return {
            "rankings":       rankings,
            "components":     components,
            "excel_bytes":    excel_bytes,
            "use_census":     use_census,
            "unmatched":      unmatched,
            "combined_raw":   combined,  # kept for re-runs
        }

    if process_btn or cbsa_rerun:
        cbsa_overrides = st.session_state.get("cbsa_overrides", {})
        result = _run_pipeline(uploaded_files, census_key, cbsa_overrides, model_type)
        for k, v in result.items():
            st.session_state[k] = v

    # ── Display results ──────────────────────────────────────────────────
    if "rankings" not in st.session_state:
        st.stop()

    rankings = st.session_state["rankings"]
    use_census = st.session_state.get("use_census", False)

    st.divider()
    st.subheader("3 · Results")

    # Summary metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Markets Ranked", len(rankings))
    m2.metric("Index Model", "Full (Census + CoStar)" if use_census else "CoStar Only")

    top_market = rankings.iloc[0]
    bottom_market = rankings.iloc[-1]
    m3.metric("Top Market", f"{top_market['Market']}", f"{top_market['Demand_Index']:.1f}")
    m4.metric("Bottom Market", f"{bottom_market['Market']}", f"{bottom_market['Demand_Index']:.1f}")

    # Rankings table
    st.markdown("#### Market Rankings")

    display_cols = ["Rank", "Market", "Demand_Index", "Tier"]
    optional_display = [
        "Period", "Effective Rent % Growth/Yr", "Vacancy Percent",
        "Occupancy Percent", "Absorption Units", "Under Construction Percent",
        "Deliveries Percent", "Inventory Units",
    ]
    display_cols += [c for c in optional_display if c in rankings.columns]

    col_cfg = {
        "Demand_Index": st.column_config.ProgressColumn(
            "Demand Index",
            help="0 = lowest demand, 100 = highest demand",
            min_value=0,
            max_value=100,
            format="%.1f",
        ),
        "Rank": st.column_config.NumberColumn("Rank", format="%d"),
    }
    # Friendly percent/units formatting via column_config
    for c in display_cols:
        if ("Percent" in c or "Growth" in c) and c not in col_cfg:
            col_cfg[c] = st.column_config.NumberColumn(c, format="%.2f %%")
        elif "Units" in c and c not in col_cfg:
            col_cfg[c] = st.column_config.NumberColumn(c, format="%d")

    st.dataframe(
        rankings[display_cols],
        column_config=col_cfg,
        use_container_width=True,
        height=min(len(rankings) * 40 + 40, 600),
    )

    # Bar chart — top / bottom 10
    import matplotlib.pyplot as plt
    col_chart1, col_chart2 = st.columns(2)

    def _hbar(ax, data, color, title):
        """Horizontal bar chart sorted so highest score is at top."""
        data = data.sort_values("Demand_Index", ascending=True)  # barh draws bottom-up
        labels = data["Market"].str.replace("_", " ")
        values = data["Demand_Index"]
        bars = ax.barh(labels, values, color=color, height=0.6)
        ax.set_xlim(0, 100)
        ax.set_xlabel("Demand Index", fontsize=9)
        ax.set_title(title, fontsize=10, fontweight="bold", pad=8)
        ax.tick_params(axis="y", labelsize=8)
        ax.tick_params(axis="x", labelsize=8)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        # Value labels inside bars
        for bar, val in zip(bars, values):
            ax.text(
                min(val + 1, 98), bar.get_y() + bar.get_height() / 2,
                f"{val:.1f}", va="center", ha="left", fontsize=7.5, color="#333333",
            )

    with col_chart1:
        top10 = rankings.head(10)[["Market", "Demand_Index"]].copy()
        fig, ax = plt.subplots(figsize=(5.5, 4))
        _hbar(ax, top10, "#4ade80", "🟢 Top 10 Markets")
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    with col_chart2:
        bot10 = rankings.tail(10)[["Market", "Demand_Index"]].copy()
        fig, ax = plt.subplots(figsize=(5.5, 4))
        _hbar(ax, bot10, "#f87171", "🔴 Bottom 10 Markets")
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # ── Download ─────────────────────────────────────────────────────────
    st.divider()
    st.subheader("4 · Download")

    st.download_button(
        label="📥 Download Excel Workbook",
        data=st.session_state["excel_bytes"],
        file_name="multifamily_demand_index.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    sheets = ["Market Rankings", "Index Components"]
    if use_census:
        sheets.append("Census Demographics")
    sheets.append("Methodology")
    st.caption(f"Workbook contains {len(sheets)} sheets: {', '.join(sheets)}")

    # ── Validation backtest ──────────────────────────────────────────────
    st.divider()
    with st.expander("📊 5 · Validation Backtest", expanded=False):
        st.markdown(
            "Tests whether the Demand Index would have predicted **forward rent "
            "growth** in the uploaded historical data.\n\n"
            "Forward growth is computed from effective rent *levels* — not CoStar's "
            "YoY column — to avoid information overlap. "
            "(CoStar's YoY at $t+1$ shares 3 of 4 quarters with the value at $t$, "
            "making it a poor out-of-sample target.)\n\n"
            "Rent growth is **excluded** from the index inputs, so any correlation "
            "represents genuine predictive signal."
        )

        horizon = st.radio(
            "Forward horizon",
            options=["4Q", "1Q"],
            index=0,
            horizontal=True,
            help=(
                "**4Q (1 year)**: rent 4 quarters ahead ÷ rent today − 1. "
                "Aligns with investment timelines and has zero information overlap.\n\n"
                "**1Q (1 quarter)**: rent next quarter ÷ rent today − 1. "
                "Shorter signal, noisier, but also zero overlap."
            ),
        )

        backtest_btn = st.button("🔬 Run Backtest", type="secondary")

        if backtest_btn:
            combined_raw = st.session_state.get("combined_raw")
            if combined_raw is None or combined_raw.empty:
                st.error("No raw data available — please re-run the pipeline above.")
            else:
                with st.status("Running backtest…", expanded=True) as bt_status:
                    st.write("🔢 Computing Demand Index per quarter…")
                    bt_use_census = st.session_state.get("use_census", False)
                    results, bt_stats = run_backtest(combined_raw, bt_use_census)

                    if results.empty:
                        st.error(bt_stats.get("error", "Backtest produced no results."))
                    else:
                        st.write(
                            f"✅ {bt_stats['total_quarters']} quarters × "
                            f"{bt_stats['total_markets']} markets = "
                            f"{len(results):,} observations"
                        )
                        st.write(f"Model: **{bt_stats['model']}**")
                        bt_status.update(label="✅ Backtest complete!", state="complete")

                st.session_state["bt_results"] = results
                st.session_state["bt_stats"] = bt_stats
                st.session_state["bt_horizon"] = horizon

        # Display results if available
        if "bt_results" in st.session_state:
            results = st.session_state["bt_results"]
            bt_stats = st.session_state["bt_stats"]
            bt_horizon = st.session_state.get("bt_horizon", "4Q")
            h_stats = bt_stats.get(bt_horizon, {})

            if "error" not in h_stats and h_stats.get("n", 0) >= 10:
                # ── Metrics row ────────────────────────────────────────
                horizon_label = "1-Quarter" if bt_horizon == "1Q" else "4-Quarter"
                st.markdown(f"#### {horizon_label} Forward Results")

                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("Pearson r", f"{h_stats['pearson_r']:.3f}")
                mc2.metric("Spearman ρ", f"{h_stats['spearman_rho']:.3f}")
                mc3.metric("R²", f"{h_stats['r_squared']:.3f}")
                mc4.metric("Observations", f"{h_stats['n']:,}")

                p_val = h_stats["pearson_p"]
                if p_val < 0.001:
                    st.success(f"Statistically significant (p = {p_val:.2e})")
                elif p_val < 0.05:
                    st.success(f"Statistically significant (p = {p_val:.4f})")
                else:
                    st.warning(f"Not statistically significant (p = {p_val:.4f})")

                # ── Charts side-by-side ────────────────────────────────
                ch1, ch2 = st.columns(2)

                with ch1:
                    fig_scatter = scatter_chart(results, bt_stats, bt_horizon)
                    st.pyplot(fig_scatter, use_container_width=True)

                    # Download button for scatter
                    import io as _io
                    buf_s = _io.BytesIO()
                    fig_scatter.savefig(buf_s, format="png", dpi=300, bbox_inches="tight",
                                        facecolor="white")
                    buf_s.seek(0)
                    st.download_button(
                        "📥 Download scatter chart",
                        data=buf_s,
                        file_name=f"validation_scatter_{bt_horizon}.png",
                        mime="image/png",
                    )
                    plt.close(fig_scatter)

                with ch2:
                    fig_box = tier_boxplot(results, bt_horizon)
                    st.pyplot(fig_box, use_container_width=True)

                    buf_b = _io.BytesIO()
                    fig_box.savefig(buf_b, format="png", dpi=300, bbox_inches="tight",
                                    facecolor="white")
                    buf_b.seek(0)
                    st.download_button(
                        "📥 Download tier box plot",
                        data=buf_b,
                        file_name=f"validation_boxplot_{bt_horizon}.png",
                        mime="image/png",
                    )
                    plt.close(fig_box)

                # ── Tier summary table ─────────────────────────────────
                st.markdown("#### Per-Tier Summary")
                tier_data = h_stats.get("tiers", {})
                tier_rows = []
                for tier in ["High Demand", "Moderate Demand", "Low Demand"]:
                    t = tier_data.get(tier, {})
                    if t.get("n", 0) > 0:
                        tier_rows.append({
                            "Tier": tier,
                            "Observations": t["n"],
                            f"Median Fwd {bt_horizon} Growth": f"{t['median'] * 100:.2f}%",
                            f"Mean Fwd {bt_horizon} Growth": f"{t['mean'] * 100:.2f}%",
                            "IQR (25th–75th)": f"{t['q25'] * 100:.2f}% – {t['q75'] * 100:.2f}%",
                        })
                if tier_rows:
                    st.dataframe(pd.DataFrame(tier_rows), use_container_width=True)

            else:
                n = h_stats.get("n", 0)
                if n < 10:
                    st.warning(
                        f"Only {n} observations available for the {bt_horizon} horizon. "
                        "Need at least 10 for meaningful statistics. "
                        "Try uploading more historical quarters."
                    )
                else:
                    st.error(h_stats.get("error", "Unknown error during backtest."))

    # ── Missing-market CBSA override UI ──────────────────────────────────
    unmatched = st.session_state.get("unmatched", [])
    if use_census and unmatched:
        st.divider()
        st.subheader("6 · Missing Markets")
        st.warning(
            f"**{len(unmatched)} market(s) could not be matched to a Census CBSA code "
            "and were excluded from the ranked results above.**\n\n"
            "If you know the CBSA code(s), enter them below and re-run to include "
            "those markets in the full Census-enhanced index.\n\n"
            "You can look up CBSA codes at "
            "[census.gov/geographies/reference-files.html]"
            "(https://www.census.gov/geographies/reference-files.html) "
            "or by searching for the metro area on "
            "[data.census.gov](https://data.census.gov).",
            icon="⚠️",
        )

        overrides = dict(st.session_state.get("cbsa_overrides", {}))
        any_filled = False
        for mkt in unmatched:
            code = st.text_input(
                f"CBSA code for **{mkt}**",
                value=overrides.get(mkt, ""),
                placeholder="e.g. 35620",
                key=f"cbsa_input_{mkt}",
            )
            clean = code.strip()
            if clean:
                overrides[mkt] = clean
                any_filled = True
            else:
                overrides.pop(mkt, None)

        rerun_btn = st.button(
            "🔄 Re-run with CBSA codes",
            type="primary",
            disabled=not any_filled,
        )
        if rerun_btn and any_filled:
            st.session_state["cbsa_overrides"] = overrides
            # Clear previous pipeline results so _run_pipeline fires again
            for key in ("rankings", "components", "excel_bytes", "use_census", "unmatched"):
                st.session_state.pop(key, None)
            st.rerun()


if __name__ == "__main__":
    main()
