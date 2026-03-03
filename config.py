"""
Configuration for the Multifamily Demand Index App.

Index methodology grounded in:
- NMHC Research Notes (Bruen, Dec 2025): Employment growth ↔ rent growth significant
  in 94/99 quarters; deliveries ↔ lower rent growth in 93/99 quarters.
- NMHC Research Notes (Bruen, Dec 2024): Absorptions, employment growth, income growth,
  and vacancy rate all statistically significant predictors across 491 markets / 10 years.
- Real estate economics (De Leeuw 1971, Polinsky & Ellwood 1979): Housing demand driven
  by demographics (population), income, price; supply adjusts via construction pipeline.
- Apartment List National Rent Report (Feb 2026): Vacancy rate as primary market
  tightness indicator; construction surge → soft conditions in Sun Belt.
- Harvard JCHS America's Rental Housing 2024: Cost burdens, affordability, supply gaps.
"""

# ── Index Component Weights ──────────────────────────────────────────────────
# Positive = higher value → stronger demand signal
# Negative = higher value → supply headwind (drag on demand)

INDEX_WEIGHTS_FULL = {
    # Demand Fundamentals (CoStar)
    # Absorption (+16.7 bps per ppt — NMHC Dec 2024 panel regression)
    "absorption_pct":       0.25,   # Net absorption as % of inventory

    # Demand Drivers (Census)
    # Employment growth is the *change* in employment rate, not its level.
    # NMHC Dec 2024 alternate model: +19.8 bps — largest single demand predictor.
    "employment_growth":    0.25,   # YoY change in employment rate
    # Population growth: De Leeuw 1971, Polinsky & Ellwood 1979
    "population_growth":    0.15,   # YoY population growth
    # Income growth: +5.5 bps (NMHC Dec 2024) — real but smaller effect
    "income_growth":        0.10,   # YoY median household income growth

    # Supply / Slack Pressure
    # Prior-year vacancy is the LARGEST suppressive factor in the NMHC research:
    # −24.5 to −27.7 bps per ppt (NMHC Dec 2024).  Using a lagged value avoids
    # simultaneity with the current period's demand signals.
    "prior_vacancy":       -0.15,   # Prior-year vacancy rate (CoStar, lagged 4 quarters)
    # Delivery rate: −5.2 to −7.7 bps (NMHC Dec 2024), confirmed in 93/99 quarters
    # (NMHC Dec 2025).  Under-construction is excluded — it reflects future, not
    # current, supply pressure and would double-count with deliveries.
    "deliveries_pct":      -0.10,   # Deliveries as % of inventory
}

# ── Alternative Model: Absorption / Supply Pressure ────────────────────────
# score = absorption_units / (inventory_units * vacancy_pct/100 + uc_units)
#
# Numerator  : quarterly net absorption (units actually leased)
# Denominator: total "contested" unit pool = units already vacant today
#              + under-construction pipeline arriving soon
#
# Interpretation: higher ratio → demand is outrunning available + incoming
# supply; conceptually the inverse of months-of-supply at current absorption.
# Entirely CoStar-derived. Rescaled to 0-100 the same way as the weighted model.

# When Census data is unavailable, rebalance to CoStar-only weights.
# Absorption is the dominant observable demand signal; occupancy captures
# current tightness; prior-year vacancy provides the key lagged slack indicator
# found in the NMHC research; deliveries capture near-term supply impact.
# Rent growth is intentionally excluded from both models — it is the *outcome*
# variable in the research and including it as an input would create circularity.
INDEX_WEIGHTS_COSTAR_ONLY = {
    "absorption_pct":       0.40,   # Primary observable demand signal
    "occupancy":            0.25,   # Current market tightness (inverse of vacancy)
    "prior_vacancy":       -0.20,   # Prior-year vacancy, largest suppressive factor
    "deliveries_pct":      -0.15,   # Current supply delivery pressure
}

# ── CoStar Column Mapping ────────────────────────────────────────────────────
# Maps our internal names to expected CoStar spreadsheet column names.
# Column names have leading whitespace stripped during ingestion.

COSTAR_COLUMNS = {
    "period":               "Period",
    "inventory_units":      "Inventory Units",
    "inventory_bldgs":      "Inventory Bldgs",
    "inventory_avg_sf":     "Inventory Avg SF",
    "asking_rent_unit":     "Asking Rent Per Unit",
    "asking_rent_sf":       "Asking Rent Per SF",
    "asking_rent_growth":   "Asking Rent % Growth/Yr",
    "effective_rent_unit":  "Effective Rent Per Unit",
    "effective_rent_sf":    "Effective Rent Per SF",
    "effective_rent_growth":"Effective Rent % Growth/Yr",
    "concessions":          "Effective Rent Concessions %",
    "vacancy_units":        "Vacancy Units",
    "vacancy_pct":          "Vacancy Percent",
    "vacancy_growth":       "Vacancy % Growth/Yr",
    "occupancy_units":      "Occupancy Units",
    "occupancy_pct":        "Occupancy Percent",
    "occupancy_growth":     "Occupancy % Growth/Yr",
    "absorption_units":     "Absorption Units",
    "absorption_pct":       "Absorption Percent",
    "uc_bldgs":             "Under Construction Bldgs",
    "uc_units":             "Under Construction Units",
    "uc_pct":               "Under Construction Percent",
    "deliveries_bldgs":     "Deliveries Bldgs",
    "deliveries_units":     "Deliveries Units",
    "deliveries_pct":       "Deliveries Percent",
}

# Columns that may arrive as mixed types and need coercion
NUMERIC_COERCE_COLUMNS = [
    "Asking Rent % Growth/Yr",
    "Effective Rent % Growth/Yr",
    "Vacancy % Growth/Yr",
    "Occupancy % Growth/Yr",
    "Absorption Units",
    "Absorption Percent",
]

# ── Key output columns for the Excel workbook ────────────────────────────────
OUTPUT_METRICS = [
    ("Inventory Units",              "{:,.0f}",    "Total apartment units in the market"),
    ("Asking Rent Per Unit",         "${:,.0f}",   "Average asking rent per unit"),
    ("Effective Rent Per SF",        "${:,.2f}",   "Effective rent per square foot"),
    ("Effective Rent % Growth/Yr",   "{:.1%}",     "Year-over-year effective rent growth"),
    ("Vacancy Percent",              "{:.1%}",     "Current vacancy rate"),
    ("Occupancy Percent",            "{:.1%}",     "Current occupancy rate"),
    ("Absorption Units",             "{:,.0f}",    "Net units absorbed (latest quarter)"),
    ("Under Construction Units",     "{:,.0f}",    "Units currently under construction"),
    ("Under Construction Percent",   "{:.1%}",     "Under construction as % of inventory"),
    ("Deliveries Units",             "{:,.0f}",    "Units delivered (latest quarter)"),
    ("Deliveries Percent",           "{:.1%}",     "Deliveries as % of inventory"),
    ("Effective Rent Concessions %", "{:.1%}",     "Concession rate"),
]

# ── CBSA Code Mapping ────────────────────────────────────────────────────────
# Maps Market filename stems to Census Bureau CBSA search terms.
# Extended to cover common US metros beyond the original 29.

CBSA_SEARCH_TERMS = {
    "Albuquerque_NM":             "Albuquerque",
    "Austin_TX":                  "Austin-Round Rock",
    "Bismark_ND":                 "Bismarck",
    "Boise_ID":                   "Boise City",
    "Charlottesville_VA":         "Charlottesville",
    "Chattanooga_TN":             "Chattanooga",
    "Clarksville_TN":             "Clarksville",
    "College_Station_Bryan_TX":   "College Station-Bryan",
    "Columbus_OH":                "Columbus",
    "Dallas_Fort_Worth_TX":       "Dallas-Fort Worth-Arlington",
    "Denver_CO":                  "Denver-Aurora",
    "Farmington_NM":              "Farmington",
    "Helena_MT":                  "Helena",
    "Houston_TX":                 "Houston",
    "Kansas_City_MO":             "Kansas City",
    "Knoxville_TN":               "Knoxville",
    "Nashville_TN":               "Nashville-Davidson",
    "Northwest_Arkansas_AK":      "Fayetteville-Springdale-Rogers",
    "Oklahoma_City_OK":           "Oklahoma City",
    "Omaha_NE":                   "Omaha-Council Bluffs",
    "Phoenix_AZ":                 "Phoenix-Mesa",
    "Rapid_City_SD":              "Rapid City",
    "Richmond_VA":                "Richmond",
    "Saint_Louis_MO":             "St. Louis",
    "Salt_Lake_City_UT":          "Salt Lake City",
    "San_Antonio_TX":             "San Antonio",
    "Sioux_Falls_SD":             "Sioux Falls",
    "Topeka_KS":                  "Topeka",
    "Wilmington_NC":              "Wilmington",
    # Common additional metros an analyst might upload
    "Atlanta_GA":                 "Atlanta",
    "Charlotte_NC":               "Charlotte",
    "Chicago_IL":                 "Chicago",
    "Indianapolis_IN":            "Indianapolis",
    "Jacksonville_FL":            "Jacksonville",
    "Las_Vegas_NV":               "Las Vegas",
    "Los_Angeles_CA":             "Los Angeles",
    "Miami_FL":                   "Miami",
    "Minneapolis_MN":             "Minneapolis",
    "New_York_NY":                "New York",
    "Orlando_FL":                 "Orlando",
    "Philadelphia_PA":            "Philadelphia",
    "Portland_OR":                "Portland",
    "Raleigh_NC":                 "Raleigh",
    "San_Diego_CA":               "San Diego",
    "San_Francisco_CA":           "San Francisco",
    "Seattle_WA":                 "Seattle",
    "Tampa_FL":                   "Tampa",
    "Washington_DC":              "Washington",
}

# Census ACS variables needed for the demand index
CENSUS_VARIABLES = {
    "B01003_001E": "Population",
    "B19013_001E": "Median_Household_Income",
    "B23025_002E": "Labor_Force",
    "B23025_004E": "Employed",
    "B07001_065E": "In_Migration",
}

# ── Demand Index Tier Thresholds ─────────────────────────────────────────────
TIER_THRESHOLDS = {
    "High Demand":     67,   # Index >= 67
    "Moderate Demand": 33,   # Index >= 33
    "Low Demand":       0,   # Index < 33
}

# ── App Metadata ─────────────────────────────────────────────────────────────
APP_TITLE = "Multifamily Demand Index"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = """
Upload CoStar multifamily market spreadsheets to generate a research-backed
**Demand Index** that ranks markets by the strength of apartment demand
relative to supply conditions.
"""

METHODOLOGY_TEXT = """
# Demand Index Methodology

## Overview
The Multifamily Demand Index is a composite score (0–100) that ranks markets by
the relative strength of apartment demand versus supply conditions. Higher scores
indicate stronger demand fundamentals and more favorable supply-demand dynamics
for landlords and investors.

## Research Foundation
This index is grounded in peer-reviewed and institutional research on multifamily
housing demand drivers:

1. **NMHC Research Notes** (Bruen, December 2025) — "Unpacking the Relationship
   Between Jobs and Apartment Demand": Panel regression across 150 largest
   apartment markets (2001–2025) found employment growth has a meaningful,
   positive impact on rents in 94 of 99 quarters studied. Higher delivery rates
   were associated with lower rent growth in 93 of 99 quarters.
   → [nmhc.org](https://www.nmhc.org/research-insight/research-notes/2025/unpacking-the-relationship-between-jobs-and-apartment-demand/)

2. **NMHC Research Notes** (Bruen, December 2024) — "How New Supply Impacts
   Affordability Across the Board": Statistical model across 491 CoStar markets
   over 10 years found that absorptions (+16.7 bps), employment growth (+19.8 bps),
   income growth (+5.5 bps), and prior-year vacancy rate (−24.5 to −27.7 bps) are
   all statistically significant predictors of rent growth.
   → [nmhc.org](https://www.nmhc.org/research-insight/research-notes/2024/how-new-supply-impacts-affordability-across-the-board/)

3. **NMHC Research Notes** (Bruen, October 2025) — "Reconciling the Apartment
   Shortage with Recent Record Apartment Completions": Distinguishes short-term
   supply surpluses from structural housing shortages driven by latent demand.
   → [nmhc.org](https://www.nmhc.org/research-insight/research-notes/2025/reconciling-the-apartment-shortage-with-recent-record-apartment-completions/)

4. **Real Estate Economics** (De Leeuw 1971; Polinsky & Ellwood 1979) — Housing
   demand is a function of demographics (population size & growth), income
   (elasticity 0.5–0.9), and price.
   → De Leeuw (1971) "The Demand for Housing: A Review of Cross-Section Evidence"
     *Review of Economics and Statistics* 53(1):1–10
     [jstor.org/stable/1925374](https://www.jstor.org/stable/1925374)
   → Polinsky & Ellwood (1979) "An Empirical Reconciliation of Micro and Grouped Estimates of the Demand for Housing"
     *Review of Economics and Statistics* 61(2):199–205
     [jstor.org/stable/1924587](https://www.jstor.org/stable/1924587)

5. **Apartment List National Rent Report** (February 2026) — Vacancy rate as
   primary market tightness indicator; markets with highest construction (Austin,
   Denver, Phoenix) show sharpest rent declines.
   → [apartmentlist.com](https://www.apartmentlist.com/research/national-rent-data)

6. **Harvard JCHS** — America's Rental Housing 2024: Documents the relationship
   between supply constraints, cost burdens, and market-level affordability.
   → [jchs.harvard.edu](https://www.jchs.harvard.edu/sites/default/files/reports/files/Harvard_JCHS_Americas_Rental_Housing_2024.pdf)

## Components & Weights

Weights are proportional to empirical coefficient magnitudes from NMHC Dec 2024.
Rent growth is **not** an input — it is the *dependent variable* in the research
regressions and would create circularity. It is better used as a validation check
on index rankings.

### With Census Data (Full Model)
| Component | Weight | Direction | Source | Research basis |
|-----------|--------|-----------|--------|----------------|
| Net Absorption (% of Inventory) | 25% | + | CoStar | +16.7 bps (NMHC Dec 2024) |
| Employment Growth (YoY) | 25% | + | Census ACS | +19.8 bps — *change*, not level (NMHC Dec 2024) |
| Population Growth (YoY) | 15% | + | Census ACS | De Leeuw 1971; Polinsky & Ellwood 1979 |
| Income Growth (YoY) | 10% | + | Census ACS | +5.5 bps (NMHC Dec 2024) |
| Prior-Year Vacancy Rate | 15% | − | CoStar (lagged) | −24.5 to −27.7 bps — largest factor (NMHC Dec 2024) |
| Deliveries (% of Inventory) | 10% | − | CoStar | −5.2 to −7.7 bps; 93/99 quarters (NMHC Dec 2025) |

### CoStar-Only Model (No Census Data)
| Component | Weight | Direction | Source | Notes |
|-----------|--------|-----------|--------|-------|
| Net Absorption (% of Inventory) | 40% | + | CoStar | Primary observable demand signal |
| Occupancy Rate | 25% | + | CoStar | Current market tightness |
| Prior-Year Vacancy Rate | 20% | − | CoStar (lagged) | Largest suppressive factor per research |
| Deliveries (% of Inventory) | 15% | − | CoStar | Near-term supply glut |

**Note on excluded components:**
- *Rent growth*: excluded as an input — it is the outcome variable in both NMHC
  regressions. Use it to validate index rankings post-hoc.
- *Under-construction %*: excluded — it reflects future supply not yet hitting the
  market, and would double-count with deliveries (which already capture the
  near-term supply impact with a measured, significant coefficient).
- *Employment rate level*: replaced by employment growth — NMHC Dec 2024 found
  the year-over-year *change* in employment (+19.8 bps) matters, not the level.

## Calculation
1. Extract the **latest complete quarter** per market (excluding QTD rows).
2. Compute **prior-year vacancy** from the same quarter one year earlier in the
   historical dataset — the single strongest suppressive predictor in the research.
3. For each component, compute the **cross-sectional z-score** across all markets:
   z = (x − μ) / σ
4. Multiply each z-score by its signed weight (negative for supply/vacancy).
5. Sum the weighted z-scores into a raw composite score.
6. **Rescale to 0–100**: Index = (raw − min) / (max − min) × 100.
7. Classify into tiers: **High Demand** (≥67), **Moderate Demand** (33–66),
   **Low Demand** (<33).
"""
