# Multifamily Demand Index App

A research-backed tool that scores and ranks multifamily markets on a **0–100 Demand Index** using CoStar data — optionally enriched with U.S. Census demographics. Designed for investment teams who need a fast, auditable way to screen metros for value-add multifamily expansion.

---

## What It Does

1. You upload one or more **CoStar multifamily Excel files** (one per market)
2. The app computes a **Demand Index score (0–100)** for each market using absorption, occupancy, vacancy, deliveries, and (optionally) Census demographics
3. Markets are classified as **High Demand** (≥ 67), **Moderate Demand** (33–67), or **Low Demand** (< 33)
4. You download a formatted **Excel workbook** with rankings, component breakdowns, and a methodology tab

---

## Before You Start — What You Need

### 1. Python 3.9 or newer

Check if you already have it by opening Terminal (Mac/Linux) or Command Prompt (Windows) and typing:

```
python3 --version
```

If it says `Python 3.9.x` or higher, you're good. If not — or if you get an error — download and install Python from **[python.org/downloads](https://www.python.org/downloads/)**.
➡️ During installation on Windows, check the box that says **"Add Python to PATH"**.

### 2. CoStar Multifamily Export Files

You need one `.xlsx` file per market, exported from CoStar's multifamily platform. The file should contain quarterly historical data with columns like:

- `Period` (e.g. `2025 Q4`)
- `Vacancy Percent`, `Occupancy Percent`
- `Absorption Units`
- `Effective Rent/Unit`
- `Under Construction Units`, `Deliveries Units`

**The market name is read from the filename** — so name your files clearly:
✅ `Austin_TX.xlsx` → market name: `Austin TX`
✅ `Denver_CO.xlsx` → market name: `Denver CO`

### 3. (Optional) A Free Census API Key

The Census API key unlocks demographic enrichment (population growth, income, employment). The app works without it but will use CoStar-only weights.

To get a free key:
1. Go to **[api.census.gov/data/key_signup.html](https://api.census.gov/data/key_signup.html)**
2. Fill in your name and email
3. Check your email — the key arrives in a few minutes
4. It looks like: `a1b2c3d4e5f6...` (a long string of letters and numbers)

---

## Getting Started — Step by Step

### Mac or Linux

1. **Download or clone this folder** to your computer
2. Open **Terminal** and navigate to the folder:
   ```
   cd path/to/demand_index_app
   ```
3. Run the launcher:
   ```
   ./run.sh
   ```
4. The app will open automatically in your browser at `http://localhost:8501`

> The first time you run it, the launcher creates a virtual environment and installs all dependencies — this takes about 1–2 minutes. Subsequent runs are instant.

### Windows

1. **Download or clone this folder** to your computer
2. Double-click **`run.bat`**
3. A terminal window opens and installs dependencies, then the app opens in your browser at `http://localhost:8501`

> If you see a "Windows protected your PC" warning, click **"More info" → "Run anyway"** — this is a standard Python script, not a virus.

### Manual Setup (Any OS)

If the launcher scripts don't work:

```bash
# From inside the demand_index_app folder:
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate.bat
pip install -r requirements.txt
streamlit run app.py
```

---

## Using the App — Step by Step

Once the app is open in your browser:

**Step 1 — Upload CoStar Files**
Click "Browse files" and select one or more `.xlsx` CoStar export files. You'll see a preview table confirming the data was read correctly.

**Step 2 — (Optional) Enter Census API Key**
Paste your Census API key into the text box. The app will automatically pull population, income, and employment data for each market. If you skip this, the app uses CoStar-only weights.

**Step 3 — Compute the Index**
Click the **"Compute Demand Index"** button. Results appear in seconds.

**Step 4 — Review Results**
- A ranked table shows every market with its score (0–100) and tier (High / Moderate / Low)
- Color-coded bar charts show each component's contribution
- Expand any market to see its full component breakdown

**Step 5 — Run Backtest (Optional)**
Click "Run Backtest" to validate the index's predictive accuracy against the historical data you uploaded. This shows you how well the index would have predicted rent growth in prior quarters.

**Step 6 — Download the Excel Report**
Click **"Download Excel Report"** to save a formatted workbook with:
- Market Rankings tab (color-coded by tier)
- Component Breakdown tab (z-scores and weighted contributions)
- Census Demographics tab (if Census key was used)
- Methodology tab (full documentation for LP / IC review)

---

## Demand Index Methodology

### Components & Weights

**Full Model (with Census data — 6 components):**

| Component | Source | Weight | Logic |
|---|---|---|---|
| Absorption % of Inventory | CoStar | +25% | Higher absorption = units being leased, real demand |
| Employment Growth | Census ACS | +25% | Job growth drives household formation and rent capacity |
| Population Growth | Census ACS | +15% | More people = more housing demand |
| Income Growth | Census ACS | +10% | Rising incomes = tenants can afford higher rents |
| Prior-Year Vacancy | CoStar | −15% | Higher vacancy = weaker market (inverted) |
| Deliveries % of Inventory | CoStar | −10% | More new supply = near-term rent pressure (inverted) |

**CoStar-Only Model (without Census key — 4 components):**

| Component | Weight |
|---|---|
| Absorption % of Inventory | +40% |
| Occupancy | +25% |
| Prior-Year Vacancy | −20% |
| Deliveries % of Inventory | −15% |

### How the Score Is Calculated

1. For each component, compute the **z-score** across all markets in the uploaded set (this standardizes different units — percent, units, dollars — onto a common scale)
2. Multiply each z-score by its weight
3. Sum weighted z-scores into a raw composite
4. Rescale to 0–100: `Index = (raw − min) / (max − min) × 100`
5. Classify: **High Demand** ≥ 67 · **Moderate Demand** 33–66 · **Low Demand** < 33

> **Important:** Scores are relative — they rank markets against each other within your uploaded set. Adding or removing markets will shift scores slightly.

### Research Basis

Weights are derived from peer-reviewed academic and industry research:
- **NMHC Research Notes** (Bruen, Dec 2024): Panel regression across 491 metros — absorption, employment growth, income growth, and vacancy are all statistically significant predictors of rent growth
- **NMHC Research Notes** (Bruen, Dec 2025): Employment growth significant in 94 of 99 quarters; deliveries suppress rent growth in 93 of 99 quarters
- **Harvard JCHS** (2024): Structural housing shortage reinforces fundamental demand drivers
- **Apartment List** (Feb 2026): National vacancy at 7.4%; Sun Belt supply wave validates deliveries as a risk factor

### Backtested Predictive Accuracy (CoStar-Only Model)

| Metric | Value |
|---|---|
| Pearson r | 0.141 |
| p-value | 2.49 × 10⁻¹⁴ |
| R² | 0.020 |
| Observations | 2,900 market-quarters (29 markets, ~100 quarters) |
| Horizon | 4-quarter forward rent growth |

**Tier median forward rent growth (4 quarters out):**
- High Demand: **2.18%**
- Moderate Demand: **1.87%**
- Low Demand: **1.40%**

The index produces statistically significant, monotonically ordered tier separation. It is a **screening tool**, not a rent forecast — it narrows the aperture for deeper underwriting, it does not replace it.

---

## Excel Output

| Sheet | Contents |
|---|---|
| **Market Rankings** | All markets ranked by Demand Index, color-coded by tier |
| **Index Components** | Raw input values, z-scores, and weighted contributions per component |
| **Census Demographics** | Population, income, employment data per market (if Census key provided) |
| **Methodology** | Full documentation of weights, formulas, and research citations |

---

## Troubleshooting

| Issue | What to Do |
|---|---|
| `python3: command not found` | Install Python 3.9+ from [python.org](https://python.org) — on Windows check "Add to PATH" |
| `Permission denied: ./run.sh` | Run `chmod +x run.sh` in Terminal first |
| App doesn't open in browser | Go to [http://localhost:8501](http://localhost:8501) manually |
| File upload fails | Make sure it's a CoStar `.xlsx` export (not `.csv` or `.xls`) |
| Census data shows "N/A" | Check your API key is correct and has been activated (check email) |
| Market name looks wrong | Rename the file — `Austin_TX.xlsx` → displayed as `Austin TX` |
| Port already in use | Close other browser tabs running Streamlit, or run `streamlit run app.py --server.port 8502` |

---

## Requirements

- Python 3.9+
- Internet connection (for Census API enrichment only — the core app works offline)

All Python dependencies are in `requirements.txt` and installed automatically by the launcher scripts.
