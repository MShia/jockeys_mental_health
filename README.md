# Jockey Mental Wellbeing Dashboard
## IHRB × DCU School of Health & Human Performance

Interactive analytics dashboard for the Jockey Mental Wellbeing Survey data, covering common mental disorders (CMDs), risk factors, coping strategies, social support, weight-making behaviours, and multivariable regression modelling.

---

## Prerequisites

You need **Node.js** (version 18 or later) installed on your computer.

### Check if Node.js is installed
Open a terminal (Command Prompt on Windows, Terminal on Mac/Linux) and run:

```bash
node --version
```

If you see a version number like `v18.x.x` or `v20.x.x`, you're good to go.

### If Node.js is NOT installed

- **Windows / Mac**: Download from https://nodejs.org (choose the LTS version)
- **Linux**: `sudo apt install nodejs npm` (Ubuntu/Debian) or use [nvm](https://github.com/nvm-sh/nvm)

---

## Quick Start (3 steps)

### Step 1 — Open terminal in the project folder

```bash
cd jockey-dashboard
```

### Step 2 — Install dependencies

```bash
npm install
```

This downloads React, Recharts, PapaParse, and Vite. Takes about 30–60 seconds.

### Step 3 — Start the dashboard

```bash
npm run dev
```

This will show output like:

```
  VITE v5.x.x  ready in 300 ms

  ➜  Local:   http://localhost:3000/
  ➜  Network: http://192.168.x.x:3000/
```

**Open http://localhost:3000 in your browser.** The dashboard will open automatically in most cases.

---

## Using the Dashboard

### 1. Upload your data
- Click **"Upload CSV Data"** on the landing page
- Select your survey CSV file (exported from Qualtrics/Excel)
- The dashboard loads instantly — all processing happens in your browser (no data leaves your machine)

### 2. Apply filters
- Use the top filter bar to slice data by **Gender**, **Education**, **Location**, **Licence Type**, and **Age Range**
- All tabs update in real time when filters change
- The sample counter (N = x / total) shows filtered vs total count

### 3. Navigate the tabs

| Tab | What it shows |
|---|---|
| **Overview** | Key stats, CMD prevalence chart, demographics, summary statistics table |
| **CMD Prevalence** | PHQ-9/GAD-7/K10 severity distributions, WEMWBS histogram, CES-D item means, clinical cut-off table |
| **Risk Factors** | SARRS stress categories, Sense of Agency, Career Satisfaction, Sleep distribution |
| **Coping & Support** | Brief-COPE 14 subscales, PASS-Q radar plot (4 support dimensions) |
| **Weight Making** | Prevalence, method usage ranked, cutting frequency, weight stats |
| **Logistic Regression** | Binary CMD outcomes (clinical cut-offs) → ORs, 95% CIs, forest plot |
| **Linear Regression** | Continuous CMD totals → OLS coefficients, R², RMSE |
| **Missing Data** | Variable-level missingness with colour-coded bars |

### 4. Statistical modelling

**Logistic Regression (OR tab):**
1. Select a binary outcome (e.g., CES-D Risk ≥16, PHQ-9 ≥10)
2. Tick predictors from any risk factor group (SARRS, SoA, CSS, PASS-Q, COPE, Sleep, Weight, Demographics)
3. Categorical variables (Gender, Education) are automatically dummy-coded with the first category as reference
4. Click **Run Logistic Regression**
5. Results: β, SE, z, p-value, OR, 95% CI, significance stars, forest plot
6. Export to CSV

**Linear Regression (β tab):**
1. Select a continuous outcome (CES-D, PHQ-9, GAD-7, K10, or WEMWBS total)
2. Tick predictors
3. Click **Run Linear Regression**
4. Results: β, SE, t, p-value, R², Adj R², RMSE, coefficient bar chart
5. Export to CSV

### 5. Export data
Every section has an **"Export CSV"** button for:
- Summary statistics
- CMD prevalence
- Risk factor summaries
- Coping subscales
- Weight-making methods
- Regression results
- Missing data reports

---

## CSV Format Requirements

Your CSV should have column headers matching the variable names from the survey codebook. Key columns include:

**Demographics:** `age`, `gender`, `education_level`, `licence_type`, `primary_race_location`

**CMD Scales:** `cesd_total`, `cesd_depressive_risk`, `phq9_total`, `gad7_total`, `k10_total`, `wemwbs_total`

**Risk Factors:** `sarrs_total_stress_score`, `sarrs_n_events`, `soa_positive_agency_total`, `soa_negative_agency_total`, `css_total`, `passq_total`, `sleep_hours`

**Coping:** `cope_self_distraction_score`, `cope_active_coping_score`, ... (all 14 COPE subscale scores)

**Weight Making:** `ever_cut_weight`, `wm_food_restriction`, `wm_sauna`, ... (all method flags)

The dashboard handles missing values gracefully — missing cells are excluded from calculations per-variable.

---

## Building for Deployment

To create a static build you can host anywhere:

```bash
npm run build
```

The output goes to the `dist/` folder. Upload it to any web server, GitHub Pages, Netlify, or Vercel.

To preview the build locally:

```bash
npm run preview
```

---

## Troubleshooting

| Issue | Fix |
|---|---|
| `npm: command not found` | Install Node.js from https://nodejs.org |
| Port 3000 already in use | Edit `vite.config.js` → change `port: 3000` to another number |
| CSV not loading | Ensure file is UTF-8 encoded CSV with headers in the first row |
| Regression shows "Insufficient cases" | Too many missing values in selected variables — reduce predictors or check data |
| Charts look empty | Check that column names match exactly (case-sensitive) |

---

## Tech Stack

- **React 18** — UI framework
- **Recharts** — Charts and visualizations
- **PapaParse** — CSV parsing
- **Vite** — Build tool and dev server
- All statistical computations (OLS, IRLS logistic regression, matrix inversion) run client-side in JavaScript — no backend needed, no data transmitted

---

*IHRB × DCU School of Health & Human Performance*
*Jockey Mental Wellbeing Research Programme*
