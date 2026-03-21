# CBA-Correlation

This code makes up the analysis part of a project aimed at pitching a Convertible Bond Arbitrage (CBA) fund to a pension fund with long-only positions in equities and bonds. We use monthly return data. The goal is to demonstrate how well the strategy diversifies a portfolio and holds up during adverse market conditions (hedge efficiency).

---

## Data Sources

- **ARBIX** — Monthly return data for the ARBIX fund, sourced from Absolute Investment Advisers
- **Equities** — MSCI World Index (broad global equity proxy)
- **Bonds** — iShares 7–10 Year Treasury Bond ETF (chosen as a middle ground between cash-like behaviour and interest rate sensitivity)

---

## Graphs

### Figure 1 — Correlation Chart (`Correlation_Chart.png`)
A clustered bar chart showing the Pearson correlation between ARBIX and each of the two asset class proxies (equities and bonds). Correlations are calculated across:
- A **full sample** (entire available date range)
- Four **sub-periods** to test behaviour across different market regimes:
  - GFC (Dec 2007 – Jun 2009) — *stress*
  - Post-GFC Expansion (Jan 2012 – Dec 2013) — *growth*
  - Covid Stress (Feb 2020 – Dec 2020) — *stress*
  - CB Renaissance (Jan 2023 – Dec 2025) — *growth*

Stress periods are shaded red; growth periods are shaded green. A correlation near zero indicates the strategy moves independently of that asset class.

### Figure 2 — Volatility Chart (`Volatility_Chart.png`)
A horizontal bar chart comparing the **annualised return volatility** (standard deviation) of ARBIX, equities, and bonds. Lower volatility relative to equities indicates the strategy is less prone to large swings in value, which is useful for portfolio diversification.

### Figure 3 — Down Market Chart (`Down_Market_Chart.png`)
A vertical bar chart comparing ARBIX's **average monthly return across all months** versus only **equity down months** (months where the equity index fell). This tests how well the strategy holds up when equity markets are falling.

---

## Calculated Ratios

| Statistic | Value | Based On |
|---|---|---|
| Downside Capture Ratio (ARBIX vs Equities) | 2.88% | 104 equity down months |
| Down-Market Beta (ARBIX vs Equities) | 0.1827 | 104 equity down months |

- **Downside Capture Ratio** — measures how much of the equity benchmark's decline ARBIX captures during down months. A value closer to 0% is better, indicating the fund loses far less than equities when markets fall.
- **Down-Market Beta** — a version of standard Beta calculated only during equity down months. A value below 1 (and especially near 0) indicates low sensitivity to adverse equity movements.

---

## Data Validation

The code automatically checks for the following before running any analysis:
- All required columns (`Month`, `Year`, `Return`) are present in each data sheet
- No duplicate dates exist within a series
- The dataset is trimmed to the **common date range** shared across all three series
- Any months with missing values are flagged and excluded, with a count printed to the console
- Sub-periods with fewer than 11 observations are flagged and excluded from the correlation results

---

## Outputs

Running `Code.py` produces the following files:
- `Correlation_Chart.png`
- `Volatility_Chart.png`
- `Down_Market_Chart.png`
- `Clean_Data.xlsx` — contains three sheets: cleaned return data, correlation results, and downside statistics
