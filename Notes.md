# CBA Correlation – Code Notes

## Configuration

All runtime parameters are declared at the top of `Code.py`:

| Variable | Value | Purpose |
|----------|-------|---------|
| `INPUT_FILE` | `Data.xlsx` | Source workbook |
| `OUTPUT_FILE` | `Clean_Data.xlsx` | Final output workbook |
| `SHEETS` | `{"ARBIX", "Bonds", "Equities"}` | Sheet name mapping |
| `MIN_OBS` | `11` | Minimum observations required for a valid correlation (lowered from 12 to accommodate the 11-month Covid Stress window Feb–Dec 2020) |
| `PERIODS` | Dict of sub-period labels → (start, end) timestamps | Subperiods for correlation analysis |

The full-sample period label is generated dynamically from the actual cleaned data range and prepended to `PERIODS` at runtime.

### Sub-Period Definitions

| Period Label | Type | Date Range |
|---|---|---|
| GFC | Period of Stress | Dec 2007 – Jun 2009 |
| Post-GFC Expansion | Period of Growth | Jan 2012 – Dec 2013 |
| Covid Stress | Period of Stress | Feb 2020 – Dec 2020 |
| CB Renaissance | Period of Growth | Jan 2023 – Dec 2025 |

Covid Stress starts Feb 2020 (first month of market stress) and ends Dec 2020 (pre-vaccine rollout normalisation). The 11-month window is intentional and drives the `MIN_OBS = 11` setting.

`_PERIOD_TYPES` and the colour constants `_STRESS_COLOR` / `_GROWTH_COLOR` are module-level helpers used by `plot_correlation_bars` to apply background shading.

### Colour Palette

Three teal shades matched to the reference Combo Chart and Clustered Bar Chart inputs:

| Constant | Hex | Role |
|----------|-----|------|
| `TEAL_DARK` | `#1A5F5F` | Primary series (ARBIX–Equity, All Months, ARBIX volatility) |
| `TEAL_MID` | `#2E9E8A` | Third series (Bonds in volatility chart) |
| `TEAL_LIGHT` | `#93D4CB` | Secondary series (ARBIX–Bond, Equity Down Months, Equities volatility) |

---

## Functions

### `load_raw_data()`
Reads each sheet listed in `SHEETS` from `INPUT_FILE` into a dict of raw DataFrames. No transformation applied.

### `validate_series(name, df)`
Checks that the DataFrame contains the three required columns — `Month`, `Year`, `Return` (after stripping whitespace). Raises `ValueError` if any are absent.

### `parse_series(name, df)`
1. Strips whitespace from column names.
2. Calls `validate_series`.
3. Constructs a `Date` column by concatenating `Month` (title-cased, stripped) and `Year`, parsed with `pd.to_datetime(..., format="mixed")`.
4. Checks for duplicate dates; raises `ValueError` if any found.
5. Returns a DataFrame with a `datetime` index and a single column named after the series.

### `build_clean_dataset(raw)`
1. Calls `parse_series` for each raw DataFrame.
2. Outer-joins all three series on the date index to expose any gaps.
3. Trims to the common date window: latest start and earliest end across all series.
4. Logs any months with missing values, then drops them.
5. Returns `clean` — a DataFrame with a **datetime** index and columns `ARBIX`, `Bonds`, `Equities`.

Dates remain as `datetime` objects throughout; string formatting happens only at export.

### `calculate_correlations(clean)`
1. Builds the full-sample label dynamically from `clean.index.min()` and `clean.index.max()`.
2. Prepends it to `PERIODS`.
3. For each period, slices `clean` and checks observation count against `MIN_OBS`. If below threshold, returns `NaN` and prints a warning.
4. Computes Pearson correlation (rounded to 4 d.p.) between `ARBIX` and each of `Equities` and `Bonds`.
5. Returns a `corr_df` summary table and the complete `all_periods` dict (used for chart y-labels).

### `_circle_legend(ax, handles, loc, ncol, bbox_to_anchor)`
Helper that renders a legend using filled-circle markers to match the style of the reference Combo Chart (Input 1). Called by all three chart functions. Parameters: `ax` (axes), `handles` (list of `Line2D` with `marker="o"`), `loc` (legend anchor string), `ncol` (number of columns in legend), `bbox_to_anchor` (optional tuple to anchor the legend outside the axes area — used when the legend must not overlap chart content).

### `plot_correlation_bars(corr_df, all_periods)`
Vertical clustered bar chart of ARBIX–Equity and ARBIX–Bond correlations across all periods. Bars stand upright and are grouped along a horizontal x-axis (one cluster per period).

**Key design choices:**
- **Vertical bar layout**: `ax.bar()` with two bars per x-position (`x ± width/2`). ARBIX–Equity bar uses `TEAL_DARK` (left of cluster); ARBIX–Bond bar uses `TEAL_LIGHT` (right of cluster).
- **Shaded vertical column bands** applied per sub-period using `ax.axvspan()`: soft red (`_STRESS_COLOR`, α = 0.08) for Periods of Stress (GFC, Covid Stress); soft green (`_GROWTH_COLOR`, α = 0.08) for Periods of Growth (Post-GFC Expansion, CB Renaissance). Position 0 (Full Sample) is unshaded.
- **Legend labels**: "ARBIX – MSCI World Index" (formerly CBA-Equity) and "ARBIX – iShares Treasury Bond ETF" (formerly CBA-Bond), plus stress/growth shading swatches.
- **Circle legend** floated above the chart area using `bbox_to_anchor=(0.5, 1.02)` with `loc="lower center"` and `ncol=4` (all four items in one row). This ensures the legend never overlaps bars regardless of bar height.
- **Y-axis label**: "Correlation Coefficients".
- No chart title.

Saves `Correlation_Chart.png`.

### `plot_volatility(clean)`
Horizontal bar chart of **annualised** return standard deviations for ARBIX, Equities, and Bonds. Bars run left-to-right with ARBIX at the top, Equities in the middle, and Bonds at the bottom.

**Processing step:** `monthly_std × √12` converts monthly standard deviation to an annualised figure. Values are multiplied by 100 and displayed as percentage points (e.g., `"5.24%"`). X-axis is labelled "Annualised Volatility (%)".

**Layout:** `ax.barh()` plots bottom-to-top, so the data lists are ordered `[Bonds, Equities, ARBIX]` to produce the desired top-to-bottom visual order (ARBIX → Equities → Bonds). Data labels are placed to the right of each bar end.

**Legend:** Stacked vertically (`ncol=1`) in the order ARBIX → Equities → Bonds, anchored to the right of the axes via `bbox_to_anchor=(1.02, 0.5)` with `loc="center left"` so it never overlaps the bars. No chart title.

Saves `Volatility_Chart.png`.

### `plot_down_market(clean)`
Vertical bar chart comparing average ARBIX monthly return across all months versus months when Equities had a negative return.

**Layout:** Bars positioned at `x = [0.42, 0.84]` (width = 0.32) with `xlim = [0.05, 1.25]`, reducing the centre gap and moving bars away from the y-axis to prevent x-tick label overlap.  Data-label y-offset set dynamically as `2% of yrange`.

**Visual changes:** `TEAL_DARK` for "All Months", `TEAL_LIGHT` for "Equity Down Months". Circle legend at upper right with two columns (Input 1 style). No chart title.

Saves `Down_Market_Chart.png`.

### `calculate_downside_stats(clean)`
Computes two downside-risk statistics using only the equity-down months (months where `Equities < 0`):

1. **Downside Capture Ratio** — compounded ARBIX return over equity-down months divided by the compounded Equities return over the same months:
   - `arbix_compound = ∏(1 + ARBIX_i) − 1` for all equity-down months
   - `equity_compound = ∏(1 + Equities_i) − 1` for all equity-down months
   - `DCR = arbix_compound / equity_compound`
2. **Down-Market Beta** — OLS beta restricted to equity-down months:
   - `β = Cov(ARBIX, Equities) / Var(Equities)` over equity-down months

Returns a two-row DataFrame with columns `Statistic`, `Value`, and `N (down months)`. Values are rounded to 4 d.p.

### `export_outputs(clean, corr_df, downside_df)`
Single export step: converts the datetime index to `"Mon YYYY"` strings, then writes three sheets to `OUTPUT_FILE` in one pass — `Clean_Data`, `Correlations`, and `Downside_Stats`. No intermediate file is written during processing.

---

## Processing Steps

1. Load raw sheets from `Data.xlsx`.
2. Validate required columns; parse Month + Year into a datetime index per series.
3. Outer-join → trim to common window → drop incomplete rows → produce `clean` DataFrame.
4. Keep `clean` in memory; pass it directly to correlation and chart steps (no write/re-read of `Clean_Data.xlsx` mid-workflow).
5. Compute correlations for the dynamic full-sample label and all configured sub-periods.
6. Compute downside capture ratio and down-market beta over equity-down months.
7. Generate three charts: `Correlation_Chart.png`, `Volatility_Chart.png`, `Down_Market_Chart.png`.
8. Write `Clean_Data.xlsx` once, at the end (three sheets: `Clean_Data`, `Correlations`, `Downside_Stats`).

---

## Assumptions

- Pearson correlation is appropriate (linear co-movement of return series).
- Equity-down months are defined as months where the Equities return is strictly negative (`< 0`); zero-return months are excluded from this filter.
- Downside capture ratio uses geometric (compounded) returns, not arithmetic averages. Both the numerator and denominator will be negative (ARBIX earns less than the benchmark in down markets), so a ratio below 1.0 indicates ARBIX loses proportionally less than equities during drawdowns.
- Down-market beta uses pandas `.cov()` and `.var()` (both default to `ddof=1`, sample statistics), consistent with the rest of the analysis.
- Sub-period date boundaries are inclusive of both endpoints.
- Named periods are defined as follows: GFC = peak-to-trough of the Global Financial Crisis; Post-GFC Expansion = QE-driven equity recovery period; Covid Stress = initial market dislocation through peak pandemic uncertainty (Feb–Dec 2020); CB Renaissance = renewed CBA outperformance in the post-rate-hike environment.
- Volatility is annualised as `monthly std × √12`, assuming i.i.d. monthly returns. This is the standard square-root-of-time scaling convention.
- Data represents monthly simple returns (decimals); no log-return conversion applied.
- `format="mixed"` date parsing is required because ARBIX uses full month names (`JANUARY`) while Bonds and Equities use three-letter abbreviations (`AUG`); `str.title()` normalises case before parsing.
- Subperiods with fewer than `MIN_OBS` (11) observations produce `NaN` correlation rather than a potentially unreliable estimate.
- The Correlation chart uses vertical bars (value axis = y-axis), so "Correlation Coefficients" is applied as the y-axis label.
- The Volatility chart uses horizontal bars (value axis = x-axis), so "Annualised Volatility (%)" is applied as the x-axis label.
- Scatter plots and the 12-month rolling correlation chart have been removed as they are no longer required by the analysis.

---

## Inputs

| Sheet | Series | Raw date range |
|-------|--------|----------------|
| ARBIX | Absolute Arbitrage Fund monthly returns | Jan 2003 – Feb 2026 |
| Bonds | iShares 7-10 Year Treasury Bond ETF | Aug 2002 – Feb 2026 |
| Equities | MSCI World Index | Feb 1997 – Feb 2026 |

Common window after trimming: **Jan 2003 – Feb 2026** (driven by ARBIX on both ends).

> **Note:** The raw ARBIX sheet contained an asterisk on `AUGUST 2017` (`*AUGUST`), which was a source-data footnote marker. Corrected directly in `Data.xlsx` before processing.

---

## Outputs

| File | Content |
|------|---------|
| `Clean_Data.xlsx` — sheet `Clean_Data` | 278-row monthly returns table (ARBIX, Bonds, Equities) with `"Mon YYYY"` string index |
| `Clean_Data.xlsx` — sheet `Correlations` | Pearson correlation table by period |
| `Clean_Data.xlsx` — sheet `Downside_Stats` | Downside capture ratio and down-market beta vs Equities, with observation count |
| `Correlation_Chart.png` | Horizontal clustered bar chart: ARBIX–Equity and ARBIX–Bond correlations by period; stress/growth periods highlighted with horizontal background shading |
| `Volatility_Chart.png` | Vertical bar chart of annualised return standard deviations (ARBIX, Equities, Bonds), displayed as % |
| `Down_Market_Chart.png` | Vertical bar chart: average ARBIX return — all months vs equity down months |
