# CBA Correlation – Code Notes

## Configuration

All runtime parameters are declared at the top of `Code.py`:

| Variable | Value | Purpose |
|----------|-------|---------|
| `INPUT_FILE` | `Data.xlsx` | Source workbook |
| `OUTPUT_FILE` | `Clean_Data.xlsx` | Final output workbook |
| `SHEETS` | `{"ARBIX", "Bonds", "Equities"}` | Sheet name mapping |
| `MIN_OBS` | `12` | Minimum observations required for a valid correlation |
| `ROLLING_WINDOW` | `12` | Rolling window length (months) for rolling correlation charts |
| `PERIODS` | Dict of sub-period labels → (start, end) timestamps | Subperiods for correlation analysis |

The full-sample period label is generated dynamically from the actual cleaned data range and prepended to `PERIODS` at runtime.

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

Note: the original redundant `.rename()` inside the parsing step (which renamed `Return` to `df.columns[2]` only for it to be overwritten immediately) has been removed.

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
5. Returns a `corr_df` summary table and the complete `all_periods` dict (used for chart x-labels).

### `plot_correlation_bars(corr_df, all_periods)`
Clustered bar chart of CBA-Equity and CBA-Bond correlations across all periods. Saves `Correlation_Chart.png`.

### `plot_rolling_correlations(clean)`
Line chart of rolling `ROLLING_WINDOW`-month Pearson correlations for ARBIX vs Equities and ARBIX vs Bonds. Saves `Rolling_Correlation_Chart.png`.

### `plot_scatter(clean, other, color, filename)`
Scatter plot of ARBIX monthly returns against a specified series (`Equities` or `Bonds`), with an OLS trend line and correlation coefficient in the legend. Saves to the given filename.

### `plot_volatility(clean)`
Bar chart of monthly return standard deviations for ARBIX, Equities, and Bonds. Saves `Volatility_Chart.png`.

### `plot_down_market(clean)`
Bar chart comparing average ARBIX monthly return across all months versus months when Equities had a negative return. Saves `Down_Market_Chart.png`.

### `export_outputs(clean, corr_df)`
Single export step: converts the datetime index to `"Mon YYYY"` strings, then writes both `Clean_Data` and `Correlations` sheets to `OUTPUT_FILE` in one pass. No intermediate file is written during processing.

---

## Processing Steps

1. Load raw sheets from `Data.xlsx`.
2. Validate required columns; parse Month + Year into a datetime index per series.
3. Outer-join → trim to common window → drop incomplete rows → produce `clean` DataFrame.
4. Keep `clean` in memory; pass it directly to correlation and chart steps (no write/re-read of `Clean_Data.xlsx` mid-workflow).
5. Compute correlations for the dynamic full-sample label and all configured sub-periods.
6. Generate six charts (see Outputs).
7. Write `Clean_Data.xlsx` once, at the end.

---

## Assumptions

- Pearson correlation is appropriate (linear co-movement of return series).
- Sub-period date boundaries are inclusive of both endpoints.
- Recession and boom periods match the assignment brief; no adjustments made.
- Data represents monthly simple returns (decimals); no log-return conversion applied.
- `format="mixed"` date parsing is required because ARBIX uses full month names (`JANUARY`) while Bonds and Equities use three-letter abbreviations (`AUG`); `str.title()` normalises case before parsing.
- Subperiods with fewer than `MIN_OBS` (12) observations produce `NaN` correlation rather than a potentially unreliable estimate.

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
| `Correlation_Chart.png` | Clustered bar chart: CBA-Equity and CBA-Bond correlations by period |
| `Rolling_Correlation_Chart.png` | 12-month rolling correlation lines: ARBIX vs Equities and ARBIX vs Bonds |
| `Scatter_Equities.png` | Scatter of ARBIX vs Equities monthly returns with OLS trend line |
| `Scatter_Bonds.png` | Scatter of ARBIX vs Bonds monthly returns with OLS trend line |
| `Volatility_Chart.png` | Bar chart of monthly return standard deviations (ARBIX, Equities, Bonds) |
| `Down_Market_Chart.png` | Bar chart: average ARBIX return — all months vs equity down months |
