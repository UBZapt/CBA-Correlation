# CBA Correlation – Code Notes

## Task 1: Data Cleaning

### Inputs
| Sheet | Series | Raw date range |
|-------|--------|----------------|
| ARBIX | Absolute Arbitrage Fund monthly returns | Jan 2003 – Feb 2026 |
| Bonds | iShares 7-10 Year Treasury Bond ETF | Aug 2002 – Feb 2026 |
| Equities | MSCI World Index | Feb 1997 – Feb 2026 |

### Process

1. **Load** — `pd.read_excel()` reads each sheet by name into a separate DataFrame.

2. **Column standardisation** — `df.columns.str.strip()` removes trailing whitespace from column names (ARBIX sheet had `'Year '` and `'Return '`).

3. **Date parsing** — Month and Year columns are concatenated as a string (`"January 2003"`) and passed to `pd.to_datetime(..., format='mixed')`.
   `format='mixed'` is required because ARBIX uses full month names (`JANUARY`) while Bonds and Equities use 3-letter abbreviations (`AUG`). `str.title()` normalises case before parsing.

4. **Outer join** — The three series are merged on the `Date` index using `DataFrame.join(..., how='outer')` to expose any gaps across all series simultaneously.

5. **Date range trimming** — The common window is defined as:
   - **Start**: latest of the three series' first dates (`max(starts.values())`)
   - **End**: earliest of the three series' last dates (`min(ends.values())`)
   - Result: **Jan 2003 – Feb 2026** (driven by ARBIX on both ends), giving **278 months**.

6. **Missing month removal** — `df.isnull().any(axis=1)` flags any row where at least one series has `NaN`. Flagged rows are logged and then removed via `df.dropna()`. No months were omitted in this dataset.

### Omitted Data Log
None — all three series had complete data within the Jan 2003 – Feb 2026 window.

> **Note:** The raw ARBIX sheet contained an asterisk on `AUGUST 2017` (`*AUGUST`), which was a source-data footnote marker. This was corrected directly in `Data.xlsx` before processing.

7. **Export** — `clean.index` is reformatted to `"Mon YYYY"` strings via `strftime("%b %Y")` for readability, then written to `Clean_Data.xlsx` using `DataFrame.to_excel()` (requires `openpyxl`).

### Output
- `clean` — a 278-row DataFrame indexed by `Date` with columns `ARBIX`, `Bonds`, `Equities`, containing monthly returns expressed as decimals.
- `Clean_Data.xlsx` — the same data persisted to disk for reuse without re-running the cleaning pipeline.

---

## Task 2: Correlation Analysis & Chart

### Inputs
- `Clean_Data.xlsx` — cleaned monthly returns for `ARBIX`, `Bonds`, `Equities` (Jan 2003 – Feb 2026, 278 months).
- Data is already in monthly return form (decimals); no further transformation applied.

### Periods Analysed
| Label | Start | End | N |
|-------|-------|-----|---|
| Full Sample | Jan 2003 | Feb 2026 | 278 |
| Recession 1 | Dec 2007 | Jun 2009 | 19 |
| Recession 2 | Jan 2020 | Jun 2021 | 18 |
| Boom 1 | Jan 2012 | Dec 2013 | 24 |
| Boom 2 | Jan 2023 | Dec 2025 | 36 |

### Process

1. **Reload** — `pd.read_excel("Clean_Data.xlsx", index_col=0)` loads the cleaned data; the string date index is re-parsed to `datetime` with `pd.to_datetime(..., format="%b %Y")`.

2. **Period subsetting** — each period is sliced with `df.loc[start:end]`. `.dropna()` is applied per slice to ensure only matched (complete) observations are used.

3. **Pearson correlation** — `Series.corr()` computes the Pearson correlation coefficient between `ARBIX` and each of `Equities` and `Bonds` for each period. Results are rounded to 4 decimal places.

4. **Excel export** — `pd.ExcelWriter` in append mode (`mode="a", if_sheet_exists="replace"`) adds a `"Correlations"` sheet to `Clean_Data.xlsx` without disturbing existing sheets.

5. **Chart** — `matplotlib` clustered bar chart:
   - Two bar groups per period: CBA-Equity (navy `#1F3864`) and CBA-Bond (red `#C00000`).
   - Bar width 0.35; bars offset by ±width/2 around integer x-positions.
   - Data labels above/below each bar (`fontweight="bold"`, offset ±0.012).
   - Zero reference line via `ax.axhline(0)`.
   - Gridlines removed (`ax.grid(False)`); top and right spines hidden.
   - Saved as `Correlation_Chart.png` at 300 dpi.

6. **Chart embedding** — `openpyxl.load_workbook` reopens the file; `openpyxl.drawing.image.Image` attaches the PNG to cell `F2` of the `Correlations` sheet.

### Assumptions
- Pearson correlation is appropriate (standard for linear co-movement of return series).
- Sub-period date boundaries are inclusive of both endpoints.
- The recession and boom periods are as specified in the assignment brief; no adjustments were made.
- Data is treated as already representing monthly simple returns; no log-return conversion applied.

### Outputs
- `Clean_Data.xlsx` — updated with a `"Correlations"` sheet containing the results table and the embedded chart.
- `Correlation_Chart.png` — standalone 300 dpi chart file.
