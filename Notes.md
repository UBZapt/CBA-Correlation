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
