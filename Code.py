import pandas as pd
import numpy as np

# ── 1. Load raw sheets ────────────────────────────────────────────────────────
raw = {
    "ARBIX": pd.read_excel("Data.xlsx", sheet_name="ARBIX"),
    "Bonds": pd.read_excel("Data.xlsx", sheet_name="Bonds"),
    "Equities": pd.read_excel("Data.xlsx", sheet_name="Equities"),
}

# ── 2. Standardise columns and parse dates ────────────────────────────────────
def parse_dates(df):
    df.columns = df.columns.str.strip()
    df["Date"] = pd.to_datetime(
        df["Month"].str.strip().str.title() + " " + df["Year"].astype(str),
        format="mixed"       # handles both full names (JANUARY) and abbreviations (AUG)
    )
    return df[["Date", "Return"]].rename(columns={"Return": df.columns[2]})

series = {}
for name, df in raw.items():
    parsed = parse_dates(df)
    parsed.columns = ["Date", name]
    parsed = parsed.set_index("Date")
    series[name] = parsed

# ── 3. Merge on date index (outer join to expose all gaps) ────────────────────
merged = series["ARBIX"].join(series["Bonds"], how="outer").join(series["Equities"], how="outer")
merged = merged.sort_index()

# ── 4. Trim to common date range (latest start, earliest end) ─────────────────
starts = {name: s.index.min() for name, s in series.items()}
ends   = {name: s.index.max() for name, s in series.items()}

start_date = max(starts.values())
end_date   = min(ends.values())

print(f"Latest start  : {max(starts, key=starts.get)} -> {start_date.strftime('%b %Y')}")
print(f"Earliest end  : {min(ends,   key=ends.get)}   -> {end_date.strftime('%b %Y')}")

merged = merged.loc[start_date:end_date]

# ── 5. Identify and drop months with any missing data ─────────────────────────
missing_mask = merged.isnull().any(axis=1)
omitted = merged[missing_mask].copy()

if not omitted.empty:
    print(f"\nOmitted months ({len(omitted)}):")
    for dt, row in omitted.iterrows():
        missing_cols = row.index[row.isnull()].tolist()
        print(f"  {dt.strftime('%b %Y')} — missing in: {', '.join(missing_cols)}")
else:
    print("\nNo missing months within the common date range.")

clean = merged.dropna()

print(f"\nClean dataset : {clean.index[0].strftime('%b %Y')} -> {clean.index[-1].strftime('%b %Y')}  ({len(clean)} months)")
print(clean.head())
