import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

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

# ── 6. Export cleaned data ────────────────────────────────────────────────────
clean.index = clean.index.strftime("%b %Y")
clean.index.name = "Date"
clean.to_excel("Clean_Data.xlsx")

# ═════════════════════════════════════════════════════════════════════════════
# TASK 2: CORRELATION ANALYSIS
# ═════════════════════════════════════════════════════════════════════════════

# ── 7. Load clean data ────────────────────────────────────────────────────────
df = pd.read_excel("Clean_Data.xlsx", index_col=0)
df.index = pd.to_datetime(df.index, format="%b %Y")

# Data is already in monthly return form (decimals) — no transformation needed.

# ── 8. Define periods ─────────────────────────────────────────────────────────
periods = {
    "Full Sample\n(Jan 2003 - Feb 2026)":     (pd.Timestamp("2003-01-01"), pd.Timestamp("2026-02-01")),
    "Recession 1\n(Dec 2007 - Jun 2009)":     (pd.Timestamp("2007-12-01"), pd.Timestamp("2009-06-01")),
    "Recession 2\n(Jan 2020 - Jun 2021)":     (pd.Timestamp("2020-01-01"), pd.Timestamp("2021-06-01")),
    "Boom 1\n(Jan 2012 - Dec 2013)":          (pd.Timestamp("2012-01-01"), pd.Timestamp("2013-12-01")),
    "Boom 2\n(Jan 2023 - Dec 2025)":          (pd.Timestamp("2023-01-01"), pd.Timestamp("2025-12-01")),
}

# ── 9. Calculate correlations ─────────────────────────────────────────────────
records = []
for label, (start, end) in periods.items():
    sub = df.loc[start:end].dropna(subset=["ARBIX", "Equities", "Bonds"])
    records.append({
        "Period":      label.replace("\n", " "),
        "CBA-Equity":  round(sub["ARBIX"].corr(sub["Equities"]), 4),
        "CBA-Bond":    round(sub["ARBIX"].corr(sub["Bonds"]),    4),
        "N (months)":  len(sub),
    })

corr_df = pd.DataFrame(records)
print(corr_df.to_string(index=False))

# ── 10. Write correlations to new Excel sheet ──────────────────────────────────
with pd.ExcelWriter("Clean_Data.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    corr_df.to_excel(writer, sheet_name="Correlations", index=False)

# ── 11. Clustered bar chart ───────────────────────────────────────────────────
chart_labels = list(periods.keys())          # keep \n for multi-line x-ticks
x     = np.arange(len(chart_labels))
width = 0.35

fig, ax = plt.subplots(figsize=(13, 6))

bars_eq = ax.bar(x - width / 2, corr_df["CBA-Equity"], width,
                 label="CBA-Equity", color="#1F3864")
bars_bd = ax.bar(x + width / 2, corr_df["CBA-Bond"],   width,
                 label="CBA-Bond",   color="#C00000")

# Data labels
for bar in list(bars_eq) + list(bars_bd):
    h = bar.get_height()
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        h + (0.012 if h >= 0 else -0.012),
        f"{h:.3f}",
        ha="center",
        va="bottom" if h >= 0 else "top",
        fontsize=8.5,
        fontweight="bold",
    )

ax.axhline(0, color="black", linewidth=0.8)
ax.set_xticks(x)
ax.set_xticklabels(chart_labels, fontsize=9)
ax.set_xlabel("Period", fontsize=11, labelpad=8)
ax.set_ylabel("Pearson Correlation Coefficient", fontsize=11)
ax.set_title(
    "Convertible Bond Arbitrage Fund: Correlations with Equity and Bond Indices\n"
    "Full Sample and Economic Sub-Periods",
    fontsize=12, fontweight="bold", pad=14,
)
ax.legend(fontsize=10, frameon=False)
ax.grid(False)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

plt.tight_layout()
plt.savefig("Correlation_Chart.png", dpi=300, bbox_inches="tight")
plt.close()
print("Chart saved: Correlation_Chart.png")
