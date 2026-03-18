import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from scipy import stats

# ── Configuration ─────────────────────────────────────────────────────────────
INPUT_FILE     = "Data.xlsx"
OUTPUT_FILE    = "Clean_Data.xlsx"
SHEETS         = {"ARBIX": "ARBIX", "Bonds": "Bonds", "Equities": "Equities"}
MIN_OBS        = 11
ROLLING_WINDOW = 12

PERIODS = {
    "GFC\n(Dec 2007 – Jun 2009)":                (pd.Timestamp("2007-12-01"), pd.Timestamp("2009-06-01")),
    "Post-GFC Expansion\n(Jan 2012 – Dec 2013)":  (pd.Timestamp("2012-01-01"), pd.Timestamp("2013-12-01")),
    "Covid Stress\n(Feb 2020 – Dec 2020)":         (pd.Timestamp("2020-02-01"), pd.Timestamp("2020-12-01")),
    "CB Renaissance\n(Jan 2023 – Dec 2025)":       (pd.Timestamp("2023-01-01"), pd.Timestamp("2025-12-01")),
}

# Period type for shading (index in all_periods after full-sample at index 0)
# 1=GFC (stress), 2=Post-GFC Expansion (growth), 3=Covid Stress (stress), 4=CB Renaissance (growth)
_PERIOD_TYPES = {1: "stress", 2: "growth", 3: "stress", 4: "growth"}
_STRESS_COLOR = "#B22222"
_GROWTH_COLOR = "#2E7D32"


# ── Functions ─────────────────────────────────────────────────────────────────

def load_raw_data():
    return {name: pd.read_excel(INPUT_FILE, sheet_name=sheet) for name, sheet in SHEETS.items()}


def validate_series(name, df):
    required = {"Month", "Year", "Return"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Sheet '{name}' is missing columns: {missing}")


def parse_series(name, df):
    df.columns = df.columns.str.strip()
    validate_series(name, df)
    df["Date"] = pd.to_datetime(
        df["Month"].str.strip().str.title() + " " + df["Year"].astype(str),
        format="mixed",
    )
    dupes = df["Date"].duplicated()
    if dupes.any():
        raise ValueError(f"Sheet '{name}' has duplicate dates: {df.loc[dupes, 'Date'].tolist()}")
    result = df[["Date", "Return"]].copy()
    result.columns = ["Date", name]
    return result.set_index("Date")


def build_clean_dataset(raw):
    series = {name: parse_series(name, df) for name, df in raw.items()}

    merged = (
        series["ARBIX"]
        .join(series["Bonds"],    how="outer")
        .join(series["Equities"], how="outer")
        .sort_index()
    )

    start_date = max(s.index.min() for s in series.values())
    end_date   = min(s.index.max() for s in series.values())
    merged = merged.loc[start_date:end_date]

    omitted = merged[merged.isnull().any(axis=1)]
    if not omitted.empty:
        print(f"Omitted months ({len(omitted)}):")
        for dt, row in omitted.iterrows():
            cols = row.index[row.isnull()].tolist()
            print(f"  {dt.strftime('%b %Y')} — missing: {', '.join(cols)}")
    else:
        print("No missing months within the common date range.")

    clean = merged.dropna()
    print(f"Clean dataset: {clean.index[0].strftime('%b %Y')} – {clean.index[-1].strftime('%b %Y')} ({len(clean)} months)")
    return clean


def calculate_correlations(clean):
    full_label = (
        f"Full Sample\n({clean.index[0].strftime('%b %Y')} – {clean.index[-1].strftime('%b %Y')})"
    )
    all_periods = {full_label: (clean.index.min(), clean.index.max()), **PERIODS}

    records = []
    for label, (start, end) in all_periods.items():
        sub = clean.loc[start:end].dropna(subset=["ARBIX", "Equities", "Bonds"])
        n = len(sub)
        if n < MIN_OBS:
            eq_corr = bd_corr = np.nan
            print(f"Warning: '{label.replace(chr(10), ' ')}' has only {n} observations (< {MIN_OBS}); returning NaN.")
        else:
            eq_corr = round(sub["ARBIX"].corr(sub["Equities"]), 4)
            bd_corr = round(sub["ARBIX"].corr(sub["Bonds"]),    4)
        records.append({
            "Period":     label.replace("\n", " "),
            "CBA-Equity": eq_corr,
            "CBA-Bond":   bd_corr,
            "N (months)": n,
        })

    return pd.DataFrame(records), all_periods


def plot_correlation_bars(corr_df, all_periods):
    chart_labels = list(all_periods.keys())
    x     = np.arange(len(chart_labels))
    width = 0.35

    fig, ax = plt.subplots(figsize=(14, 7))

    bars_eq = ax.bar(x - width / 2, corr_df["CBA-Equity"], width,
                     label="CBA–Equity", color="#1F3864")
    bars_bd = ax.bar(x + width / 2, corr_df["CBA-Bond"],   width,
                     label="CBA–Bond",   color="#C00000")

    for bar in list(bars_eq) + list(bars_bd):
        h = bar.get_height()
        if np.isnan(h):
            continue
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h + (0.012 if h >= 0 else -0.012),
            f"{h:.3f}",
            ha="center", va="bottom" if h >= 0 else "top",
            fontsize=8.5, fontweight="bold",
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

    # Colour-code x-tick labels by period type — no background fill to avoid bar colour mixing
    tick_colors = {0: "black", 1: _STRESS_COLOR, 2: _GROWTH_COLOR,
                   3: _STRESS_COLOR, 4: _GROWTH_COLOR}
    for i, tick in enumerate(ax.get_xticklabels()):
        tick.set_color(tick_colors.get(i, "black"))

    from matplotlib.lines import Line2D
    legend_elements = [
        mpatches.Patch(facecolor="#1F3864", label="CBA–Equity"),
        mpatches.Patch(facecolor="#C00000", label="CBA–Bond"),
        Line2D([0], [0], marker="s", color="w", markerfacecolor=_STRESS_COLOR,
               markersize=10, label="Period of Stress"),
        Line2D([0], [0], marker="s", color="w", markerfacecolor=_GROWTH_COLOR,
               markersize=10, label="Period of Growth"),
    ]
    ax.legend(handles=legend_elements, fontsize=9.5, frameon=False, ncol=2, loc="upper right")

    ax.grid(False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    plt.savefig("Correlation_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def plot_rolling_correlations(clean):
    roll_eq = clean["ARBIX"].rolling(ROLLING_WINDOW).corr(clean["Equities"])
    roll_bd = clean["ARBIX"].rolling(ROLLING_WINDOW).corr(clean["Bonds"])

    fig, ax = plt.subplots(figsize=(13, 5))
    ax.plot(roll_eq.index, roll_eq, label=f"ARBIX vs Equities ({ROLLING_WINDOW}m rolling)", color="#1F3864", linewidth=1.5)
    ax.plot(roll_bd.index, roll_bd, label=f"ARBIX vs Bonds ({ROLLING_WINDOW}m rolling)",    color="#C00000", linewidth=1.5)
    ax.axhline(0, color="black", linewidth=0.8, linestyle="--")
    ax.set_xlabel("Date", fontsize=11)
    ax.set_ylabel("Pearson Correlation", fontsize=11)
    ax.set_title(
        f"Rolling {ROLLING_WINDOW}-Month Correlation: ARBIX vs Equities and Bonds",
        fontsize=12, fontweight="bold", pad=12,
    )
    ax.legend(fontsize=10, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    plt.savefig("Rolling_Correlation_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def plot_scatter(clean, other, color, filename):
    x = clean[other]
    y = clean["ARBIX"]
    slope, intercept, r, _, _ = stats.linregress(x, y)
    x_line = np.linspace(x.min(), x.max(), 200)

    fig, ax = plt.subplots(figsize=(7, 6))
    ax.scatter(x, y, color=color, alpha=0.5, s=20)
    ax.plot(x_line, slope * x_line + intercept, color="black", linewidth=1.2, label=f"r = {r:.3f}")
    ax.set_xlabel(f"{other} Monthly Return", fontsize=11)
    ax.set_ylabel("ARBIX Monthly Return", fontsize=11)
    ax.set_title(f"ARBIX vs {other}: Monthly Returns", fontsize=12, fontweight="bold", pad=12)
    ax.legend(fontsize=10, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, dpi=300, bbox_inches="tight")
    plt.close()


def plot_volatility(clean):
    # Annualise: monthly std × √12
    vols   = clean[["ARBIX", "Equities", "Bonds"]].std() * np.sqrt(12)
    vols_pct = vols * 100          # express as percentage points
    colors = ["#1F3864", "#2E75B6", "#C00000"]

    fig, ax = plt.subplots(figsize=(7, 5))
    bars = ax.bar(vols_pct.index, vols_pct.values, color=colors)
    for bar, v in zip(bars, vols_pct.values):
        ax.text(bar.get_x() + bar.get_width() / 2, v + 0.05, f"{v:.2f}%",
                ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_ylabel("Annualised Volatility (%)", fontsize=11)
    ax.set_title("Volatility Comparison: Annualised Return Standard Deviations",
                 fontsize=12, fontweight="bold", pad=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)
    plt.tight_layout()
    plt.savefig("Volatility_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def plot_down_market(clean):
    down = clean[clean["Equities"] < 0]
    values = [clean["ARBIX"].mean(), down["ARBIX"].mean()]
    labels = ["All Months", f"Equity Down Months (n={len(down)})"]
    colors = ["#1F3864", "#C00000"]

    fig, ax = plt.subplots(figsize=(6, 5))

    x = [0.42, 0.84]
    width = 0.32
    bars = ax.bar(x, values, color=colors, width=width)

    # Generous y-axis range: 40% headroom above max and below min so labels
    # never crowd the title or overlap the zero line
    ymax = max(values)
    ymin = min(values)
    yspan = ymax - ymin if ymax != ymin else abs(ymax) or 0.01
    ax.set_ylim(ymin - yspan * 0.4, ymax + yspan * 0.4)

    # Small fixed offset: 2% of the visible y-span keeps labels tight to bars
    offset = yspan * 0.02

    for bar, v in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            v + (offset if v >= 0 else -offset),
            f"{v:.4f}",
            ha="center", va="bottom" if v >= 0 else "top",
            fontsize=10, fontweight="bold",
        )

    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9.5)
    ax.set_xlim(0.05, 1.25)
    ax.set_ylabel("Avg ARBIX Monthly Return", fontsize=11)
    ax.set_title("ARBIX Performance in Equity Down Months", fontsize=12, fontweight="bold", pad=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)
    plt.tight_layout()
    plt.savefig("Down_Market_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def calculate_downside_stats(clean):
    down = clean[clean["Equities"] < 0]

    # Downside capture ratio: compounded ARBIX return / compounded Equity return, equity-down months only
    arbix_compound  = (1 + down["ARBIX"]).prod() - 1
    equity_compound = (1 + down["Equities"]).prod() - 1
    dcr = round((arbix_compound / equity_compound)*100, 4)

    # Down-market beta: Cov(ARBIX, Equities) / Var(Equities), equity-down months only
    down_beta = round(down["ARBIX"].cov(down["Equities"]) / down["Equities"].var(), 4)

    return pd.DataFrame([
        {"Statistic": "Downside Capture Ratio (ARBIX vs Equities)", "Value": dcr,      "N (down months)": len(down)},
        {"Statistic": "Down-Market Beta (ARBIX vs Equities)",        "Value": down_beta, "N (down months)": len(down)},
    ])


def export_outputs(clean, corr_df, downside_df):
    export = clean.copy()
    export.index = export.index.strftime("%b %Y")
    export.index.name = "Date"

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name="Clean_Data")
        corr_df.to_excel(writer, sheet_name="Correlations", index=False)
        downside_df.to_excel(writer, sheet_name="Downside_Stats", index=False)


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    raw              = load_raw_data()
    clean            = build_clean_dataset(raw)
    corr_df, periods = calculate_correlations(clean)
    downside_df      = calculate_downside_stats(clean)

    plot_correlation_bars(corr_df, periods)
    plot_rolling_correlations(clean)
    plot_scatter(clean, "Equities", "#1F3864", "Scatter_Equities.png")
    plot_scatter(clean, "Bonds",    "#C00000", "Scatter_Bonds.png")
    plot_volatility(clean)
    plot_down_market(clean)

    export_outputs(clean, corr_df, downside_df)
    print("Done.")
