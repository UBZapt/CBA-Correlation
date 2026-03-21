import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

# ── Configuration ─────────────────────────────────────────────────────────────
INPUT_FILE  = "Data.xlsx"
OUTPUT_FILE = "Clean_Data.xlsx"
SHEETS      = {"ARBIX": "ARBIX", "Bonds": "Bonds", "Equities": "Equities"}
MIN_OBS     = 11

PERIODS = {
    "GFC\n(Dec 2007 – Jun 2009)":                ("stress", pd.Timestamp("2007-12-01"), pd.Timestamp("2009-06-01")),
    "Post-GFC Expansion\n(Jan 2012 – Dec 2013)":  ("growth", pd.Timestamp("2012-01-01"), pd.Timestamp("2013-12-01")),
    "Covid Stress\n(Feb 2020 – Dec 2020)":         ("stress", pd.Timestamp("2020-02-01"), pd.Timestamp("2020-12-01")),
    "CB Renaissance\n(Jan 2023 – Dec 2025)":       ("growth", pd.Timestamp("2023-01-01"), pd.Timestamp("2025-12-01")),
}
_STRESS_COLOR = "#B22222"
_GROWTH_COLOR = "#2E7D32"

# Teal colour palette matched to reference Combo / Clustered Bar Chart inputs
TEAL_DARK  = "#1A5F5F"   # dark teal  — primary series
TEAL_MID   = "#2E9E8A"   # medium teal — third series (volatility chart)
TEAL_LIGHT = "#93D4CB"   # light mint  — secondary series


# ── Data Functions ─────────────────────────────────────────────────────────────

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

    # Build flat dicts so the loop and chart function don't need to know the tuple shape
    all_periods  = {full_label: (clean.index.min(), clean.index.max())}
    period_types = {}  # label -> "stress" | "growth"; full sample has no entry (no shading)
    for label, (ptype, start, end) in PERIODS.items():
        all_periods[label]  = (start, end)
        period_types[label] = ptype

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

    return pd.DataFrame(records), all_periods, period_types


# ── Chart Helpers ──────────────────────────────────────────────────────────────

def _circle_legend(ax, handles, loc="upper left", ncol=None, bbox_to_anchor=None):
    """Render a legend with filled-circle markers (Input 1 style).

    Pass bbox_to_anchor to anchor the legend outside the axes area.
    """
    ncol = ncol or len(handles)
    kwargs = dict(
        handles=handles,
        fontsize=9.5,
        frameon=False,
        ncol=ncol,
        loc=loc,
        handletextpad=0.5,
        columnspacing=1.4,
    )
    if bbox_to_anchor is not None:
        kwargs["bbox_to_anchor"] = bbox_to_anchor
    ax.legend(**kwargs)


# ── Chart Functions ────────────────────────────────────────────────────────────

def plot_correlation_bars(corr_df, all_periods, period_types):
    """
    Vertical clustered bar chart of ARBIX–Equity and ARBIX–Bond correlations.
    Bars are grouped vertically along a horizontal x-axis. Shaded column bands
    mark stress / growth periods. Circle legend floated above the chart area
    (Input 1 style) to avoid overlap with bars.
    """
    chart_labels = list(all_periods.keys())
    n_periods    = len(chart_labels)
    x            = np.arange(n_periods)
    width        = 0.35

    fig, ax = plt.subplots(figsize=(14, 7))

    # Shaded vertical column bands: stress = soft red, growth = soft green
    for i, label in enumerate(chart_labels):
        ptype = period_types.get(label)
        if ptype is None:
            continue  # Full Sample — no shading
        shade_color = _STRESS_COLOR if ptype == "stress" else _GROWTH_COLOR
        ax.axvspan(i - 0.48, i + 0.48, color=shade_color, alpha=0.08, zorder=0)

    # Vertical bars — ARBIX vs Equities (left), ARBIX vs Bonds (right)
    bars_eq = ax.bar(x - width / 2, corr_df["CBA-Equity"], width,
                     color=TEAL_DARK,  zorder=2)
    bars_bd = ax.bar(x + width / 2, corr_df["CBA-Bond"],   width,
                     color=TEAL_LIGHT, zorder=2)

    # Data labels
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
    ax.set_ylabel("Correlation Coefficients", fontsize=11)

    # Circle legend anchored above the axes — clears all bars regardless of height
    legend_elements = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_DARK,
               markersize=11, label="ARBIX – MSCI World Index"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_LIGHT,
               markersize=11, label="ARBIX – iShares Treasury Bond ETF"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=_STRESS_COLOR,
               markersize=10, label="Period of Stress"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=_GROWTH_COLOR,
               markersize=10, label="Period of Growth"),
    ]
    _circle_legend(ax, legend_elements,
                   loc="lower center", bbox_to_anchor=(0.5, 1.02), ncol=4)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)
    plt.tight_layout()
    plt.savefig("Correlation_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def plot_volatility(clean):
    """
    Horizontal bar chart of annualised return standard deviations.
    Bars run left-to-right; ARBIX at the top, Equities in the middle, Bonds at
    the bottom. Circle legend stacked vertically (ncol=1) to the right of the
    chart area (Input 1 style).
    """
    vols     = clean[["ARBIX", "Equities", "Bonds"]].std() * np.sqrt(12)
    vols_pct = vols * 100

    # Display order top-to-bottom: ARBIX, Equities, Bonds.
    # barh plots bottom-to-top, so reverse the lists so ARBIX ends up on top.
    labels = ["Bonds", "Equities", "ARBIX"]
    colors = [TEAL_LIGHT, TEAL_MID, TEAL_DARK]
    values = [vols_pct["Bonds"], vols_pct["Equities"], vols_pct["ARBIX"]]

    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.barh(labels, values, color=colors)

    for bar, v in zip(bars, values):
        ax.text(v + 0.1, bar.get_y() + bar.get_height() / 2,
                f"{v:.2f}%", ha="left", va="center", fontsize=9, fontweight="bold")

    # Extend x-axis to 1.45× the max bar so data labels don't reach the legend area
    ax.set_xlim(0, max(values) * 1.45)
    ax.set_xlabel("Annualised Volatility (%)", fontsize=11)

    # Legend: stacked vertically (ncol=1), top-to-bottom order ARBIX → Equities → Bonds,
    # anchored to the right of the axes so it never overlaps the bars.
    legend_elements = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_DARK,
               markersize=11, label="ARBIX"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_MID,
               markersize=11, label="Equities"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_LIGHT,
               markersize=11, label="Bonds"),
    ]
    _circle_legend(ax, legend_elements,
                   loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=1)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)
    plt.tight_layout()
    plt.savefig("Volatility_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def plot_down_market(clean):
    """
    Vertical bar chart comparing average ARBIX return — all months vs equity-down months.
    Teal colour palette; circle legend at upper right (Input 1 style).
    """
    down   = clean[clean["Equities"] < 0]
    values = [clean["ARBIX"].mean(), down["ARBIX"].mean()]
    labels = ["All Months", f"Equity Down Months\n(n={len(down)})"]
    colors = [TEAL_DARK, TEAL_LIGHT]

    fig, ax = plt.subplots(figsize=(6, 5))

    x     = [0.42, 0.84]
    width = 0.32
    bars  = ax.bar(x, values, color=colors, width=width)

    ymax  = max(values)
    ymin  = min(values)
    yspan = ymax - ymin if ymax != ymin else abs(ymax) or 0.01
    ax.set_ylim(ymin - yspan * 0.4, ymax + yspan * 0.4)
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

    legend_elements = [
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_DARK,
               markersize=11, label="All Months"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=TEAL_LIGHT,
               markersize=11, label="Equity Down Months"),
    ]
    _circle_legend(ax, legend_elements, loc="upper right", ncol=2)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)
    plt.tight_layout()
    plt.savefig("Down_Market_Chart.png", dpi=300, bbox_inches="tight")
    plt.close()


def calculate_downside_stats(clean):
    down = clean[clean["Equities"] < 0]

    arbix_compound  = (1 + down["ARBIX"]).prod() - 1
    equity_compound = (1 + down["Equities"]).prod() - 1
    dcr = round((arbix_compound / equity_compound) * 100, 4) if equity_compound != 0 else np.nan

    down_beta = round(down["ARBIX"].cov(down["Equities"]) / down["Equities"].var(), 4)

    return pd.DataFrame([
        {"Statistic": "Downside Capture Ratio (ARBIX vs Equities)", "Value": dcr,       "N (down months)": len(down)},
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
    corr_df, periods, period_types = calculate_correlations(clean)
    downside_df                    = calculate_downside_stats(clean)

    plot_correlation_bars(corr_df, periods, period_types)
    plot_volatility(clean)
    plot_down_market(clean)

    export_outputs(clean, corr_df, downside_df)
    print("Done.")
