import pandas as pd, matplotlib, os, tempfile, io, warnings
warnings.filterwarnings("ignore")
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import inspect as _insp

PALETTE  = ["#7c3aed","#06b6d4","#10b981","#f59e0b","#ef4444","#ec4899","#8b5cf6","#f97316"]
STYLE = {
    "figure.facecolor":"#0f0f1e","axes.facecolor":"#0f0f1e",
    "axes.edgecolor":"#2d2d55","axes.labelcolor":"#c4c4e0",
    "xtick.color":"#9898c0","ytick.color":"#9898c0",
    "text.color":"#e0e0f8","grid.color":"#1e1e3a","grid.alpha":0.6,
    "font.family":"sans-serif",
}

def _s():
    for k,v in STYLE.items():
        plt.rcParams[k] = v

def _save(fig, name, tmpdir):
    p = os.path.join(tmpdir, name)
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return p

def generate_all_charts(df: pd.DataFrame, tmpdir: str = None) -> dict:
    """
    Returns a dict mapping chart_name -> file_path.
    Generates: bar, line, pie, scatter, heatmap, histogram, box, area
    """
    if tmpdir is None:
        tmpdir = tempfile.mkdtemp()

    num_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(include=["object","category"]).columns.tolist()
    charts = {}

    if not num_cols:
        return charts

    # ── 1. Bar Chart ──────────────────────────────────────────────────────────
    try:
        _s(); fig, ax = plt.subplots(figsize=(10, 5.2))
        col = num_cols[0]
        if cat_cols and df[cat_cols[0]].nunique() <= 20:
            g = df.groupby(cat_cols[0])[col].mean().nlargest(10)
            labels = [str(x)[:14] for x in g.index]
            bars = ax.bar(labels, g.values, color=PALETTE[:len(g)], edgecolor="none", width=0.62)
            for bar, val in zip(bars, g.values):
                ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+max(g.values)*0.01,
                        f"{val:,.1f}", ha="center", va="bottom", fontsize=8, color="#e0e0f8")
            ax.set_xlabel(cat_cols[0], fontsize=10, labelpad=8)
        else:
            v = df[col].dropna().head(12)
            bars = ax.bar(range(len(v)), v.values, color=PALETTE[0], edgecolor="none", width=0.7)
        ax.set_title(f"Bar Chart — {col}", fontsize=13, fontweight="bold", pad=14)
        ax.set_ylabel(col, fontsize=10)
        ax.grid(axis="y", alpha=0.4); ax.spines[["top","right"]].set_visible(False)
        plt.xticks(rotation=30, ha="right", fontsize=8); plt.tight_layout()
        charts["bar"] = _save(fig, "chart_bar.png", tmpdir)
    except Exception as e:
        print(f"bar error: {e}")

    # ── 2. Line Chart ─────────────────────────────────────────────────────────
    if len(num_cols) >= 1:
        try:
            _s(); fig, ax = plt.subplots(figsize=(10, 5.2))
            for i, col in enumerate(num_cols[:4]):
                d = df[col].dropna().reset_index(drop=True)
                ax.plot(d.index, d.values, color=PALETTE[i], linewidth=2.5, label=col, alpha=0.95)
                ax.fill_between(d.index, d.values, alpha=0.07, color=PALETTE[i])
            ax.set_title("Trend Analysis", fontsize=13, fontweight="bold", pad=14)
            ax.legend(framealpha=0.15, edgecolor="#444466", fontsize=9)
            ax.grid(alpha=0.35); ax.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            charts["line"] = _save(fig, "chart_line.png", tmpdir)
        except Exception as e:
            print(f"line error: {e}")

    # ── 3. Pie Chart ──────────────────────────────────────────────────────────
    if cat_cols:
        try:
            _s(); fig, ax = plt.subplots(figsize=(8, 5.5))
            vc = df[cat_cols[0]].value_counts().head(7)
            wedges, texts, auto = ax.pie(
                vc.values, labels=[str(x)[:13] for x in vc.index],
                colors=PALETTE[:len(vc)], autopct="%1.1f%%",
                startangle=130, pctdistance=0.8,
                wedgeprops={"edgecolor":"#0f0f1e","linewidth":2.5})
            for at in auto:
                at.set_fontsize(9); at.set_color("white"); at.set_fontweight("bold")
            ax.set_title(f"Distribution — {cat_cols[0]}", fontsize=13, fontweight="bold", pad=14)
            plt.tight_layout()
            charts["pie"] = _save(fig, "chart_pie.png", tmpdir)
        except Exception as e:
            print(f"pie error: {e}")

    # ── 4. Scatter Plot ───────────────────────────────────────────────────────
    if len(num_cols) >= 2:
        try:
            _s(); fig, ax = plt.subplots(figsize=(9, 5.2))
            x, y = df[num_cols[0]].dropna(), df[num_cols[1]].dropna()
            mn = min(len(x), len(y))
            ax.scatter(x[:mn], y[:mn], c=PALETTE[0], alpha=0.65, s=45, edgecolors="none")
            # trend line
            try:
                z = np.polyfit(x[:mn], y[:mn], 1)
                p = np.poly1d(z)
                xs = np.linspace(x[:mn].min(), x[:mn].max(), 100)
                ax.plot(xs, p(xs), color=PALETTE[1], linewidth=1.8, linestyle="--", alpha=0.8, label="Trend")
                ax.legend(fontsize=9, framealpha=0.15)
            except:
                pass
            ax.set_xlabel(num_cols[0], fontsize=10); ax.set_ylabel(num_cols[1], fontsize=10)
            ax.set_title(f"Scatter — {num_cols[0]} vs {num_cols[1]}", fontsize=13, fontweight="bold", pad=14)
            ax.grid(alpha=0.3); ax.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            charts["scatter"] = _save(fig, "chart_scatter.png", tmpdir)
        except Exception as e:
            print(f"scatter error: {e}")

    # ── 5. Histogram ──────────────────────────────────────────────────────────
    if num_cols:
        try:
            _s(); fig, ax = plt.subplots(figsize=(9, 5.2))
            col = num_cols[0]
            ax.hist(df[col].dropna(), bins=20, color=PALETTE[0], edgecolor="#0f0f1e", alpha=0.9)
            ax.axvline(df[col].mean(), color=PALETTE[2], linewidth=2, linestyle="--", label=f"Mean: {df[col].mean():.2f}")
            ax.axvline(df[col].median(), color=PALETTE[3], linewidth=2, linestyle=":", label=f"Median: {df[col].median():.2f}")
            ax.set_xlabel(col, fontsize=10); ax.set_ylabel("Frequency", fontsize=10)
            ax.set_title(f"Distribution — {col}", fontsize=13, fontweight="bold", pad=14)
            ax.legend(framealpha=0.15, fontsize=9); ax.grid(axis="y", alpha=0.35)
            ax.spines[["top","right"]].set_visible(False); plt.tight_layout()
            charts["histogram"] = _save(fig, "chart_histogram.png", tmpdir)
        except Exception as e:
            print(f"histogram error: {e}")

    # ── 6. Box Plot ───────────────────────────────────────────────────────────
    if len(num_cols) >= 2:
        try:
            _s(); fig, ax = plt.subplots(figsize=(10, 5.2))
            cols_to_plot = num_cols[:6]
            data = [df[c].dropna().values for c in cols_to_plot]
            # matplotlib 3.9 renamed boxplot(labels=) to tick_labels= and 3.11
            # removed the old name, so this chart was failing silently and every
            # generated deck quietly lost one chart type. Resolve the keyword at
            # call time rather than pinning matplotlib. Fixed 2026-08-18.
            _lbl = ('tick_labels'
                    if 'tick_labels' in _insp.signature(ax.boxplot).parameters
                    else 'labels')
            bp = ax.boxplot(data, patch_artist=True, **{_lbl: cols_to_plot},
                            medianprops=dict(color="white", linewidth=2),
                            whiskerprops=dict(color="#666699"),
                            capprops=dict(color="#666699"),
                            flierprops=dict(marker="o", color=PALETTE[4], markersize=4, alpha=0.5))
            for patch, color in zip(bp['boxes'], PALETTE):
                patch.set_facecolor(color); patch.set_alpha(0.75)
            ax.set_title("Box Plot — Spread & Outliers", fontsize=13, fontweight="bold", pad=14)
            ax.set_ylabel("Value", fontsize=10); ax.grid(axis="y", alpha=0.35)
            ax.spines[["top","right"]].set_visible(False)
            plt.xticks(rotation=20, ha="right", fontsize=9); plt.tight_layout()
            charts["box"] = _save(fig, "chart_box.png", tmpdir)
        except Exception as e:
            print(f"box error: {e}")

    # ── 7. Area / Stacked ─────────────────────────────────────────────────────
    if len(num_cols) >= 2:
        try:
            _s(); fig, ax = plt.subplots(figsize=(10, 5.2))
            sample = df[num_cols[:3]].dropna().head(50).reset_index(drop=True)
            for i, col in enumerate(sample.columns):
                ax.fill_between(sample.index, sample[col], alpha=0.35, color=PALETTE[i], label=col)
                ax.plot(sample.index, sample[col], color=PALETTE[i], linewidth=1.5, alpha=0.9)
            ax.set_title("Area Chart — Volume Over Time", fontsize=13, fontweight="bold", pad=14)
            ax.legend(framealpha=0.15, fontsize=9); ax.grid(alpha=0.3)
            ax.spines[["top","right"]].set_visible(False); plt.tight_layout()
            charts["area"] = _save(fig, "chart_area.png", tmpdir)
        except Exception as e:
            print(f"area error: {e}")

    # ── 8. Correlation Heatmap ────────────────────────────────────────────────
    if len(num_cols) >= 3:
        try:
            _s(); fig, ax = plt.subplots(figsize=(8, 6))
            corr = df[num_cols[:8]].corr()
            cmap = matplotlib.colors.LinearSegmentedColormap.from_list(
                "neu", ["#06b6d4","#0f0f1e","#7c3aed"])
            im = ax.imshow(corr.values, cmap=cmap, vmin=-1, vmax=1)
            ax.set_xticks(range(len(corr.columns))); ax.set_yticks(range(len(corr.columns)))
            ax.set_xticklabels(corr.columns, rotation=35, ha="right", fontsize=8)
            ax.set_yticklabels(corr.columns, fontsize=8)
            for i in range(len(corr)):
                for j in range(len(corr.columns)):
                    ax.text(j, i, f"{corr.iloc[i,j]:.2f}", ha="center", va="center",
                            fontsize=8, color="white" if abs(corr.iloc[i,j]) > 0.4 else "#aaaacc")
            plt.colorbar(im, ax=ax, shrink=0.8)
            ax.set_title("Correlation Matrix", fontsize=13, fontweight="bold", pad=14)
            plt.tight_layout()
            charts["heatmap"] = _save(fig, "chart_heatmap.png", tmpdir)
        except Exception as e:
            print(f"heatmap error: {e}")

    # ── 9. Summary Stats Bar ──────────────────────────────────────────────────
    if len(num_cols) >= 2:
        try:
            _s(); fig, ax = plt.subplots(figsize=(10, 5))
            stats_df = df[num_cols[:6]].describe().loc[["mean","std","min","max"]]
            x = np.arange(len(num_cols[:6])); width = 0.2
            for i, (stat, color) in enumerate(zip(["mean","std","min","max"], PALETTE)):
                ax.bar(x + i*width, stats_df.loc[stat], width, label=stat.capitalize(),
                       color=color, edgecolor="none", alpha=0.85)
            ax.set_xticks(x + width*1.5); ax.set_xticklabels(num_cols[:6], rotation=20, ha="right", fontsize=9)
            ax.set_title("Summary Statistics Comparison", fontsize=13, fontweight="bold", pad=14)
            ax.legend(framealpha=0.15, fontsize=9); ax.grid(axis="y", alpha=0.3)
            ax.spines[["top","right"]].set_visible(False); plt.tight_layout()
            charts["stats"] = _save(fig, "chart_stats.png", tmpdir)
        except Exception as e:
            print(f"stats error: {e}")

    return charts


def pick_chart_for_slide(charts: dict, hint: str) -> str | None:
    """Pick the best chart path for a slide based on AI hint."""
    hint = (hint or "").lower()
    order = []
    if "bar" in hint:     order = ["bar","stats","area","line","pie","scatter","histogram","box","heatmap"]
    elif "line" in hint:  order = ["line","area","bar","scatter","stats","pie","histogram","box","heatmap"]
    elif "pie" in hint:   order = ["pie","bar","stats","line","area","scatter","histogram","box","heatmap"]
    elif "scatter" in hint: order = ["scatter","line","area","bar","stats","pie","histogram","box","heatmap"]
    elif "heat" in hint:  order = ["heatmap","scatter","line","bar","stats","pie","histogram","box","area"]
    elif "box" in hint:   order = ["box","histogram","scatter","line","bar","stats","pie","area","heatmap"]
    else:                 order = ["bar","line","pie","scatter","histogram","area","box","stats","heatmap"]

    for name in order:
        if name in charts:
            return charts[name]
    # fallback: any available
    return next(iter(charts.values()), None)
