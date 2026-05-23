"""Generate evaluation metric figures for both reports.

Produces:
  Figures/Reliability Figures/
    01 Agreement Summary by Track and Country.png       (grouped bar, both reports)
    02 Content Alpha by Field Nigeria.png               (Content report)
    03 Content Alpha by Field Kenya.png                 (Content report)
    04 Audience Alpha by Field Nigeria.png              (Audience report)
    05 Audience Alpha by Field Kenya.png                (Audience report)
    06 F1 Summary by Track and Country.png              (both reports)
    07 Alpha Distribution Boxplot.png                   (both reports)
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import warnings
warnings.filterwarnings("ignore")

ROOT    = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "Figures" / "Reliability Figures"
OUT_DIR.mkdir(parents=True, exist_ok=True)
NB_OUT  = ROOT / "Notebooks" / "outputs"

# ── data ──────────────────────────────────────────────────────────────────────
summary  = pd.read_excel(NB_OUT / "Research Grade Metrics Summary.xlsx")
alpha_fld= pd.read_excel(NB_OUT / "Research Grade Alpha By Field.xlsx")
f1       = pd.read_excel(NB_OUT / "Precision Recall F1 Summary.xlsx")

# ── palette ───────────────────────────────────────────────────────────────────
BLUE   = "#1a3a5c"
ORANGE = "#c0504d"
GREEN  = "#2e7d4f"
GOLD   = "#c9a84c"
LGREY  = "#f5f5f5"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.color": "#e8e8e8",
    "grid.linewidth": 0.7,
    "axes.labelcolor": "#333",
    "xtick.color": "#555",
    "ytick.color": "#555",
})

def save(fig, name):
    p = OUT_DIR / f"{name}.png"
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  saved: {p.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 1 — Agreement Summary (exact agreement %, kappa, alpha)
# ═══════════════════════════════════════════════════════════════════════════════
def fig_agreement_summary():
    hh   = summary[summary.metric_group == "human_human"].copy()
    hllm = summary[summary.metric_group == "human_llm"].copy()

    labels = ["Content\nNigeria", "Content\nKenya", "Audience\nNigeria", "Audience\nKenya"]
    hh_exact   = hh["exact_agreement_pct"].values
    hllm_exact = hllm["exact_agreement_pct"].values
    hh_alpha   = hh["alpha_mean"].values
    hh_kappa   = hh["kappa_mean"].values

    x = np.arange(len(labels))
    w = 0.25

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor("white")

    # Panel 1: exact agreement
    b1 = axes[0].bar(x - w/2, hh_exact,   w, label="Human-Human", color=BLUE,   zorder=3)
    b2 = axes[0].bar(x + w/2, hllm_exact, w, label="Human-LLM",   color=ORANGE, zorder=3)
    for bar, v in list(zip(b1, hh_exact)) + list(zip(b2, hllm_exact)):
        axes[0].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
                     f"{v:.1f}%", ha="center", va="bottom", fontsize=9, fontweight="bold")
    axes[0].set_xticks(x); axes[0].set_xticklabels(labels, fontsize=10)
    axes[0].set_ylabel("Exact Agreement (%)", fontsize=11)
    axes[0].set_ylim(0, 100)
    axes[0].set_title("Exact Agreement by Track and Country", fontsize=12, fontweight="bold", pad=10)
    axes[0].legend(fontsize=10)
    axes[0].set_facecolor(LGREY)

    # Panel 2: alpha + kappa (human-human only)
    b3 = axes[1].bar(x - w/2, hh_alpha, w, label="Mean Krippendorff Alpha", color=GREEN, zorder=3)
    b4 = axes[1].bar(x + w/2, hh_kappa, w, label="Mean Cohen Kappa",        color=GOLD,  zorder=3)
    for bar, v in list(zip(b3, hh_alpha)) + list(zip(b4, hh_kappa)):
        axes[1].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.005,
                     f"{v:.3f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    axes[1].set_xticks(x); axes[1].set_xticklabels(labels, fontsize=10)
    axes[1].set_ylabel("Coefficient Value", fontsize=11)
    axes[1].set_ylim(-0.05, 0.5)
    axes[1].axhline(0, color="#aaa", linewidth=0.8, linestyle="--")
    axes[1].set_title("Human-Human IRR Coefficients by Track and Country", fontsize=12, fontweight="bold", pad=10)
    axes[1].legend(fontsize=10)
    axes[1].set_facecolor(LGREY)

    fig.suptitle("Table 7 Summary: Inter-Rater Reliability by Track and Country",
                 fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    save(fig, "01 Agreement Summary by Track and Country")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 2/3/4/5 — Alpha by field (one per track-country)
# ═══════════════════════════════════════════════════════════════════════════════
def fig_alpha_by_field(track, country, fig_num, color):
    sub = alpha_fld[(alpha_fld.track == track) & (alpha_fld.country == country)].copy()
    sub = sub.dropna(subset=["kripp_alpha"]).sort_values("kripp_alpha", ascending=True)

    fig, ax = plt.subplots(figsize=(10, max(5, len(sub) * 0.38)))
    fig.patch.set_facecolor("white")

    colors = [ORANGE if v < 0 else color for v in sub["kripp_alpha"]]
    bars = ax.barh(sub["field"], sub["kripp_alpha"], color=colors, height=0.6, zorder=3)

    for bar, val in zip(bars, sub["kripp_alpha"]):
        x_pos = bar.get_width() + 0.005 if val >= 0 else bar.get_width() - 0.005
        ha = "left" if val >= 0 else "right"
        ax.text(x_pos, bar.get_y() + bar.get_height()/2,
                f"{val:.3f}", va="center", ha=ha, fontsize=9)

    ax.axvline(0,    color="#aaa",    linewidth=0.8, linestyle="--")
    ax.axvline(0.2,  color=GREEN,     linewidth=0.8, linestyle=":", alpha=0.7, label="Fair (0.20)")
    ax.axvline(0.4,  color="#2e7d4f", linewidth=0.8, linestyle=":", alpha=0.5, label="Moderate (0.40)")
    ax.set_xlabel("Krippendorff Alpha", fontsize=11)
    ax.set_title(f"Krippendorff Alpha by Question Field\n{track.title()} Track, {country}",
                 fontsize=12, fontweight="bold", pad=12)
    ax.set_facecolor(LGREY)
    ax.legend(fontsize=9, loc="lower right")
    ax.set_xlim(-0.25, 0.85)

    # Colour legend
    neg_patch = mpatches.Patch(color=ORANGE, label="Below chance (alpha < 0)")
    pos_patch = mpatches.Patch(color=color,  label="Above chance (alpha >= 0)")
    ax.legend(handles=[pos_patch, neg_patch], fontsize=9, loc="lower right")

    plt.tight_layout()
    save(fig, f"{fig_num:02d} {track.title()} Alpha by Field {country}")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 6 — F1 Summary
# ═══════════════════════════════════════════════════════════════════════════════
def fig_f1_summary():
    labels = [f"{r.track.title()}\n{r.country}" for _, r in f1.iterrows()]
    micro  = f1["single micro f1"].values * 100
    multi  = f1["multi token f1"].values * 100
    macro  = f1["single macro f1"].values * 100

    x = np.arange(len(labels))
    w = 0.25

    fig, ax = plt.subplots(figsize=(11, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor(LGREY)

    b1 = ax.bar(x - w,   micro, w, label="Single-choice Micro-F1", color=BLUE,   zorder=3)
    b2 = ax.bar(x,       macro, w, label="Single-choice Macro-F1", color=GREEN,  zorder=3)
    b3 = ax.bar(x + w,   multi, w, label="Multi-choice Token-F1",  color=GOLD,   zorder=3)

    for bars in [b1, b2, b3]:
        for bar in bars:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                    f"{bar.get_height():.1f}", ha="center", va="bottom", fontsize=9, fontweight="bold")

    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=11)
    ax.set_ylabel("F1 Score (%)", fontsize=11)
    ax.set_ylim(0, 100)
    ax.set_title("Table 8 Summary: Human-LLM Classification Performance (F1) by Track and Country",
                 fontsize=12, fontweight="bold", pad=12)
    ax.legend(fontsize=10)

    plt.tight_layout()
    save(fig, "06 F1 Summary by Track and Country")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 7 — Alpha distribution boxplot
# ═══════════════════════════════════════════════════════════════════════════════
def fig_alpha_boxplot():
    groups = {
        "Content\nNigeria": alpha_fld[(alpha_fld.track == "content") & (alpha_fld.country == "Nigeria")]["kripp_alpha"].dropna(),
        "Content\nKenya":   alpha_fld[(alpha_fld.track == "content") & (alpha_fld.country == "Kenya")]["kripp_alpha"].dropna(),
        "Audience\nNigeria":alpha_fld[(alpha_fld.track == "audience") & (alpha_fld.country == "Nigeria")]["kripp_alpha"].dropna(),
        "Audience\nKenya":  alpha_fld[(alpha_fld.track == "audience") & (alpha_fld.country == "Kenya")]["kripp_alpha"].dropna(),
    }

    fig, ax = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor(LGREY)

    colors = [BLUE, GREEN, BLUE, GREEN]
    bp = ax.boxplot(
        [g.values for g in groups.values()],
        labels=list(groups.keys()),
        patch_artist=True,
        medianprops=dict(color="white", linewidth=2),
        whiskerprops=dict(linewidth=1.2),
        capprops=dict(linewidth=1.2),
        flierprops=dict(marker="o", markersize=5, alpha=0.5),
        zorder=3,
    )
    for patch, c in zip(bp["boxes"], colors):
        patch.set_facecolor(c)
        patch.set_alpha(0.75)

    ax.axhline(0,   color="#aaa", linewidth=0.8, linestyle="--", label="Chance (0)")
    ax.axhline(0.2, color=GOLD,   linewidth=0.8, linestyle=":",  label="Fair threshold (0.20)")
    ax.axhline(0.4, color=ORANGE, linewidth=0.8, linestyle=":",  label="Moderate threshold (0.40)")
    ax.set_ylabel("Krippendorff Alpha", fontsize=11)
    ax.set_title("Distribution of Field-Level Krippendorff Alpha by Track and Country",
                 fontsize=12, fontweight="bold", pad=12)
    ax.legend(fontsize=9)

    plt.tight_layout()
    save(fig, "07 Alpha Distribution Boxplot")


# ── run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"Output: {OUT_DIR.relative_to(ROOT)}\n")
    fig_agreement_summary()
    fig_alpha_by_field("content",  "Nigeria", 2, BLUE)
    fig_alpha_by_field("content",  "Kenya",   3, GREEN)
    fig_alpha_by_field("audience", "Nigeria", 4, BLUE)
    fig_alpha_by_field("audience", "Kenya",   5, GREEN)
    fig_f1_summary()
    fig_alpha_boxplot()
    print(f"\nDone. {len(list(OUT_DIR.glob('*.png')))} figures saved.")
