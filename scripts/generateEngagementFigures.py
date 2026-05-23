"""Generate engagement and content analysis visualizations for README and reports.

Outputs to: Figures/Engagement Figures/
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import warnings
warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "Figures" / "Engagement Figures"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── palette ───────────────────────────────────────────────────────────────────
NIGERIA_BLUE  = "#1a3a5c"
KENYA_GREEN   = "#2e7d4f"
ACCENT_GOLD   = "#c9a84c"
LIGHT_GREY    = "#f4f4f4"
MID_GREY      = "#888888"
RED_WARN      = "#c0392b"
PURPLE        = "#6c4f9e"
CORAL         = "#e07b54"

CREATOR_COLORS = {
    "Shola":            "#1a3a5c",
    "Wizarab":          "#2e7d4f",
    "Agba John Doe":    "#c9a84c",
    "Deyemi Okanlawon": "#6c4f9e",
}

STYLE = dict(facecolor="white", edgecolor="none")

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.color": "#e8e8e8",
    "grid.linewidth": 0.7,
    "axes.labelcolor": "#333333",
    "xtick.color": "#555555",
    "ytick.color": "#555555",
})


def save(fig, name):
    p = OUT_DIR / f"{name}.png"
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  saved: {p.name}")


# ── load data ─────────────────────────────────────────────────────────────────

tw  = pd.read_excel(ROOT / "Research Assets/Engagement Metrics/Nigeria/twitter post engagement.xlsx")
yt  = pd.read_excel(ROOT / "Research Assets/Engagement Metrics/youtube video metrics.xlsx")
eda = pd.read_excel(ROOT / "scripts/eda_output/nigeria content eda.xlsx")

# Only good Twitter rows
tw_good = tw[tw["error"].isna()].copy()
tw_good["likes"]   = pd.to_numeric(tw_good["likes"],   errors="coerce").fillna(0)
tw_good["views"]   = pd.to_numeric(tw_good["views"],   errors="coerce").fillna(0)
tw_good["retweets"]= pd.to_numeric(tw_good["retweets"],errors="coerce").fillna(0)
tw_good["replies"] = pd.to_numeric(tw_good["replies"], errors="coerce").fillna(0)


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 1 — Twitter engagement by creator (median likes + views)
# ═══════════════════════════════════════════════════════════════════════════════

def fig_twitter_engagement_by_creator():
    creators = ["Shola", "Wizarab", "Agba John Doe", "Deyemi Okanlawon"]
    med_likes = [tw_good[tw_good.creator==c]["likes"].median() for c in creators]
    med_views = [tw_good[tw_good.creator==c]["views"].median() / 1000 for c in creators]

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    fig.patch.set_facecolor("white")

    colors = [CREATOR_COLORS[c] for c in creators]

    # Likes
    bars = axes[0].barh(creators, med_likes, color=colors, height=0.55, zorder=3)
    axes[0].set_xlabel("Median Likes per Tweet", fontsize=11)
    axes[0].set_title("Median Likes by Creator", fontsize=13, fontweight="bold", pad=10)
    for bar, val in zip(bars, med_likes):
        axes[0].text(bar.get_width() + 10, bar.get_y() + bar.get_height()/2,
                     f"{int(val):,}", va="center", fontsize=10, color="#333")

    # Views
    bars2 = axes[1].barh(creators, med_views, color=colors, height=0.55, zorder=3)
    axes[1].set_xlabel("Median Views per Tweet (thousands)", fontsize=11)
    axes[1].set_title("Median Views by Creator", fontsize=13, fontweight="bold", pad=10)
    for bar, val in zip(bars2, med_views):
        axes[1].text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2,
                     f"{val:.0f}K", va="center", fontsize=10, color="#333")

    for ax in axes:
        ax.set_facecolor(LIGHT_GREY)
        ax.invert_yaxis()

    fig.suptitle("Nigeria Twitter — Engagement by Creator\n(Focused Dataset, 203 tweets)",
                 fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    save(fig, "01 Twitter Engagement by Creator")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 2 — YouTube views comparison Nigeria vs Kenya
# ═══════════════════════════════════════════════════════════════════════════════

def fig_youtube_views():
    fig, ax = plt.subplots(figsize=(13, 6))
    fig.patch.set_facecolor("white")

    labels = [t[:45] + "…" if len(t) > 45 else t for t in yt["title"]]
    views  = yt["views"] / 1000
    likes  = yt["likes"]
    colors = [NIGERIA_BLUE if c == "Nigeria" else KENYA_GREEN for c in yt["country"]]

    y = np.arange(len(yt))
    bars = ax.barh(y, views, color=colors, height=0.6, zorder=3)

    for bar, v, l in zip(bars, views, likes):
        ax.text(bar.get_width() + 2, bar.get_y() + bar.get_height()/2,
                f"{v:.0f}K views · {l:,} likes", va="center", fontsize=9, color="#333")

    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel("Views (thousands)", fontsize=11)
    ax.set_title("YouTube Video Engagement — Nigeria & Kenya\n(Focused Content Analysis Videos)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_facecolor(LIGHT_GREY)
    ax.invert_yaxis()

    patches = [mpatches.Patch(color=NIGERIA_BLUE, label="Nigeria (MENtality)"),
               mpatches.Patch(color=KENYA_GREEN,  label="Kenya")]
    ax.legend(handles=patches, loc="lower right", fontsize=10)

    plt.tight_layout()
    save(fig, "02 YouTube Views Nigeria Kenya")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 3 — Content framing distribution (Nigeria)
# ═══════════════════════════════════════════════════════════════════════════════

def fig_framing_distribution():
    framing = eda["framing__frame"].value_counts()
    labels  = [f.replace("_", " ").title() for f in framing.index]
    colors  = [NIGERIA_BLUE, KENYA_GREEN, ACCENT_GOLD, PURPLE, CORAL,
               RED_WARN, "#999", "#555"][:len(framing)]

    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor("white")
    bars = ax.barh(labels[::-1], framing.values[::-1], color=colors[::-1], height=0.6, zorder=3)
    for bar, val in zip(bars, framing.values[::-1]):
        ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
                str(val), va="center", fontsize=10, color="#333")
    ax.set_xlabel("Number of Content Segments", fontsize=11)
    ax.set_title("Nigeria — Content Framing Distribution\n(LLM Focused Coding, 310 segments)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_facecolor(LIGHT_GREY)
    plt.tight_layout()
    save(fig, "03 Nigeria Content Framing Distribution")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 4 — Misogyny type distribution
# ═══════════════════════════════════════════════════════════════════════════════

def fig_misogyny_distribution():
    miso  = eda["misogyny__misogyny"].value_counts()
    labels = [m.replace("_", " ").title() for m in miso.index]
    total  = miso.sum()
    pcts   = miso.values / total * 100

    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_facecolor("white")
    colors = [LIGHT_GREY if "none" in l.lower() else RED_WARN
              for l in [m.lower() for m in miso.index]]
    colors = [LIGHT_GREY, "#e07b54", "#c0392b", "#9b59b6", "#e74c3c", "#f39c12"][:len(miso)]

    bars = ax.bar(labels, pcts, color=colors, zorder=3, width=0.6)
    for bar, pct, n in zip(bars, pcts, miso.values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                f"{pct:.0f}%\n(n={n})", ha="center", fontsize=9, color="#333")

    ax.set_ylabel("% of Segments", fontsize=11)
    ax.set_title("Nigeria — Misogyny Type by Content Segment\n(LLM Focused Coding)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_facecolor(LIGHT_GREY)
    plt.xticks(rotation=20, ha="right", fontsize=9)
    plt.tight_layout()
    save(fig, "04 Nigeria Misogyny Type Distribution")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 5 — Orientation breakdown by creator
# ═══════════════════════════════════════════════════════════════════════════════

def fig_orientation_by_creator():
    creators = ["Banky Wellington", "Deyemi Okanlawon", "Ebuka Obi-Uchendu",
                "Shola", "Wizarab", "Agba John Doe"]
    prog, regr = [], []
    for c in creators:
        sub = eda[eda.creator == c]["orientation"].value_counts()
        prog.append(sub.get("progressive", 0))
        regr.append(sub.get("regressive", 0))

    x = np.arange(len(creators))
    w = 0.35

    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor("white")
    ax.bar(x - w/2, prog, w, label="Progressive", color=KENYA_GREEN, zorder=3)
    ax.bar(x + w/2, regr, w, label="Regressive",  color=RED_WARN, zorder=3)

    ax.set_xticks(x)
    ax.set_xticklabels(creators, rotation=18, ha="right", fontsize=10)
    ax.set_ylabel("Number of Segments", fontsize=11)
    ax.set_title("Nigeria — Progressive vs Regressive Content by Creator\n(LLM Focused Coding)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.legend(fontsize=11)
    ax.set_facecolor(LIGHT_GREY)
    plt.tight_layout()
    save(fig, "05 Nigeria Orientation by Creator")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 6 — Engagement vs framing (Twitter, Nigeria)
# ═══════════════════════════════════════════════════════════════════════════════

def fig_engagement_vs_framing():
    merged = tw_good.merge(
        eda[["content_id", "framing__frame"]].rename(columns={"framing__frame": "framing"}),
        left_on="source_url",
        right_on="content_id",
        how="inner"
    )
    # Better join via EDA content_id → source_url mapping
    # Use the merged workbook instead
    try:
        m = pd.read_excel(ROOT / "Research Assets/Engagement Metrics/Nigeria/Engagement ContentAnalysis Merged.xlsx",
                          sheet_name="Nigeria_Twitter")
        m["likes"] = pd.to_numeric(m["likes"], errors="coerce").fillna(0)
        m["views"] = pd.to_numeric(m["views"], errors="coerce").fillna(0)
        m = m[m["framing_frame"].notna() & (m["likes"] > 0)]

        framing_eng = m.groupby("framing_frame")[["likes", "views"]].median().sort_values("likes", ascending=False)
        framing_eng.index = [f.replace("_", " ").title() for f in framing_eng.index]

        fig, axes = plt.subplots(1, 2, figsize=(14, 5))
        fig.patch.set_facecolor("white")

        colors = [NIGERIA_BLUE, KENYA_GREEN, ACCENT_GOLD, PURPLE, CORAL, RED_WARN, "#888", "#555"]

        # Likes by framing
        axes[0].barh(framing_eng.index[::-1], framing_eng["likes"][::-1],
                     color=colors[:len(framing_eng)], height=0.55, zorder=3)
        axes[0].set_xlabel("Median Likes", fontsize=11)
        axes[0].set_title("Median Likes by Framing Type", fontsize=12, fontweight="bold")

        # Views by framing
        axes[1].barh(framing_eng.index[::-1], framing_eng["views"][::-1] / 1000,
                     color=colors[:len(framing_eng)], height=0.55, zorder=3)
        axes[1].set_xlabel("Median Views (thousands)", fontsize=11)
        axes[1].set_title("Median Views by Framing Type", fontsize=12, fontweight="bold")

        for ax in axes:
            ax.set_facecolor(LIGHT_GREY)

        fig.suptitle("Nigeria Twitter — Engagement by Content Framing\n(tweets with LLM codes and >0 likes)",
                     fontsize=13, fontweight="bold", y=1.02)
        plt.tight_layout()
        save(fig, "06 Nigeria Engagement by Framing")
    except Exception as e:
        print(f"  skipping fig 6: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 7 — Sentiment distribution across creators
# ═══════════════════════════════════════════════════════════════════════════════

def fig_sentiment_distribution():
    sentiment = eda["sentiment__sentiment"].value_counts()
    labels = [s.title() for s in sentiment.index]
    colors = {"Negative": RED_WARN, "Positive": KENYA_GREEN,
              "Neutral": MID_GREY, "Mixed": ACCENT_GOLD}
    bar_colors = [colors.get(l, NIGERIA_BLUE) for l in labels]

    fig, ax = plt.subplots(figsize=(8, 5))
    fig.patch.set_facecolor("white")
    bars = ax.bar(labels, sentiment.values, color=bar_colors, zorder=3, width=0.5)
    for bar, val in zip(bars, sentiment.values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                str(val), ha="center", fontsize=11, fontweight="bold", color="#333")
    ax.set_ylabel("Number of Segments", fontsize=11)
    ax.set_title("Nigeria — Sentiment Distribution Across All Segments\n(LLM Focused Coding)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_facecolor(LIGHT_GREY)
    plt.tight_layout()
    save(fig, "07 Nigeria Sentiment Distribution")


# ═══════════════════════════════════════════════════════════════════════════════
# FIG 8 — Top tweets by likes (Nigeria Twitter)
# ═══════════════════════════════════════════════════════════════════════════════

def fig_top_tweets():
    top = tw_good[tw_good["likes"] > 0].nlargest(10, "likes")[["creator", "likes", "views", "text"]]
    top["short_text"] = top["text"].str[:55] + "…"
    top = top.reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(13, 6))
    fig.patch.set_facecolor("white")
    colors = [CREATOR_COLORS.get(c, NIGERIA_BLUE) for c in top["creator"]]
    bars = ax.barh(top["short_text"][::-1], top["likes"][::-1], color=colors[::-1], height=0.6, zorder=3)
    for bar, val, views in zip(bars, top["likes"][::-1], top["views"][::-1]):
        ax.text(bar.get_width() + 50, bar.get_y() + bar.get_height()/2,
                f"{int(val):,} likes · {int(views/1000):.0f}K views",
                va="center", fontsize=8.5, color="#333")

    ax.set_xlabel("Likes", fontsize=11)
    ax.set_title("Nigeria — Top 10 Tweets by Likes\n(Focused Dataset)",
                 fontsize=13, fontweight="bold", pad=12)
    ax.set_facecolor(LIGHT_GREY)

    patches = [mpatches.Patch(color=v, label=k) for k, v in CREATOR_COLORS.items()]
    ax.legend(handles=patches, loc="lower right", fontsize=9)
    plt.tight_layout()
    save(fig, "08 Nigeria Top Tweets by Likes")


# ═══════════════════════════════════════════════════════════════════════════════
# Run all
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"Generating figures → {OUT_DIR.relative_to(ROOT)}\n")
    fig_twitter_engagement_by_creator()
    fig_youtube_views()
    fig_framing_distribution()
    fig_misogyny_distribution()
    fig_orientation_by_creator()
    fig_engagement_vs_framing()
    fig_sentiment_distribution()
    fig_top_tweets()
    print(f"\nDone. {len(list(OUT_DIR.glob('*.png')))} figures saved.")
