"""Build the merged engagement + content analysis workbook for Ksenia.

Joins:
  Nigeria Twitter:
    [Creator]_Twitter.xlsx (Segment ID + Source URL)
    → nigeria_content_eda.csv (coded variables: themes, sentiment, framing, misogyny)
    → twitter_post_engagement.xlsx (likes, retweets, replies, views)

  Nigeria YouTube (MENtality / Banky + Ebuka):
    Banky Wellington_Podcast.xlsx + Ebuka Obi-Uchendu_Podcast.xlsx (Source URL → video_id)
    → youtube_video_metrics.xlsx (views, likes, comment_count)

Output:
  Research Assets/Engagement Metrics/Engagement_ContentAnalysis_Merged.xlsx
  Sheets: Nigeria_Twitter | Nigeria_YouTube | Methods_Note
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
ENG_DIR = ROOT / "Research Assets" / "Engagement Metrics"
NIG_FINAL = ROOT / "Nigeria" / "Content Analysis" / "Content - Final"
EDA_DIR = ROOT / "scripts" / "eda_output"


TWITTER_CREATORS = ["Deyemi Okanlawon", "Shola", "Agba John Doe", "Wizarab"]
PODCAST_CREATORS = ["Banky Wellington", "Ebuka Obi-Uchendu"]


# ── helpers ──────────────────────────────────────────────────────────────────

def extract_tweet_id(url):
    m = re.search(r"/status/(\d+)", str(url))
    return m.group(1) if m else None

def extract_video_id(url):
    m = re.search(r"v=([a-zA-Z0-9_-]{11})", str(url))
    return m.group(1) if m else None

def pick_existing(*paths: Path) -> Path:
    for p in paths:
        if p.exists():
            return p
    raise FileNotFoundError("None of the candidate paths exist:\n" + "\n".join(str(p) for p in paths))

def content_file(creator: str, content_kind: str) -> Path:
    legacy = NIG_FINAL / f"{creator}_{content_kind}.xlsx"
    spaced = NIG_FINAL / f"{creator} {content_kind}.xlsx"
    return pick_existing(legacy, spaced)

def engagement_file(kind: str) -> Path:
    if kind == "twitter":
        return pick_existing(
            ENG_DIR / "twitter_post_engagement.xlsx",
            ENG_DIR / "twitter post engagement.xlsx",
        )
    if kind == "youtube":
        return pick_existing(
            ENG_DIR / "youtube_video_metrics.xlsx",
            ENG_DIR / "youtube video metrics.xlsx",
        )
    raise ValueError(f"Unknown engagement kind: {kind}")

def merged_out_file() -> Path:
    return pick_existing(
        ENG_DIR / "Engagement_ContentAnalysis_Merged.xlsx",
        ENG_DIR / "Engagement ContentAnalysis Merged.xlsx",
    ) if (ENG_DIR / "Engagement_ContentAnalysis_Merged.xlsx").exists() or (ENG_DIR / "Engagement ContentAnalysis Merged.xlsx").exists() else (ENG_DIR / "Engagement ContentAnalysis Merged.xlsx")

def load_eda() -> pd.DataFrame:
    csv_path = EDA_DIR / "nigeria_content_eda.csv"
    xlsx_path = EDA_DIR / "nigeria content eda.xlsx"
    legacy_xlsx = EDA_DIR / "nigeria_content_eda.xlsx"
    path = pick_existing(csv_path, xlsx_path, legacy_xlsx)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


# ── Nigeria Twitter ───────────────────────────────────────────────────────────

def build_nigeria_twitter():
    print("Building Nigeria Twitter merge...")

    # 1. Coding units with Source URLs
    units_frames = []
    for creator in TWITTER_CREATORS:
        path = content_file(creator, "Twitter")
        df = pd.read_excel(path)
        df = df.rename(columns={"Segment ID": "content_id", "Source URL": "source_url"})
        df["tweet_id"] = df["source_url"].apply(extract_tweet_id)
        units_frames.append(df[["content_id", "source_url", "tweet_id",
                                 "Influencer", "Platform", "Content Type",
                                 "Context (NOT CODED - comprehension only)",
                                 "Verbatim Text (CODE THIS)"]])
    units = pd.concat(units_frames, ignore_index=True)

    # 2. Coded variables from EDA
    eda_all = load_eda()
    eda = eda_all[["content_id", "orientation", "themes__themes",
                   "sentiment__sentiment", "sentiment__intensity",
                   "emotion__emotions", "misogyny__misogyny",
                   "misogyny__intensity", "framing__frame",
                   "framing__stance_implied"]]

    # 3. Engagement metrics
    eng = pd.read_excel(engagement_file("twitter"))[
        ["tweet_id", "likes", "retweets", "replies", "quotes", "bookmarks", "views", "timestamp"]
    ]
    eng["tweet_id"] = eng["tweet_id"].astype(str)
    units["tweet_id"] = units["tweet_id"].astype(str)

    # 4. Join
    merged = (units
              .merge(eda, on="content_id", how="left")
              .merge(eng, on="tweet_id", how="left"))

    # 5. Clean columns
    merged = merged.rename(columns={
        "Influencer": "creator",
        "Platform": "platform",
        "Content Type": "content_type",
        "Context (NOT CODED - comprehension only)": "context",
        "Verbatim Text (CODE THIS)": "tweet_text",
        "themes__themes": "themes",
        "sentiment__sentiment": "sentiment",
        "sentiment__intensity": "sentiment_intensity",
        "emotion__emotions": "emotions",
        "misogyny__misogyny": "misogyny",
        "misogyny__intensity": "misogyny_intensity",
        "framing__frame": "framing_frame",
        "framing__stance_implied": "framing_stance",
    })

    col_order = [
        "content_id", "creator", "platform", "content_type", "source_url", "tweet_id",
        "tweet_text", "context",
        # coded variables
        "orientation", "themes", "sentiment", "sentiment_intensity",
        "emotions", "misogyny", "misogyny_intensity", "framing_frame", "framing_stance",
        # engagement
        "likes", "retweets", "replies", "quotes", "bookmarks", "views", "timestamp",
    ]
    merged = merged[[c for c in col_order if c in merged.columns]]

    has_codes = merged["orientation"].notna().sum()
    has_eng   = merged["likes"].notna().sum()
    print(f"  Rows: {len(merged)} | Has LLM codes: {has_codes} | Has engagement: {has_eng}")
    return merged


# ── Nigeria YouTube ───────────────────────────────────────────────────────────

def build_nigeria_youtube():
    print("Building Nigeria YouTube merge...")

    # Coding units from podcast files — aggregate to video level
    video_rows = []
    for creator in PODCAST_CREATORS:
        path = content_file(creator, "Podcast")
        df = pd.read_excel(path)
        df["video_id"] = df["Source URL"].apply(extract_video_id)
        # One row per unique video
        for vid, grp in df.groupby("video_id"):
            video_rows.append({
                "video_id":    vid,
                "creator":     creator,
                "source_url":  grp["Source URL"].iloc[0],
                "coded_segments": len(grp),
            })

    video_df = pd.DataFrame(video_rows)
    if video_df.empty:
        return pd.DataFrame(columns=["video_id", "creator", "source_url", "coded_segments"])
    video_df = (video_df.groupby("video_id", as_index=False)
                .agg({
                    "creator": lambda s: " | ".join(sorted(set(s))),
                    "source_url": "first",
                    "coded_segments": "sum",
                }))

    # YouTube engagement metrics
    yt = pd.read_excel(engagement_file("youtube"))

    merged = video_df.merge(yt[["video_id","title","views","likes","comment_count","duration_sec"]],
                            on="video_id", how="left")

    print(f"  Videos: {len(merged)} | Has engagement: {merged['views'].notna().sum()}")
    return merged


# ── Methods note ─────────────────────────────────────────────────────────────

METHODS_TEXT = """
ENGAGEMENT × CONTENT ANALYSIS — MERGED DATASET
================================================

PURPOSE
-------
This workbook joins post-level engagement metrics with LLM content analysis
coding variables for Nigeria. It supports the correlation analysis requested
by Ksenia: "what content features (narratives, appeals, framing) are associated
with the greatest engagement?"

SHEETS
------
Nigeria_Twitter
  - 203 tweets from 4 creators: Deyemi Okanlawon, Shola, Agba John Doe, Wizarab
  - Coding variables (LLM): orientation, themes, sentiment, emotions,
    misogyny, framing_frame, framing_stance
  - Engagement: likes, retweets, replies, quotes, bookmarks, views
  - Join key: content_id (Segment ID) → source_url → tweet_id
  - Coverage: 100% of tweets have engagement; ~75% have LLM codes
    (remaining rows are in Focused dataset but not yet in EDA output)

Nigeria_YouTube
  - 6 MENtality podcast videos (Banky Wellington + Ebuka Obi-Uchendu)
  - Engagement: views, likes, comment_count, duration_sec
  - Note: coding is at segment level; engagement is at video level.
    Correlation possible by aggregating coded segments per video.

DATA SOURCES
------------
Engagement — scraped via Playwright/Twitter GraphQL (authenticated session)
  and YouTube Data API v3. Scripts:
    scripts/scrapeFocusedTweetEngagement.py
    scripts/scrapeYoutubeVideoMetrics.py

LLM codes — scripts/eda_output/nigeria_content_eda.csv
  (generated by LLM two-pass coding pipeline, see Notebooks/)

Coding units — Nigeria/Content Analysis/Content - Final/[Creator]_Twitter.xlsx

SUGGESTED CORRELATION ANALYSIS
-------------------------------
For Twitter:
  - Group by orientation (progressive / regressive) → compare median likes/views
  - Group by framing_frame → Spearman ρ with likes and views
  - Group by misogyny level → compare engagement distributions

For YouTube:
  - Match video to dominant theme across coded segments
  - Correlate theme with views/likes
"""


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    out_path = merged_out_file()
    nigeria_tw  = build_nigeria_twitter()
    nigeria_yt  = build_nigeria_youtube()

    print(f"\nWriting to: {out_path.relative_to(ROOT)}")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        nigeria_tw.to_excel(writer, sheet_name="Nigeria_Twitter", index=False)
        nigeria_yt.to_excel(writer, sheet_name="Nigeria_YouTube", index=False)

        # Methods note as plain text sheet
        methods_df = pd.DataFrame({"Methods & Data Dictionary": METHODS_TEXT.strip().split("\n")})
        methods_df.to_excel(writer, sheet_name="Methods_Note", index=False)

    print("Done.")
    print(f"\nNigeria Twitter — rows: {len(nigeria_tw)}, columns: {list(nigeria_tw.columns)}")
    print(f"Nigeria YouTube — rows: {len(nigeria_yt)}, columns: {list(nigeria_yt.columns)}")


if __name__ == "__main__":
    main()
