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
EDA_PATH = ROOT / "scripts" / "eda_output" / "nigeria_content_eda.csv"
OUT_PATH = ENG_DIR / "Engagement_ContentAnalysis_Merged.xlsx"


TWITTER_CREATORS = ["Deyemi Okanlawon", "Shola", "Agba John Doe", "Wizarab"]
PODCAST_CREATORS = ["Banky Wellington", "Ebuka Obi-Uchendu"]


# ── helpers ──────────────────────────────────────────────────────────────────

def extract_tweet_id(url):
    m = re.search(r"/status/(\d+)", str(url))
    return m.group(1) if m else None

def extract_video_id(url):
    m = re.search(r"v=([a-zA-Z0-9_-]{11})", str(url))
    return m.group(1) if m else None


# ── Nigeria Twitter ───────────────────────────────────────────────────────────

def build_nigeria_twitter():
    print("Building Nigeria Twitter merge...")

    # 1. Coding units with Source URLs
    units_frames = []
    for creator in TWITTER_CREATORS:
        path = NIG_FINAL / f"{creator}_Twitter.xlsx"
        df = pd.read_excel(path)
        df = df.rename(columns={"Segment ID": "content_id", "Source URL": "source_url"})
        df["tweet_id"] = df["source_url"].apply(extract_tweet_id)
        units_frames.append(df[["content_id", "source_url", "tweet_id",
                                 "Influencer", "Platform", "Content Type",
                                 "Context (NOT CODED - comprehension only)",
                                 "Verbatim Text (CODE THIS)"]])
    units = pd.concat(units_frames, ignore_index=True)

    # 2. Coded variables from EDA
    eda = pd.read_csv(EDA_PATH)[["content_id", "orientation", "themes__themes",
                                   "sentiment__sentiment", "sentiment__intensity",
                                   "emotion__emotions", "misogyny__misogyny",
                                   "misogyny__intensity", "framing__frame",
                                   "framing__stance_implied"]]

    # 3. Engagement metrics
    eng = pd.read_excel(ENG_DIR / "twitter_post_engagement.xlsx")[
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
        fname = "Banky Wellington_Podcast.xlsx" if creator == "Banky Wellington" else "Ebuka Obi-Uchendu_Podcast.xlsx"
        path = NIG_FINAL / fname
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

    video_df = pd.DataFrame(video_rows).drop_duplicates(subset="video_id")

    # YouTube engagement metrics
    yt = pd.read_excel(ENG_DIR / "youtube_video_metrics.xlsx")

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
    nigeria_tw  = build_nigeria_twitter()
    nigeria_yt  = build_nigeria_youtube()

    print(f"\nWriting to: {OUT_PATH.relative_to(ROOT)}")
    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
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
