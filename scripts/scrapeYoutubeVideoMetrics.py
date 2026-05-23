"""Fetch video-level engagement metrics for all YouTube videos in the Focused datasets.

Pulls views, likes, comment count, duration from YouTube Data API v3.
GOOGLE_API_KEY must be set in .env — free, 10k units/day quota.

Sources:
  Nigeria: Content - Final / Banky Wellington_Podcast.xlsx + Ebuka Obi-Uchendu_Podcast.xlsx
  Kenya:   Audience Comments - Complete / youtube_* CSVs (pageUrl column)

Output:
  Research Assets/Engagement Metrics/youtube_video_metrics.xlsx
  columns: video_id | title | country | creator | views | likes | comment_count | duration_sec | url
"""
from __future__ import annotations

import os
import re
import urllib.request
import urllib.parse
import json
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

API_KEY = os.getenv("GOOGLE_API_KEY")
assert API_KEY, "GOOGLE_API_KEY must be set in .env"

OUT_PATH = ROOT / "Research Assets" / "Engagement Metrics" / "youtube_video_metrics.xlsx"
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

YT_API = "https://www.googleapis.com/youtube/v3/videos"


def extract_video_id(url: str) -> str | None:
    m = re.search(r"v=([a-zA-Z0-9_-]{11})", str(url))
    return m.group(1) if m else None


def collect_video_ids() -> list[dict]:
    records = []
    seen = set()

    def add(vid: str, country: str, creator: str):
        if vid and vid not in seen:
            seen.add(vid)
            records.append({"video_id": vid, "country": country, "creator": creator})

    # Nigeria: Banky Wellington Podcast
    banky = ROOT / "Nigeria" / "Content Analysis" / "Content - Final" / "Banky Wellington_Podcast.xlsx"
    if banky.exists():
        df = pd.read_excel(banky)
        for url in df.get("Source URL", pd.Series(dtype=str)).dropna().unique():
            add(extract_video_id(url), "Nigeria", "Banky Wellington")

    # Nigeria: Ebuka Obi-Uchendu Podcast
    ebuka = ROOT / "Nigeria" / "Content Analysis" / "Content - Final" / "Ebuka Obi-Uchendu_Podcast.xlsx"
    if ebuka.exists():
        df = pd.read_excel(ebuka)
        for url in df.get("Source URL", pd.Series(dtype=str)).dropna().unique():
            add(extract_video_id(url), "Nigeria", "Ebuka Obi-Uchendu")

    # Kenya: Audience Comments Complete youtube_* folders
    kenya_base = ROOT / "Kenya" / "Audience Analysis" / "Audience Comments - Complete"
    if kenya_base.exists():
        for folder in kenya_base.iterdir():
            if folder.is_dir() and "youtube" in folder.name.lower():
                # derive creator name from folder
                creator = folder.name.replace("youtube_", "").replace("_", " ").title()
                for csv_file in folder.glob("*.csv"):
                    try:
                        df = pd.read_csv(csv_file)
                        if "pageUrl" in df.columns:
                            for url in df["pageUrl"].dropna().unique():
                                add(extract_video_id(url), "Kenya", creator)
                    except Exception:
                        pass

    return records


def duration_to_seconds(iso: str) -> int:
    m = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", iso)
    if not m:
        return 0
    h, mn, s = (int(x or 0) for x in m.groups())
    return h * 3600 + mn * 60 + s


def fetch_metrics(video_ids: list[str]) -> dict[str, dict]:
    results = {}
    # API allows up to 50 ids per request
    for i in range(0, len(video_ids), 50):
        batch = video_ids[i:i+50]
        params = urllib.parse.urlencode({
            "part": "snippet,statistics,contentDetails",
            "id": ",".join(batch),
            "key": API_KEY,
        })
        url = f"{YT_API}?{params}"
        with urllib.request.urlopen(url) as resp:
            data = json.loads(resp.read())
        for item in data.get("items", []):
            vid = item["id"]
            stats = item.get("statistics", {})
            snippet = item.get("snippet", {})
            cd = item.get("contentDetails", {})
            results[vid] = {
                "title": snippet.get("title", ""),
                "views": int(stats.get("viewCount", 0)),
                "likes": int(stats.get("likeCount", 0)),
                "comment_count": int(stats.get("commentCount", 0)),
                "duration_sec": duration_to_seconds(cd.get("duration", "PT0S")),
            }
    return results


def main():
    print("=== YouTube Video Metrics ===")
    records = collect_video_ids()
    print(f"  Found {len(records)} unique video IDs")
    for r in records:
        print(f"    {r['video_id']} | {r['country']} | {r['creator']}")

    video_ids = [r["video_id"] for r in records]
    print(f"\n  Fetching from YouTube Data API...")
    metrics = fetch_metrics(video_ids)
    print(f"  Got metrics for {len(metrics)} videos")

    rows = []
    for r in records:
        vid = r["video_id"]
        m = metrics.get(vid, {})
        rows.append({
            "video_id": vid,
            "title": m.get("title", "NOT FOUND"),
            "country": r["country"],
            "creator": r["creator"],
            "views": m.get("views", None),
            "likes": m.get("likes", None),
            "comment_count": m.get("comment_count", None),
            "duration_sec": m.get("duration_sec", None),
            "url": f"https://www.youtube.com/watch?v={vid}",
        })

    df = pd.DataFrame(rows)
    df.to_excel(OUT_PATH, index=False)
    print(f"\n  Saved to: {OUT_PATH.relative_to(ROOT)}")
    print(df[["title", "country", "creator", "views", "likes", "comment_count"]].to_string())


if __name__ == "__main__":
    main()
