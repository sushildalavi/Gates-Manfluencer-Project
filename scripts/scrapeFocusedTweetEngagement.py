"""Fetch per-tweet engagement metrics for every tweet in the Nigeria Focused dataset.

Uses Playwright to intercept Twitter's authenticated GraphQL API responses.
Requires auth_token + ct0 cookies in TWSCRAPE_COOKIES (.env).

Covers 4 Twitter creators: Deyemi Okanlawon, Shola, Agba John Doe, Wizarab (203 tweets).

Output:
  Research Assets/Engagement Metrics/twitter_post_engagement.xlsx
  columns: tweet_id | creator | source_url | text | likes | retweets | replies |
           quotes | bookmarks | views | timestamp
"""
from __future__ import annotations

import asyncio
import json
import os
import random
import re
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

COOKIES_STR = os.getenv("TWSCRAPE_COOKIES", "")
def pick_existing(*paths: Path) -> Path:
    for p in paths:
        if p.exists():
            return p
    return paths[0]

OUT_PATH = pick_existing(
    ROOT / "Research Assets" / "Engagement Metrics" / "twitter_post_engagement.xlsx",
    ROOT / "Research Assets" / "Engagement Metrics" / "twitter post engagement.xlsx",
)
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

FOCUSED_DIR = ROOT / "Nigeria" / "Content Analysis" / "Content - Final"
CREATORS = ["Deyemi Okanlawon", "Shola", "Agba John Doe", "Wizarab"]


def parse_cookies(cookies_str: str) -> list[dict]:
    cookies = []
    for part in cookies_str.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            cookies.append({"name": k.strip(), "value": v.strip(), "domain": ".x.com", "path": "/"})
    return cookies


def load_tweet_ids() -> pd.DataFrame:
    frames = []
    for creator in CREATORS:
        path = pick_existing(
            FOCUSED_DIR / f"{creator}_Twitter.xlsx",
            FOCUSED_DIR / f"{creator} Twitter.xlsx",
        )
        if not path.exists():
            print(f"  ! missing: {path.name}")
            continue
        df = pd.read_excel(path)
        df["creator"] = creator
        df["tweet_id"] = df["Source URL"].str.extract(r"/status/(\d+)")
        df = df.rename(columns={"Source URL": "source_url"})
        frames.append(df[["creator", "tweet_id", "source_url"]])
    combined = pd.concat(frames, ignore_index=True)
    return combined.dropna(subset=["tweet_id"]).drop_duplicates(subset="tweet_id")


def extract_metrics_from_graphql(data: dict) -> dict | None:
    """Parse TweetDetail or TweetResultByRestId GraphQL response for metrics."""
    try:
        # Navigate the nested structure
        def find_tweet_result(obj):
            if isinstance(obj, dict):
                if obj.get("__typename") in ("Tweet", "TweetWithVisibilityResults"):
                    tweet = obj.get("tweet", obj)
                    legacy = tweet.get("legacy", {})
                    views = tweet.get("views", {})
                    if legacy.get("full_text"):
                        return {
                            "text":      legacy.get("full_text", ""),
                            "likes":     legacy.get("favorite_count", 0),
                            "retweets":  legacy.get("retweet_count", 0),
                            "replies":   legacy.get("reply_count", 0),
                            "quotes":    legacy.get("quote_count", 0),
                            "bookmarks": legacy.get("bookmark_count", 0),
                            "views":     int(views.get("count", 0)) if views.get("count") else None,
                            "timestamp": legacy.get("created_at", ""),
                        }
                for v in obj.values():
                    result = find_tweet_result(v)
                    if result:
                        return result
            elif isinstance(obj, list):
                for item in obj:
                    result = find_tweet_result(item)
                    if result:
                        return result
            return None

        return find_tweet_result(data)
    except Exception:
        return None


async def scrape_tweets(tweet_df: pd.DataFrame) -> list[dict]:
    cookies = parse_cookies(COOKIES_STR)
    rows = []
    failed = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        ctx = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        await ctx.add_cookies(cookies)

        for i, row in enumerate(tweet_df.itertuples()):
            captured = {}

            async def on_response(response, _cap=captured):
                if ("TweetDetail" in response.url or "TweetResultByRestId" in response.url) and response.status == 200:
                    try:
                        body = await response.json()
                        _cap["data"] = body
                    except Exception:
                        pass

            page = await ctx.new_page()
            page.on("response", on_response)

            url = f"https://x.com/i/web/status/{row.tweet_id}"
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=20000)
                await page.wait_for_timeout(3000)
            except Exception as e:
                failed += 1
                rows.append({
                    "tweet_id": row.tweet_id, "creator": row.creator,
                    "source_url": row.source_url,
                    "text": None, "likes": None, "retweets": None,
                    "replies": None, "quotes": None, "bookmarks": None,
                    "views": None, "timestamp": None, "error": str(e)[:80],
                })
                await page.close()
                continue

            await page.close()

            if captured.get("data"):
                metrics = extract_metrics_from_graphql(captured["data"])
                if metrics:
                    rows.append({
                        "tweet_id":   row.tweet_id,
                        "creator":    row.creator,
                        "source_url": row.source_url,
                        "text":       metrics["text"],
                        "likes":      metrics["likes"],
                        "retweets":   metrics["retweets"],
                        "replies":    metrics["replies"],
                        "quotes":     metrics["quotes"],
                        "bookmarks":  metrics["bookmarks"],
                        "views":      metrics["views"],
                        "timestamp":  metrics["timestamp"],
                        "error":      None,
                    })
                else:
                    failed += 1
                    rows.append({
                        "tweet_id": row.tweet_id, "creator": row.creator,
                        "source_url": row.source_url,
                        "text": None, "likes": None, "retweets": None,
                        "replies": None, "quotes": None, "bookmarks": None,
                        "views": None, "timestamp": None, "error": "parse_failed",
                    })
            else:
                failed += 1
                rows.append({
                    "tweet_id": row.tweet_id, "creator": row.creator,
                    "source_url": row.source_url,
                    "text": None, "likes": None, "retweets": None,
                    "replies": None, "quotes": None, "bookmarks": None,
                    "views": None, "timestamp": None, "error": "no_graphql_captured",
                })

            done = i + 1
            good = done - failed
            if done % 10 == 0:
                print(f"  ... {done}/{len(tweet_df)}  good={good}  failed={failed}", flush=True)

            # Randomised delay to avoid rate limiting
            await asyncio.sleep(random.uniform(1.5, 2.5))

        await browser.close()

    return rows


def main():
    print("=== Focused Dataset Tweet Engagement (Playwright / GraphQL intercept) ===", flush=True)
    tweet_df = load_tweet_ids()
    print(f"  {len(tweet_df)} unique tweets across {len(CREATORS)} creators", flush=True)
    print(tweet_df.groupby("creator")["tweet_id"].count().to_string(), flush=True)
    print(flush=True)

    rows = asyncio.run(scrape_tweets(tweet_df))

    df = pd.DataFrame(rows)
    df.to_excel(OUT_PATH, index=False)

    good = df[df["error"].isna()]
    print(f"\n  Saved {len(df)} rows → {OUT_PATH.relative_to(ROOT)}", flush=True)
    print(f"  Success: {len(good)} | Failed: {len(df)-len(good)}", flush=True)

    if not good.empty:
        print("\n  Engagement summary by creator:", flush=True)
        print(good.groupby("creator")[["likes", "retweets", "replies", "views"]].agg(["median", "max"]).round(0).to_string(), flush=True)


if __name__ == "__main__":
    main()
