"""Retry the 12 failed tweets from the focused dataset scrape.

Uses playwright with longer delays + session refresh between each tweet.
Merges results back into the main twitter_post_engagement.xlsx.
"""
from __future__ import annotations

import asyncio
import json
import os
import random
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


def parse_cookies(s):
    return [{"name": k.strip(), "value": v.strip(), "domain": ".x.com", "path": "/"}
            for part in s.split(";") if "=" in part for k, v in [part.strip().split("=", 1)]]


def extract_metrics(data):
    def find(obj):
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
                        "error":     None,
                    }
            for v in obj.values():
                r = find(v)
                if r: return r
        elif isinstance(obj, list):
            for item in obj:
                r = find(item)
                if r: return r
        return None
    return find(data)


async def scrape_one(ctx, tweet_id, creator, source_url):
    captured = {}

    async def on_response(resp, _cap=captured):
        if ("TweetDetail" in resp.url or "TweetResultByRestId" in resp.url) and resp.status == 200:
            try:
                _cap["data"] = await resp.json()
            except Exception:
                pass

    page = await ctx.new_page()
    page.on("response", on_response)

    # First hit the home page to reset session state
    try:
        await page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(2000)
    except Exception:
        pass

    # Now navigate to the tweet
    try:
        await page.goto(f"https://x.com/i/web/status/{tweet_id}",
                        wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(5000)  # longer wait for GraphQL to fire
    except Exception as e:
        await page.close()
        return {"text": None, "likes": None, "retweets": None, "replies": None,
                "quotes": None, "bookmarks": None, "views": None,
                "timestamp": None, "error": f"goto: {str(e)[:60]}"}

    await page.close()

    if captured.get("data"):
        m = extract_metrics(captured["data"])
        if m:
            return m

    return {"text": None, "likes": None, "retweets": None, "replies": None,
            "quotes": None, "bookmarks": None, "views": None,
            "timestamp": None, "error": "no_graphql_captured"}


async def main():
    df = pd.read_excel(OUT_PATH)
    failed = df[df["error"].notna()].copy()
    print(f"Retrying {len(failed)} failed tweets with long delays + session refresh...\n")

    cookies = parse_cookies(COOKIES_STR)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox"])
        ctx = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        await ctx.add_cookies(cookies)

        for i, row in enumerate(failed.itertuples()):
            print(f"  [{i+1}/{len(failed)}] {row.tweet_id} ({row.creator})...", end=" ", flush=True)
            result = await scrape_one(ctx, row.tweet_id, row.creator, row.source_url)

            if result["error"] is None:
                print(f"OK — likes={result['likes']} retweets={result['retweets']} views={result['views']}")
                for col in ["text","likes","retweets","replies","quotes","bookmarks","views","timestamp","error"]:
                    df.loc[df["tweet_id"] == row.tweet_id, col] = result[col]
            else:
                print(f"STILL FAILED: {result['error']}")

            # Long randomised delay between each
            delay = random.uniform(8, 12)
            print(f"      waiting {delay:.1f}s...", flush=True)
            await asyncio.sleep(delay)

        await browser.close()

    df.to_excel(OUT_PATH, index=False)
    remaining = df[df["error"].notna()]
    print(f"\nDone. Saved. Still failed: {len(remaining)}")
    if not remaining.empty:
        print(remaining[["tweet_id", "creator", "error"]].to_string())

    good = df[df["error"].isna()]
    print(f"\nFinal coverage: {len(good)}/{len(df)} tweets ({len(good)/len(df)*100:.1f}%)")
    print("\nEngagement summary:")
    print(good.groupby("creator")[["likes","retweets","replies","views"]].agg(["median","max"]).round(0).to_string())


if __name__ == "__main__":
    asyncio.run(main())
