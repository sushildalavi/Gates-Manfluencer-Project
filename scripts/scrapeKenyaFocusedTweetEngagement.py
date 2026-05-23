from __future__ import annotations

import asyncio
import os
import random
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

COOKIES_STR = os.getenv("TWSCRAPE_COOKIES", "")
OUT_PATH = ROOT / "Research Assets" / "Engagement Metrics" / "Kenya" / "kenya_x_post_engagement.xlsx"
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

SRC_PATH = ROOT / "Kenya" / "Content Analysis" / "Content - Final" / "Kenya Content Analysis Snippets - With IDs.xlsx"
SRC_SHEET = "Eric (Amerix)"


def parse_cookies(cookies_str: str) -> list[dict]:
    cookies = []
    for part in cookies_str.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            cookies.append({"name": k.strip(), "value": v.strip(), "domain": ".x.com", "path": "/"})
    return cookies


def load_tweet_ids() -> pd.DataFrame:
    df = pd.read_excel(SRC_PATH, sheet_name=SRC_SHEET)
    col = "tweet_url" if "tweet_url" in df.columns else "Source URL"
    out = pd.DataFrame({
        "creator": "Amerix",
        "source_url": df[col].astype(str).str.strip(),
    })
    out = out[out["source_url"].str.contains(r"x\.com/.+/status/\d+", regex=True, na=False)].copy()
    out["tweet_id"] = out["source_url"].str.extract(r"/status/(\d+)")
    out = out.dropna(subset=["tweet_id"]).drop_duplicates(subset="tweet_id")
    return out[["creator", "tweet_id", "source_url"]]


def extract_metrics_from_graphql(data: dict) -> dict | None:
    def find_tweet_result(obj):
        if isinstance(obj, dict):
            if obj.get("__typename") in ("Tweet", "TweetWithVisibilityResults"):
                tweet = obj.get("tweet", obj)
                legacy = tweet.get("legacy", {})
                views = tweet.get("views", {})
                if legacy.get("full_text") is not None:
                    return {
                        "text": legacy.get("full_text", ""),
                        "likes": legacy.get("favorite_count", 0),
                        "retweets": legacy.get("retweet_count", 0),
                        "replies": legacy.get("reply_count", 0),
                        "quotes": legacy.get("quote_count", 0),
                        "bookmarks": legacy.get("bookmark_count", 0),
                        "views": int(views.get("count", 0)) if views.get("count") else None,
                        "timestamp": legacy.get("created_at", ""),
                    }
            for v in obj.values():
                r = find_tweet_result(v)
                if r:
                    return r
        elif isinstance(obj, list):
            for item in obj:
                r = find_tweet_result(item)
                if r:
                    return r
        return None

    try:
        return find_tweet_result(data)
    except Exception:
        return None


async def scrape_tweets(tweet_df: pd.DataFrame) -> list[dict]:
    cookies = parse_cookies(COOKIES_STR)
    rows = []
    failed = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        )
        if cookies:
            await ctx.add_cookies(cookies)

        for i, row in enumerate(tweet_df.itertuples()):
            captured = {}

            async def on_response(response, _cap=captured):
                if ("TweetDetail" in response.url or "TweetResultByRestId" in response.url) and response.status == 200:
                    try:
                        _cap["data"] = await response.json()
                    except Exception:
                        pass

            page = await ctx.new_page()
            page.on("response", on_response)

            try:
                await page.goto(f"https://x.com/i/web/status/{row.tweet_id}", wait_until="domcontentloaded", timeout=25000)
                await page.wait_for_timeout(3500)
            except Exception as e:
                failed += 1
                rows.append({
                    "tweet_id": row.tweet_id,
                    "creator": row.creator,
                    "source_url": row.source_url,
                    "text": None,
                    "likes": None,
                    "retweets": None,
                    "replies": None,
                    "quotes": None,
                    "bookmarks": None,
                    "views": None,
                    "timestamp": None,
                    "error": str(e)[:80],
                })
                await page.close()
                continue

            await page.close()

            if captured.get("data"):
                metrics = extract_metrics_from_graphql(captured["data"])
                if metrics:
                    rows.append({
                        "tweet_id": row.tweet_id,
                        "creator": row.creator,
                        "source_url": row.source_url,
                        "text": metrics["text"],
                        "likes": metrics["likes"],
                        "retweets": metrics["retweets"],
                        "replies": metrics["replies"],
                        "quotes": metrics["quotes"],
                        "bookmarks": metrics["bookmarks"],
                        "views": metrics["views"],
                        "timestamp": metrics["timestamp"],
                        "error": None,
                    })
                else:
                    failed += 1
                    rows.append({
                        "tweet_id": row.tweet_id,
                        "creator": row.creator,
                        "source_url": row.source_url,
                        "text": None,
                        "likes": None,
                        "retweets": None,
                        "replies": None,
                        "quotes": None,
                        "bookmarks": None,
                        "views": None,
                        "timestamp": None,
                        "error": "parse_failed",
                    })
            else:
                failed += 1
                rows.append({
                    "tweet_id": row.tweet_id,
                    "creator": row.creator,
                    "source_url": row.source_url,
                    "text": None,
                    "likes": None,
                    "retweets": None,
                    "replies": None,
                    "quotes": None,
                    "bookmarks": None,
                    "views": None,
                    "timestamp": None,
                    "error": "no_graphql_captured",
                })

            done = i + 1
            if done % 10 == 0:
                print(f"  ... {done}/{len(tweet_df)} good={done-failed} failed={failed}", flush=True)
            await asyncio.sleep(random.uniform(1.2, 2.2))

        await browser.close()

    return rows


def main() -> None:
    print("=== Kenya Focused X/Twitter Engagement Scrape ===", flush=True)
    if not COOKIES_STR:
        print("WARNING: TWSCRAPE_COOKIES is empty; requests will likely fail.", flush=True)

    tweet_df = load_tweet_ids()
    print(f"tweets to scrape: {len(tweet_df)}", flush=True)

    rows = asyncio.run(scrape_tweets(tweet_df))
    out = pd.DataFrame(rows)
    out.to_excel(OUT_PATH, index=False)

    good = out[out["error"].isna()]
    print(f"saved: {OUT_PATH.relative_to(ROOT)}")
    print(f"success={len(good)} failed={len(out)-len(good)}")


if __name__ == "__main__":
    main()
