from __future__ import annotations

import asyncio
import os
import random
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import async_playwright

ROOT=Path(__file__).resolve().parents[1]
load_dotenv(ROOT/'.env')
COOKIES_STR=os.getenv('TWSCRAPE_COOKIES','')
OUT=ROOT/'Research Assets'/'Engagement Metrics'/'Kenya'/'kenya_x_kibe_source_post_engagement.xlsx'
OUT.parent.mkdir(parents=True,exist_ok=True)

URLS=[
    'https://x.com/kibeandy/status/2027352101924257911',
    'https://x.com/kibeandy/status/2023656512288125173',
]

def parse_cookies(s):
    out=[]
    for p in s.split(';'):
        p=p.strip()
        if '=' in p:
            k,v=p.split('=',1)
            out.append({'name':k.strip(),'value':v.strip(),'domain':'.x.com','path':'/'})
    return out

def extract_metrics(data):
    def walk(o):
        if isinstance(o,dict):
            if o.get('__typename') in ('Tweet','TweetWithVisibilityResults'):
                tw=o.get('tweet',o); lg=tw.get('legacy',{}); vw=tw.get('views',{})
                if 'full_text' in lg:
                    return {
                        'text':lg.get('full_text',''),'likes':lg.get('favorite_count',0),'retweets':lg.get('retweet_count',0),
                        'replies':lg.get('reply_count',0),'quotes':lg.get('quote_count',0),'bookmarks':lg.get('bookmark_count',0),
                        'views':int(vw.get('count',0)) if vw.get('count') else None,'timestamp':lg.get('created_at','')
                    }
            for v in o.values():
                r=walk(v)
                if r:return r
        elif isinstance(o,list):
            for i in o:
                r=walk(i)
                if r:return r
        return None
    try:return walk(data)
    except:return None

async def main():
    rows=[]
    cookies=parse_cookies(COOKIES_STR)
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True,args=['--no-sandbox','--disable-dev-shm-usage'])
        ctx=await browser.new_context(user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        if cookies:
            await ctx.add_cookies(cookies)
        for u in URLS:
            tid=u.split('/status/')[-1]
            captured={}
            async def on_response(response,_cap=captured):
                if ('TweetDetail' in response.url or 'TweetResultByRestId' in response.url) and response.status==200:
                    try:_cap['data']=await response.json()
                    except:pass
            pg=await ctx.new_page(); pg.on('response',on_response)
            err=None
            try:
                await pg.goto(f'https://x.com/i/web/status/{tid}',wait_until='domcontentloaded',timeout=25000)
                await pg.wait_for_timeout(3500)
            except Exception as e:
                err=str(e)[:120]
            await pg.close()
            if err:
                rows.append({'creator':'Andrew Kibe','tweet_id':tid,'source_url':u,'text':None,'likes':None,'retweets':None,'replies':None,'quotes':None,'bookmarks':None,'views':None,'timestamp':None,'error':err})
            else:
                m=extract_metrics(captured.get('data',{})) if captured.get('data') else None
                if m:
                    rows.append({'creator':'Andrew Kibe','tweet_id':tid,'source_url':u,**m,'error':None})
                else:
                    rows.append({'creator':'Andrew Kibe','tweet_id':tid,'source_url':u,'text':None,'likes':None,'retweets':None,'replies':None,'quotes':None,'bookmarks':None,'views':None,'timestamp':None,'error':'parse_failed_or_no_graphql'})
            await asyncio.sleep(random.uniform(1.2,2.2))
        await browser.close()
    pd.DataFrame(rows).to_excel(OUT,index=False)
    ok=sum(1 for r in rows if not r['error'])
    print('saved',OUT)
    print('success',ok,'failed',len(rows)-ok)

if __name__=='__main__':
    asyncio.run(main())
