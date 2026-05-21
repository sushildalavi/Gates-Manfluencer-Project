"""
Unified LLM Exploratory Data Analysis — Nigeria + Kenya Content
Gates-Manfluencer-Project | Norman Lear Center × Gates Foundation

Reads directly from the Final xlsx datasets (no translated parquets needed).
Runs: themes, sentiment, emotion, misogyny, framing (content only)
Model: gpt-4o-mini via exploratoryAnalysisLib (with caching)

Outputs:
  scripts/eda_output/Nigeria Content EDA.xlsx   (310 rows)
  scripts/eda_output/Kenya Content EDA.xlsx     (394 rows)
  scripts/eda_output/theme-counts-summary.txt   (verified stats for document)
"""

from __future__ import annotations
import sys, os, json
from pathlib import Path
import pandas as pd

# ── project root & paths ─────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Nigeria" / "scripts"))

from dotenv import load_dotenv
load_dotenv(ROOT / ".env")

import exploratoryAnalysisLib as eal

OUT_DIR = ROOT / "scripts" / "eda_output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_NG = ROOT / "Nigeria"
BASE_KE = ROOT / "Kenya"

NG_CONTENT_XLSX = BASE_NG / "Content Analysis" / "Nigeria Content Analysis Final.xlsx"
KE_CONTENT_XLSX = BASE_KE / "Content Analysis" / "Kenya Content Analysis Final.xlsx"

# ── load Nigeria content ─────────────────────────────────────────────────────
print("Loading Nigeria Content...")
ng_map = {
    "Banky Wellington":   ("progressive", "Verbatim Text "),
    "Deyemi Okanlawon":   ("progressive", "Verbatim Text "),
    "Ebuka Obi-Uchendu":  ("progressive", "Verbatim Text "),
    "Shola":              ("regressive",  "Tweet"),
    "Wizarab":            ("regressive",  "Tweet"),
    "Agba John Doe":      ("regressive",  "Tweet"),
}

ng_rows = []
for sheet, (orientation, tcol) in ng_map.items():
    df = pd.read_excel(NG_CONTENT_XLSX, sheet_name=sheet)
    ctx_col = [c for c in df.columns if "context" in c.lower()][0]
    id_col  = [c for c in df.columns if "content id" in c.lower() or "id" in c.lower()][0]
    for _, row in df.iterrows():
        text = str(row[tcol]).strip()
        ctx  = str(row[ctx_col]).strip()
        ng_rows.append({
            "country":     "Nigeria",
            "content_id":  str(row[id_col]),
            "creator":     sheet,
            "orientation": orientation,
            "context":     ctx,
            "text":        text,
        })

ng_df = pd.DataFrame(ng_rows)
print(f"  {len(ng_df)} snippets | prog={( ng_df.orientation=='progressive').sum()} reg={(ng_df.orientation=='regressive').sum()}")

# ── load Kenya content ───────────────────────────────────────────────────────
print("Loading Kenya Content...")
ke_map = {
    "Eddy Kimani":              ("progressive", "Text"),
    "Onyango Otieno (Rixpoet)": ("progressive", "Text"),
    "Philip Karanja":           ("progressive", "Text"),
    "Eric (Amerix)":            ("regressive",  "Text"),
    "Andrew Kibe":              ("regressive",  "Text"),
}

ke_rows = []
for sheet, (orientation, tcol) in ke_map.items():
    df = pd.read_excel(KE_CONTENT_XLSX, sheet_name=sheet)
    ctx_col = [c for c in df.columns if "context" in c.lower()][0]
    id_col  = df.columns[0]
    for _, row in df.iterrows():
        text = str(row[tcol]).strip()
        ctx  = str(row[ctx_col]).strip()
        ke_rows.append({
            "country":     "Kenya",
            "content_id":  str(row[id_col]),
            "creator":     sheet,
            "orientation": orientation,
            "context":     ctx,
            "text":        text,
        })

ke_df = pd.DataFrame(ke_rows)
print(f"  {len(ke_df)} snippets | prog={(ke_df.orientation=='progressive').sum()} reg={(ke_df.orientation=='regressive').sum()}")

# ── QA checks ────────────────────────────────────────────────────────────────
def qa(df, name):
    empty = (df["text"].isna() | (df["text"].str.strip() == "") | (df["text"] == "nan")).sum()
    too_long = (df["text"].str.len() > 5000).sum()
    print(f"QA {name}: {len(df)} rows | empty text={empty} | too_long={too_long}")
    assert empty == 0, f"Empty text found in {name}"

qa(ng_df, "Nigeria Content")
qa(ke_df, "Kenya Content")

# ── analyses to run ──────────────────────────────────────────────────────────
ANALYSES = ["themes", "sentiment", "emotion", "misogyny", "framing"]

def run_analyses(df, name):
    out = df.copy().reset_index(drop=True)
    texts = out["text"].tolist()
    for analysis in ANALYSES:
        print(f"\n  [{name}] → {analysis} ({len(texts)} rows)")
        res = eal.run(texts, analysis)
        if "id" in res.columns:
            res = res.drop(columns=["id"])
        res = res.add_prefix(f"{analysis}__")
        out = pd.concat([out, res], axis=1)
        print(f"    done. columns added: {list(res.columns)}")
    return out

print("\n" + "="*60)
print("RUNNING NIGERIA CONTENT EDA")
print("="*60)
ng_out = run_analyses(ng_df, "Nigeria Content")

print("\n" + "="*60)
print("RUNNING KENYA CONTENT EDA")
print("="*60)
ke_out = run_analyses(ke_df, "Kenya Content")

# ── serialise list columns for Excel ────────────────────────────────────────
def listify(df):
    out = df.copy()
    for c in out.columns:
        if out[c].apply(lambda x: isinstance(x, (list, dict))).any():
            out[c] = out[c].apply(
                lambda x: json.dumps(x, ensure_ascii=False)
                if isinstance(x, (list, dict)) else x
            )
    return out

ng_out.pipe(listify).to_excel(OUT_DIR / "Nigeria Content EDA.xlsx", index=False)
ke_out.pipe(listify).to_excel(OUT_DIR / "Kenya Content EDA.xlsx", index=False)
print(f"\nSaved Nigeria Content EDA.xlsx  ({len(ng_out)} rows)")
print(f"Saved Kenya Content EDA.xlsx    ({len(ke_out)} rows)")

# ── compute and print document-relevant stats ────────────────────────────────
print("\n" + "="*60)
print("VERIFIED THEME COUNTS FOR DOCUMENT")
print("="*60)

def parse_themes(t):
    try:
        if isinstance(t, list): return t
        return json.loads(str(t))
    except:
        return [str(t)] if t else []

ng_out["theme_list"] = ng_out["themes__themes"].apply(parse_themes)
ke_out["theme_list"] = ke_out["themes__themes"].apply(parse_themes)

def theme_report(df, country):
    n      = len(df)
    n_prog = (df.orientation == "progressive").sum()
    n_reg  = (df.orientation == "regressive").sum()

    # get all unique themes
    all_themes = sorted({t for tl in df["theme_list"] for t in tl})
    print(f"\n--- {country} Content: {n} snippets "
          f"(prog={n_prog}, reg={n_reg}) ---")
    print(f"All themes found: {all_themes}")

    rows = []
    for theme in all_themes:
        mask  = df["theme_list"].apply(lambda tl: theme in tl)
        total = mask.sum()
        prog  = (mask & (df.orientation == "progressive")).sum()
        reg   = (mask & (df.orientation == "regressive")).sum()
        per_c = {}
        for c in df.creator.unique():
            sub = df[df.creator == c]
            h   = sub["theme_list"].apply(lambda tl: theme in tl).sum()
            per_c[c] = (int(h), int(len(sub)))
        rows.append(dict(theme=theme, total=total, pct=total/n*100,
                         prog=prog, prog_pct=prog/n_prog*100,
                         reg=reg, reg_pct=reg/n_reg*100,
                         per_creator=per_c))

    for r in sorted(rows, key=lambda x: -x["total"]):
        print(f"\n  {r['theme']}: {r['total']}/{n} ({r['pct']:.1f}%)")
        print(f"    Progressive: {r['prog']}/{n_prog} ({r['prog_pct']:.1f}%)"
              f"   Regressive: {r['reg']}/{n_reg} ({r['reg_pct']:.1f}%)")
        for c, (h, tot) in r["per_creator"].items():
            print(f"      {c}: {h}/{tot} ({h/tot*100:.0f}%)")
    return rows

ng_stats = theme_report(ng_out, "NIGERIA")
ke_stats = theme_report(ke_out, "KENYA")

# ── write summary to text file ───────────────────────────────────────────────
summary_lines = []
summary_lines.append("VERIFIED LLM EDA THEME COUNTS — BOTH COUNTRIES")
summary_lines.append("Model: gpt-4o-mini | Source: Final xlsx datasets")
summary_lines.append("="*60)

for country, stats, df in [("NIGERIA", ng_stats, ng_out), ("KENYA", ke_stats, ke_out)]:
    n = len(df)
    n_prog = (df.orientation=="progressive").sum()
    n_reg  = (df.orientation=="regressive").sum()
    summary_lines.append(f"\n{country} CONTENT ({n} snippets, prog={n_prog}, reg={n_reg})")
    for r in sorted(stats, key=lambda x: -x["total"]):
        summary_lines.append(
            f"  {r['theme']}: {r['total']}/{n} ({r['pct']:.1f}%) | "
            f"prog={r['prog']}/{n_prog} ({r['prog_pct']:.1f}%) | "
            f"reg={r['reg']}/{n_reg} ({r['reg_pct']:.1f}%)"
        )
        for c, (h, tot) in r["per_creator"].items():
            summary_lines.append(f"    {c}: {h}/{tot} ({h/tot*100:.0f}%)")

(OUT_DIR / "theme-counts-summary.txt").write_text("\n".join(summary_lines))
print(f"\nSaved theme-counts-summary.txt")
print("\nAll done.")
