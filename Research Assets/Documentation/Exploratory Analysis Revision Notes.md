# Exploratory Analysis Revision Notes

Date: 2026-05-17  
Project: Gates Manfluencer Research (Kenya and Nigeria)

## Why this revision was made

This revision was prepared to address manager feedback requesting:

- clearer methods and definitions,
- explicit context for regressive and progressive coding,
- country-by-country reporting rather than direct country comparisons,
- normalization context tied to purposive creator selection,
- and an engagement and reach layer with transparency about missing metrics.

## Current final file baselines used

- `Kenya/Content Analysis/Kenya Content Analysis Final.xlsx` (394 rows)
- `Kenya/Audience Analysis/Kenya Audience Analysis Final.xlsx` (412 rows)
- `Nigeria/Content Analysis/Nigeria Content Analysis Final.xlsx` (310 rows)
- `Nigeria/Audience Analysis/Nigeria Audience Analysis Final.xlsx` (417 rows)

These are the primary manager-facing final workbooks for this exploratory package.

## Orientation context used for interpretation

Creator selection is purposive, not random. The sample intentionally includes both regressive-leaning and progressive-leaning creators.

- Kenya content final rows: 188 regressive, 206 progressive
- Kenya audience final rows: 210 regressive, 202 progressive
- Nigeria content final rows: 173 regressive, 137 progressive
- Nigeria audience final rows: 227 regressive, 190 progressive

Implication: percentages in this package reflect selected corpus composition and should not be treated as national prevalence.

## Engagement and reach availability note

A metric-field audit on raw audience files shows that likes and replies are broadly available, retweets are mainly present for X-origin files, and view or impression fields are often missing.

- Kenya raw audience files scanned: 10
- Nigeria raw audience files scanned: 31

Because of this coverage variation, engagement analysis should be run on available metrics only and labeled by platform and metric coverage.

## Deliverables updated in this pass

- `Research Assets/Deliverables/LLM Exploratory Audience Findings Revised.docx`
- `Research Assets/Deliverables/LLM Exploratory Content Findings Revised.docx`

The revised documents include:

- explicit methods and definitions,
- country-separated findings,
- normalization and denominator context,
- engagement and reach caveats,
- and project visuals pulled from current repository figures.

## Reproducibility

Document generation scripts:

- `Nigeria/scripts/build-manager-ready-exploratory-doc.py`
- `Nigeria/scripts/build-manager-exploratory-docs.py`

This script rebuilds the revised exploratory `.docx` and associated generated figures under:

- `Research Assets/Deliverables/Figures/`
