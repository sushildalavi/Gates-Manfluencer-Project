# Nigeria Workspace Handover Guide

This folder is the Nigeria-side production workspace for the Manfluencer project.

## Folder map

- `Audience Analysis/`
  - `Audience Comments - Raw/`: source comment pulls.
  - `Audience Comments - Complete/`: cleaned/deduped comment workbooks.
  - `Audience Comments - Final/`: final scoped audience files used for reporting.
  - `Translated/`: English translated audience outputs.
  - `Exploratory/`: audience EDA outputs, figures, and summary workbooks.
  - `Nigeria Audience Analysis Final.xlsx`: manager-facing final audience workbook.

- `Content Analysis/`
  - `Content - Raw/`: source content assets (tweets, transcripts, captions, audio references).
  - `Content - Final/`: final per-creator coded content workbooks.
  - `Translated/`: English translated content outputs (script target).
  - `Exploratory/`: content EDA outputs and figures (script target).
  - `Nigeria Content Analysis Final.xlsx`: manager-facing final content workbook.

- `Notebooks/`
  - `Data Acquisition Pipeline.ipynb`: scraping + transcription flow.
  - `Audience Comments.ipynb`: audience cleaning/scope flow.
  - `Content Analysis.ipynb`: Nigeria content analysis workflow.
  - `Exploratory Analysis.ipynb`: exploratory analyses that feed report findings.

- `scripts/` (kept active only)
  - Acquisition/transcription: `transcribe_videos.py`, `finalize_transcripts_gemini.py`, `align_transcripts_with_captions.py`, `fix_speaker_labels.py`, `transcripts_utils.py`
  - Scraping: `scrape_all_creator_tweets.py`, `scrape_x_creator_tweets.py`, `scrape_x_tweet_replies.py`, `scrape_youtube_comments.py`
  - Processing/EDA docs: `filter_scope_relevant_comments.py`, `translate_to_english_pipeline.py`, `exploratory_analysis_lib.py`, `build-manager-ready-exploratory-doc.py`, `build-research-handover-doc.py`

## Practical sequence for a new handover owner

1. Use `Notebooks/Data Acquisition Pipeline.ipynb` for ingestion/transcription updates.
2. Use `Notebooks/Audience Comments.ipynb` for raw-to-complete/final audience pipeline.
3. Validate final workbooks:
   - `Audience Analysis/Nigeria Audience Analysis Final.xlsx`
   - `Content Analysis/Nigeria Content Analysis Final.xlsx`
4. Run translations via `scripts/translate_to_english_pipeline.py`.
5. Run `Notebooks/Exploratory Analysis.ipynb` for EDA outputs used in report narratives.

## Notes

- Kenya workspace is intentionally separate and untouched by Nigeria pipeline changes.
- `Translated/` and `Exploratory/` under Content Analysis are kept as stable output targets for scripts and notebook runs.
