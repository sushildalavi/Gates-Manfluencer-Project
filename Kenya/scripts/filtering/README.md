# Kenya audience comment filtering

Kenya-only scripts for the Gates Manfluencer project. Applies keyword- and
linguistic-signal-based filters to raw audience engagement workbooks and
produces per-piece kept workbooks for downstream coding.

(Nigeria filtering lives under `Nigeria/scripts/` in this repo.)

## Layout

`Kenya/scripts/filtering/` contains the filtering scripts and report notes:

- `kenya_audience_filter_pipeline.py`       Kenya corpus-level filter (v1/v2)
- `run_piecewise_audience_filter.py`        Per-piece Kenya runner (v1/v2)
- `audience_relevance_filter_piecewise.py`  Single-file filter (platform-agnostic)
- `filter_kenya_comments.py`      Keyword-match filter (conservative / expanded)
- `create_kenya_top4_relevant_workbook.py`  Builds top-4 most-relevant workbook
- `reports/kenya_top4_scope_audit.md`        Scope audit notes

## Running the filters

All scripts resolve their paths relative to this folder (via `__file__`),
so you can run them from anywhere without `cd`-ing first.

### Per-piece Kenya piecewise filter (v1 + v2)

```bash
python Kenya/scripts/filtering/run_piecewise_audience_filter.py
```

Reads every XLSX under `inputs/Kenya/` listed in
`kenya_audience_filter_pipeline.KENYA_FILES`, runs both v1 and v2 modes, and
writes per-piece `v1/` and `v2/` folders under
`outputs/Kenya/piecewise_filter_output/` as Excel workbooks.

### Corpus-level Kenya filter

```bash
python Kenya/scripts/filtering/kenya_audience_filter_pipeline.py --mode v2
# or override paths:
python Kenya/scripts/filtering/kenya_audience_filter_pipeline.py \
  --mode v2 \
  --input-dir  ./inputs/Kenya \
  --keyword-file "./keywords/NLC Proposed keywords.xlsx" \
  --output-dir ./outputs/Kenya/filter_output_v2
```

### Keyword-match filter (Kenya, conservative or expanded)

```bash
python Kenya/scripts/filtering/filter_kenya_comments.py --match-mode conservative
python Kenya/scripts/filtering/filter_kenya_comments.py --match-mode expanded
```

Outputs go to `outputs/Kenya/filtered_output/` or
`outputs/Kenya/filtered_output_expanded_variants/` respectively.

### Single-file piecewise filter (any platform)

```bash
python Kenya/scripts/filtering/audience_relevance_filter_piecewise.py \
  --input-file ./inputs/Kenya/"Full Tweet Stay away from vulgar women.xlsx" \
  --mode v2 \
  --output-dir ./outputs/Kenya/ad_hoc/vulgar_women
```

### Top-4 most-relevant workbook (Kenya)

```bash
python Kenya/scripts/filtering/create_kenya_top4_relevant_workbook.py
```

Reads the four priority pieces from `outputs/Kenya/filtered_output/` and
writes `outputs/Kenya/Kenya_top4_most_relevant_comments.xlsx`.

## Filter modes, at a glance

| Mode | Used by | Criteria |
|---|---|---|
| `v1` (piecewise) | `run_piecewise_audience_filter`, `audience_relevance_filter_piecewise` | Word-count floor + basic linguistic-signal hints. More permissive. |
| `v2` (piecewise) | same | Stricter. ≥ 5 words + (≥ 8 words OR at least one hint group + meaningful structure). |
| `conservative` (keyword) | `filter_kenya_comments` | Case-insensitive exact-phrase matches against the Kenya/Nigeria sheets of the keyword workbook. |
| `expanded` (keyword) | `filter_kenya_comments` | Conservative matches + morphological variants (lowercased stems, common plural/verb forms). |
| `conservative_strict` (Nigeria) | `filter_kenya_comments` + exclusion list | Conservative keyword matches minus any rows matching `keywords/conservative_excluded_keywords.txt`. |

## Dependencies

```bash
pip install pandas openpyxl
```

All scripts write their outputs to subdirectories under `outputs/` and
never mutate `inputs/`, so re-runs are safe.
