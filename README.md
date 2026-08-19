# Chewy SEO Scorecard

This project builds a local executive SEO scorecard from the current Botify crawl export and Conductor keyword workbook. It produces CSV rollups, an Excel workbook, and a local browser UI for reviewing page-type health, keyword performance, ranking opportunity, and source coverage.

## Current Outputs

- `outputs/page_type_heatmap_with_ranking_keywords.csv`: page-type scorecard rollup used by the UI.
- `outputs/url_keyword_summary.csv`: URL-level keyword drilldown.
- `outputs/page_type_issue_summary.csv`: aggregated page-type issue drivers used by the UI.
- `outputs/keyword_scorecard_workbook/chewy_seo_scorecard_with_ranking_keywords.xlsx`: formatted workbook version.
- `web/`: static local UI for reviewing the scorecard.

## Quick Start

Generate the scorecard data:

```bash
/Users/svalenza/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_scorecard_data.py
```

Generate the formatted workbook:

```bash
/Users/svalenza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node scripts/build_keyword_scorecard_workbook.mjs
```

Run the local UI:

```bash
python3 -m http.server 8787
```

Open:

```text
http://127.0.0.1:8787/web/
```

The UI currently reads `outputs/page_type_heatmap_with_ranking_keywords.csv`. It also has a CSV upload control for testing a newly generated scorecard CSV. Direct Botify ZIP upload/parsing is not implemented in the UI yet.

## Public GitHub Pages URL

The shareable team dashboard publishes from the `site/` folder through GitHub Actions:

```text
https://svalenza-chewy.github.io/SEOScorecard/
```

When the scorecard data changes, regenerate the local outputs, then rebuild the lightweight Pages artifact:

```bash
python3 scripts/build_public_site.py
```

The public artifact includes the dashboard, page-type rollups, issue summaries, and sampled URL drilldown CSVs. The full local `outputs/issue_details/` exports are intentionally ignored because they are too large for GitHub Pages.

## Source Inputs

The current generator is configured for these local files:

- Botify export: `/Users/svalenza/Downloads/export_2072255_4294835_www.chewy.com_2026-08-18_15-07.zip`
- Botify CSV inside zip: `export_2072255_4294835_www.chewy.com_2026-08-18_15-07.csv`
- Conductor keyword workbook: `/Users/svalenza/Downloads/Keywords_2026-08-18.xlsx`
- Keyword sheet: `Expanded Keyword Performance`, starting at row 8

If a new upload arrives, update the constants at the top of `scripts/build_scorecard_data.py`, then rerun the data and workbook generation commands.

## Page Types

The scorecard currently rolls up these page types:

- `Browse`: `/b/`
- `HVF`: `/f/`
- `HVSP`: `/sp/`
- `Top Picks`: `/best/`
- `PDP`: URLs containing `/dp/`
- `Brand`: `/brands` and `/brands/`
- `Editorial`: `/education/`, `/bechewy/`, and `/petmd/`
- `Deals`: `/deals/`

`Customer Care` and `Media` are intentionally removed from the scorecard for now.

## Bad Crawl Exclusion

Rows are excluded from health-score averages when they look like bad crawls. A row is treated as a bad crawl when it has no extracted title, meta description, H1, schema signal, canonical signal, or sitemap signal.

Excluded rows still appear in URL counts through `Excluded Bad Crawl URLs` and `Excluded Bad Crawl %`, but they do not pull down page-type scores.

## Score Bands

Health scores use this executive color scale:

- `85-100`: Strong
- `70-84`: Needs attention
- `0-69`: At risk

`Ranking Potential` uses a separate opportunity color scale because higher means more upside, not better health.

## Overall Scores

`Overall Excl Ranking Keywords` is the main controllable technical score. It averages the implemented technical health sections that have a value:

- Meta
- Schema
- URL
- H1
- Index Readiness
- Links
- Speed

`Overall Incl Ranking Keywords` is an optional blended view. It adds `Ranking Keywords` to the average when keyword data is available.

`Ranking Potential` is not included in either overall score because it is a prioritization signal, not a health score.

## Section Scoring

### Meta

Meta scores title and meta-description quality from Botify evidence.

Title logic:

- Title exists: 25 points
- Ideal title length, 30-65 characters: 15 points
- Acceptable title length, 20-75 characters: 8 points
- Unique title: 10 points
- Mild duplicate title: 5 points

Meta-description logic:

- Meta description exists: 25 points
- Ideal length, 70-170 characters: 15 points
- Acceptable length, 50-200 characters: 8 points
- Unique description: 10 points
- Mild duplicate description: 5 points

When a Conductor main keyword exists for the URL, the base meta score is weighted to 75%, then keyword alignment can add:

- Title includes or partially matches main keyword: up to 15 points
- Meta description includes or partially matches main keyword: up to 10 points

### Schema

Schema scoring is page-type specific.

PDP:

- Product schema: 35 points
- Offer schema: 30 points
- Breadcrumb schema: 25 points
- Rating schema: 5 points
- FAQ schema: 5 points

Browse and HVF:

- Breadcrumb schema: 80 points
- FAQ schema: 20 points

Editorial:

- Fixed at 90 for now, based on the agreed assumption that schema exists but article schema placement is not implemented correctly in the head.

HVSP, Top Picks, Brand, and Deals:

- Breadcrumb schema present: 100 points
- Breadcrumb schema missing: 0 points

### URL

URL scoring rewards clean, stable, descriptive URLs that align to the target keyword or page topic.

Signals:

- Recognized page type: 25 points
- Clean normalized URL with lowercase path and no query or fragment: 15 points
- Stable Chewy template ID pattern: 10 points
- Readable slug: up to 20 points
- Slug-to-main-keyword or topic match: up to 25 points
- Reasonable path length and depth: 5 points

Browse and Brand URLs are stricter. If the main keyword includes a species or brand modifier, the slug must carry that modifier to receive strong intent credit. For example, a Browse URL targeting `dry dog food` should not be treated as fully strong if the slug only says `dry food`.

### H1

H1 scoring uses the Botify H1 field plus keyword and template support signals.

Signals:

- H1 present: 25 points
- Single H1: 10 points, or 5 points when multiple H1s exist
- Descriptive H1 token count: up to 10 points
- H1 alignment to the Conductor main keyword or title topic: up to 25 points
- Title/H1 consistency: up to 10 points
- Page-type support signals, such as product, review, rating, FAQ, offer, or Botify audience signal: up to 15 points depending on page type
- Low H1 duplication: up to 5 points

On-page body content is not scored yet. Future scoring should add body-copy depth, uniqueness, duplicate content, freshness, authorship, moderation, and template content fields.

### Index Readiness

Index Readiness uses Botify proxy fields:

- `X-Robots-Tags: Noindex`
- `X-Robots-Tags: Nofollow`
- `Canonical Points to Self`
- `In Sitemaps`

For indexable URLs:

- Indexable baseline: 20 points
- No nofollow: 10 points
- Self-referencing canonical: 40 points
- Sitemap compliance: 30 points

Caps for indexable URLs:

- Missing self-canonical caps the score at 60
- Missing sitemap compliance caps the score at 70
- Nofollow caps the score at 85

For noindex URLs:

- The score checks whether exclusion is implemented cleanly.
- Noindex URLs are expected to be out of the sitemap.
- Noindex URLs are not penalized for sitemap exclusion.

This section is called `Index Readiness` because the upload does not include direct Google indexed/not-indexed status, HTTP status, robots.txt allow/block, or canonical target URL.

### Links

Links scoring uses page-type-specific thresholds because discovery expectations differ by template.

Signals:

- Unique inlinks: up to 45 points
- Crawl depth: up to 25 points
- Unique internal outlinks: up to 20 points
- Breadcrumb/navigation signal: 10 points

Caps:

- Zero unique inlinks caps Links at 60
- Zero unique outlinks caps Links at 75

Current limitation: Botify outlink counts are quantity-only and appear heavily template-driven. Future scoring should split contextual body links, product-grid links, navigation links, broken links, redirects, and orphan state.

### Speed

Speed averages only URLs with at least one measured Core Web Vitals metric. URLs with no LCP, CLS, or INP measurement are excluded from the Speed score rather than counted as 0.

Metric points:

- Good: 100
- Needs improvement: 70
- Poor: 0

Thresholds:

- LCP good: under `2.5s`
- LCP needs improvement: `2.5s` to `4.0s`
- LCP poor: over `4.0s`
- INP good: under `200ms`
- INP needs improvement: `200ms` to `500ms`
- INP poor: over `500ms`
- CLS good: under `0.1`
- CLS needs improvement: `0.1` to `0.25`
- CLS poor: over `0.25`

The current Botify export has LCP and CLS fields but no INP field. INP coverage is shown as 0% until a future upload includes INP from CrUX, PageSpeed, or Botify.

`Speed Coverage %` and `Speed Scored URLs` are confidence indicators. Low coverage does not cap or lower the Speed score in the current model.

### Ranking Keywords

`Ranking Keywords` is current organic keyword health from the Conductor upload.

Rank buckets:

- Position 1-3: 100
- Position 4-10: 85
- Position 11-20: 70
- Position 21-50: 45
- Position 51-100: 20
- Did not rank: 0

Page-type formula:

```text
75% average rank-bucket points
+ 15% top-10 share
+ 10% current-ranked share
```

Main keyword selection per URL:

1. `Prime = Y`
2. `Is Highest Ranking = Y`
3. `Highest Ranking = Y`
4. Highest MSV
5. Best current rank

The selected main keyword is used by Meta, URL, and H1 alignment scoring.

### Ranking Potential

`Ranking Potential` is an opportunity score. Higher means more upside, not better health.

It blends:

- Top-10 gap
- Quick-win keywords in positions 11-20
- Striking-distance keywords in positions 21-50
- Average Botify audience-keyword potential
- Total Botify audience-keyword potential

The Botify audience-potential source field is:

```text
No. of Keywords for the URL To Achieve 90% Audience (by Country)
```

This score should be used for prioritization and planning, not for grading technical quality.

## Local UI

The UI has four views:

- `Heatmap`: executive page-type score table
- `Opportunity`: selected page-type metric health bars, page-type details, and aggregated issue drivers
- `URL List`: URL-level drilldown for a selected Opportunity issue, with evidence columns and CSV export
- `Weights`: top-level metric contribution and section-level scoring decisions
- `Inputs`: source readiness and coverage summary

Clicking a page type on the Heatmap tab opens the Opportunity tab for that page type. The issue section lists aggregated reasons the page type did not score 100 by category, with affected URL counts and percentages. Clicking an issue opens the URL List tab with the affected URL evidence. Large issue files preview the first 500 rows in the browser and expose the full generated CSV for export.

The UI currently supports:

- Page-type filter
- Score-band filter
- Current CSV reload
- Manual scorecard CSV upload
- Click-through issue drilldowns from Opportunity to URL List

The UI does not yet parse the raw Botify ZIP or Conductor workbook directly. That should be a future backend step.

## Aggregated Issue Drivers

`outputs/page_type_issue_summary.csv` is generated from the same scoring inputs as the heatmap. Each row contains:

- Page type
- Scorecard category
- Issue
- Affected URL count
- Eligible URL count
- Affected percent
- Severity
- Evidence
- Recommended action
- Detail CSV path

`outputs/issue_details/` contains one generated CSV per issue callout. Each file contains the affected URLs plus supporting fields such as title, meta description, H1, schema flags, index signals, link counts, speed metrics, keyword data, and recommendation text depending on the issue category.

The issue summary uses aggregate counts only. It does not expose URL-level examples in the UI.

## Known Gaps

- No direct Botify ZIP upload flow in the UI.
- No direct Conductor workbook upload flow in the UI.
- No direct indexed/not-indexed status from Google.
- No HTTP status, robots.txt allow/block, or canonical target URL field in the current export.
- No INP data in the current Botify export.
- No on-page body content scoring yet.
- No Media column by design.
- No Customer Care page type by design.

## Recommended Next Enhancements

- Add a local backend endpoint for Botify ZIP and Conductor workbook upload.
- Persist uploaded files and generated scorecard runs by date.
- Add URL-level drilldown in the UI.
- Add issue grouping by rule, owner, and page type.
- Add direct export from the UI to CSV or workbook.
- Add on-page content fields once the source export includes body-copy evidence.
