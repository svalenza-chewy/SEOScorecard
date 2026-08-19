import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/svalenza/Documents/CEO Ask";
const outputDir = `${root}/outputs/keyword_scorecard_workbook`;
const pageCsv = `${root}/outputs/page_type_heatmap_with_ranking_keywords.csv`;
const urlCsv = `${root}/outputs/url_keyword_summary.csv`;
const outputXlsx = `${outputDir}/chewy_seo_scorecard_with_ranking_keywords.xlsx`;

await fs.mkdir(outputDir, { recursive: true });

const pageCsvText = await fs.readFile(pageCsv, "utf8");
const urlCsvText = await fs.readFile(urlCsv, "utf8");
const pageRows = pageCsvText.trimEnd().split(/\r?\n/);
const urlRows = urlCsvText.trimEnd().split(/\r?\n/);

function colName(index) {
  let name = "";
  let n = index;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

const pageColCount = pageRows[0].split(",").length;
const urlColCount = urlRows[0].split(",").length;
const pageLastCol = colName(pageColCount);
const urlLastCol = colName(urlColCount);
const pageLastRow = pageRows.length;
const urlLastRow = urlRows.length;

const workbook = await Workbook.fromCSV(pageCsvText, {
  sheetName: "Page Type Heatmap",
});
await workbook.fromCSV(urlCsvText, { sheetName: "URL Keyword Summary" });

const notes = workbook.worksheets.add("Scoring Notes");
const notesRows = [
  ["Topic", "Decision"],
  ["New heatmap columns", "Ranking Keywords and Ranking Potential"],
  ["Keyword source", "Keywords_2026-08-18.xlsx, Expanded Keyword Performance sheet, rows 8+"],
  ["Join key", "Normalize Ranking URL from the keyword workbook to Full URL from Botify: https, lowercase host, path only, no query string, no trailing slash except homepage."],
  ["Main keyword selection", "Per Ranking URL, select Prime = Y first, then Is Highest Ranking = Y, then Highest Ranking = Y, then highest MSV, then best current rank."],
  ["Meta keyword alignment", "When a main keyword exists for a URL, Meta score gives 15 points for title keyword match and 10 points for meta-description keyword match. URLs without a Conductor target keep Meta V1 and are reported in coverage."],
  ["URL keyword alignment", "URL score now rewards clean URL shape, stable Chewy template IDs, readable slugs, and slug-to-main-keyword/topic match. Browse and Brand URLs are stricter: if the target includes a species or brand modifier, the slug must carry it to receive strong intent credit."],
  ["H1 decision", "H1 score uses the Botify evidence available now: H1 presence, single-H1 setup, descriptive H1 length, H1-to-main-keyword/topic alignment, title/H1 consistency, page-type support signals, and H1 duplication."],
  ["On-page content", "Not scored yet. Add as a separate future section when the upload includes body-copy depth, unique copy, duplicate content, freshness, authorship, moderation, and template content fields."],
  ["Index decision", "Index Readiness is scored from Botify proxy evidence: noindex/nofollow controls, self-referencing canonical, and sitemap compliance. Indexable URLs are expected in sitemap; noindex URLs are expected out of sitemap, so they are not penalized for sitemap exclusion."],
  ["Index caps", "For indexable URLs, missing self-canonical is capped at 60, missing sitemap compliance is capped at 70, and nofollow is capped at 85. Noindex URLs are scored as exclusion implementation rather than forced sitemap inclusion."],
  ["Index source gap", "This upload does not include direct Google indexed/not-indexed status, HTTP status, robots.txt allow/block, or canonical target URL, so the workbook labels this as readiness rather than confirmed indexation."],
  ["Links decision", "Links score weights unique inlinks most heavily, then crawl depth, internal outlinks, and breadcrumb/navigation signal. Thresholds are page-type-specific because PDPs, Browse pages, Brand pages, and Deals naturally have different discovery patterns."],
  ["Links caveat", "Botify outlink counts are quantity-only and appear heavily template-driven. Future scoring should split contextual body links, product-grid links, navigation links, broken links, redirects, and orphan state."],
  ["Speed decision", "Speed score averages only URLs with at least one measured Core Web Vitals metric. URLs with no LCP, CLS, or INP measurement are excluded from the Speed score rather than counted as 0. Standards: LCP good under 2.5s, needs improvement 2.5s to 4.0s, poor over 4.0s; INP good under 200ms, needs improvement 200ms to 500ms, poor over 500ms; CLS good under 0.1, needs improvement 0.1 to 0.25, poor over 0.25."],
  ["Speed coverage handling", "Speed Coverage % and Speed Scored URLs are reported separately as confidence indicators. Low coverage does not cap or lower the Speed score in this version."],
  ["Speed source gap", "This Botify export has LCP and CLS columns but no INP columns with data. INP coverage is shown as 0% until the upload includes INP from CrUX/PageSpeed/Botify, and missing INP is not treated as a 0."],
  ["Ranking Keywords score buckets", "Current rank 1-3 = 100, 4-10 = 85, 11-20 = 70, 21-50 = 45, 51-100 = 20, Did not rank = 0."],
  ["Ranking Keywords formula", "Current ranking health score: 75% average rank-bucket points + 15% top-10 share + 10% current-ranked share."],
  ["Ranking Potential formula", "Opportunity score where higher means more upside, not better health. It blends top-10 gap, quick-win keywords in positions 11-20, striking-distance keywords in positions 21-50, average Botify audience-keyword potential, and total Botify audience-keyword potential."],
  ["Ranking Potential source", "Botify audience-keyword potential uses: No. of Keywords for the URL To Achieve 90% Audience (by Country). Raw supporting metrics stay visible so page-type volume does not hide the denominator."],
  ["Overall score handling", "Keep Overall Technical as the controllable Botify score. Use Overall Incl Ranking Keywords only as an optional blended view."],
  ["Bad crawl exclusion", "Rows with no extracted title, meta description, H1, schema signal, canonical signal, or sitemap signal are treated as bad crawls and excluded from health-score averages while still reported in URL counts."],
  ["Customer Care", "Removed from the scorecard page-type rollup."],
  ["HVSP caveat", "Keyword workbook has more /sp/ URLs than Botify crawl rows, so show keyword-only /sp/ URLs as a crawl coverage gap."],
  ["Schema decision", "PDP prioritizes Product, Offer, and Breadcrumb schema; Rating and FAQ are small nice-to-have signals. Browse and HVF require Breadcrumb schema; FAQ is nice to have. HVSP, Top Picks, Brand, and Deals currently score schema from Breadcrumb presence. Editorial is set to 90 for Schema."],
  ["Schema/future content caveat", "Schema remains Botify-limited until exports include validation details and richer schema fields. On-page content will be scored later when exports include word count, unique copy, duplicate content, freshness, authorship, moderation, and template content fields."],
];
notes.getRange(`A1:B${notesRows.length}`).values = notesRows;

const page = workbook.worksheets.getItem("Page Type Heatmap");
const urlSheet = workbook.worksheets.getItem("URL Keyword Summary");

const pageHeaders = [
  "Page Type",
  "Botify URLs",
  "Scorable URLs",
  "Excluded Bad Crawl URLs",
  "Excluded Bad Crawl %",
  "Keyword URLs",
  "Keyword URLs Matched",
  "Keyword Inventory Coverage %",
  "Keyword Rows",
  "Distinct Keywords",
  "Avg Current Rank",
  "Top 3 Share %",
  "Top 10 Share %",
  "Top 20 Share %",
  "Lost / Unranked Share %",
  "Ranking Keywords",
  "Keywords Not Top 10 %",
  "Quick Win 11-20 %",
  "Striking Distance 21-50 %",
  "Audience Potential URLs",
  "Audience Potential Coverage %",
  "Avg Botify Audience Potential",
  "Total Botify Audience Potential",
  "Ranking Potential",
  "Main Keyword Target Coverage %",
  "H1 Present %",
  "Single H1 %",
  "H1 Main Keyword Match %",
  "Title Main Keyword Match %",
  "Meta Main Keyword Match %",
  "URL Main Keyword Match %",
  "Index Controls Clear %",
  "Noindex %",
  "Self-Canonical %",
  "Sitemap Compliance %",
  "Canonical/Sitemap Compliance %",
  "Median Inlinks",
  "Inlink Target Met %",
  "Median Outlinks",
  "Outlink Target Met %",
  "Crawl Depth OK %",
  "Link Breadcrumb %",
  "LCP Measured %",
  "LCP Good %",
  "CLS Measured %",
  "CLS Good %",
  "INP Measured %",
  "INP Good %",
  "Speed Measured Health",
  "Speed Scored URLs",
  "Speed Coverage %",
  "Meta",
  "Schema",
  "URL",
  "H1",
  "Index Readiness",
  "Links",
  "Speed",
  "Overall Excl Ranking Keywords",
  "Overall Incl Ranking Keywords",
];
page.getRange(`A1:${pageLastCol}1`).values = [pageHeaders];

const urlHeaders = [
  "Normalized URL",
  "Page Type",
  "Matched Botify",
  "Keyword Rows",
  "Distinct Ranking Keywords",
  "Current Ranked Keywords",
  "Lost / Unranked Keywords",
  "Avg Current Rank",
  "Top 3 Keywords",
  "Top 10 Keywords",
  "Top 20 Keywords",
  "Top 50 Keywords",
  "Top 100 Keywords",
  "Ranking Keyword Score",
  "Keywords Not Top 10",
  "Quick Win 11-20 Keywords",
  "Striking Distance 21-50 Keywords",
  "Botify Audience Keywords to 90%",
  "Ranking Potential Score",
  "Main Keyword",
  "Title Matches Main Keyword",
  "Meta Matches Main Keyword",
  "H1 Matches Main Keyword",
  "URL Matches Main Keyword",
  "Top Keyword Groups",
];
urlSheet.getRange(`A1:${urlLastCol}1`).values = [urlHeaders];

for (const sheet of [page, urlSheet, notes]) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
}

page.getRange(`A1:${pageLastCol}1`).format = {
  fill: "#12345B",
  font: { bold: true, color: "#FFFFFF" },
};
page.getRange(`A1:${pageLastCol}${pageLastRow}`).format.borders = {
  insideHorizontal: { style: "thin", color: "#DFE7F1" },
  bottom: { style: "thin", color: "#C8D6E6" },
};
page.getRange("A:A").format.columnWidth = 18;
page.getRange("B:G").format.columnWidth = 14;
page.getRange(`H:${pageLastCol}`).format.columnWidth = 13;
page.getRange(`A1:${pageLastCol}${pageLastRow}`).format.wrapText = true;
page.getRange(`A1:${pageLastCol}1`).format.rowHeight = 42;
page.getRange(`B2:G${pageLastRow}`).format.numberFormat = "#,##0";
page.getRange(`H2:${pageLastCol}${pageLastRow}`).format.numberFormat = "0.0";

const pageHeaderColumns = new Map(
  pageHeaders.map((header, index) => [header, colName(index + 1)]),
);

function pageColumnRange(header) {
  const column = pageHeaderColumns.get(header);
  return column ? `${column}2:${column}${pageLastRow}` : null;
}

function formatPageColumn(header, numberFormat) {
  const address = pageColumnRange(header);
  if (address) page.getRange(address).format.numberFormat = numberFormat;
}

for (const header of [
  "Botify URLs",
  "Scorable URLs",
  "Excluded Bad Crawl URLs",
  "Keyword URLs",
  "Keyword URLs Matched",
  "Keyword Rows",
  "Distinct Keywords",
  "Audience Potential URLs",
  "Total Botify Audience Potential",
  "Speed Scored URLs",
]) {
  formatPageColumn(header, "#,##0");
}

function applyHealthScale(address) {
  if (!address) return;
  const range = page.getRange(address);
  range.conditionalFormats.add("cellIs", {
    operator: "lessThan",
    formula: 70,
    format: { fill: "#F9E1E4", font: { color: "#B02A2F", bold: true } },
  });
  range.conditionalFormats.add("cellIs", {
    operator: "between",
    formula: [70, 84.999],
    format: { fill: "#FFF0D1", font: { color: "#A46400", bold: true } },
  });
  range.conditionalFormats.add("cellIs", {
    operator: "greaterThanOrEqual",
    formula: 85,
    format: { fill: "#DFF2E7", font: { color: "#197A49", bold: true } },
  });
}

function applyOpportunityScale(address) {
  if (!address) return;
  const range = page.getRange(address);
  range.conditionalFormats.add("cellIs", {
    operator: "lessThan",
    formula: 55,
    format: { fill: "#EEF2F7", font: { color: "#526173", bold: true } },
  });
  range.conditionalFormats.add("cellIs", {
    operator: "between",
    formula: [55, 74.999],
    format: { fill: "#E6F0FF", font: { color: "#2459A6", bold: true } },
  });
  range.conditionalFormats.add("cellIs", {
    operator: "greaterThanOrEqual",
    formula: 75,
    format: { fill: "#D9E8FF", font: { color: "#0F4C9A", bold: true } },
  });
}

for (const header of [
  "Ranking Keywords",
  "Meta",
  "Schema",
  "URL",
  "H1",
  "Index Readiness",
  "Links",
  "Speed",
  "Overall Excl Ranking Keywords",
  "Overall Incl Ranking Keywords",
]) {
  applyHealthScale(pageColumnRange(header));
}
applyOpportunityScale(pageColumnRange("Ranking Potential"));

page.tables.add(`A1:${pageLastCol}${pageLastRow}`, true, "PageTypeHeatmap");

urlSheet.getRange(`A1:${urlLastCol}1`).format = {
  fill: "#12345B",
  font: { bold: true, color: "#FFFFFF" },
};
urlSheet.getRange("A:A").format.columnWidth = 72;
urlSheet.getRange(`B:${urlLastCol}`).format.columnWidth = 18;
urlSheet.getRange(`A1:${urlLastCol}1`).format.wrapText = true;
urlSheet.getRange(`A1:${urlLastCol}1`).format.rowHeight = 42;
urlSheet.getRange(`D2:M${urlLastRow}`).format.numberFormat = "#,##0";
urlSheet.getRange(`N2:N${urlLastRow}`).format.numberFormat = "0.0";
urlSheet.getRange(`O2:Q${urlLastRow}`).format.numberFormat = "#,##0";
urlSheet.getRange(`R2:S${urlLastRow}`).format.numberFormat = "0.0";
urlSheet.tables.add(`A1:${urlLastCol}${urlLastRow}`, true, "UrlKeywordSummary");

notes.getRange("A1:B1").format = {
  fill: "#12345B",
  font: { bold: true, color: "#FFFFFF" },
};
notes.getRange("A:A").format.columnWidth = 24;
notes.getRange("B:B").format.columnWidth = 118;
notes.getRange(`A1:B${notesRows.length}`).format.wrapText = true;
notes.getRange(`A1:B${notesRows.length}`).format.borders = {
  insideHorizontal: { style: "thin", color: "#DFE7F1" },
  bottom: { style: "thin", color: "#C8D6E6" },
};

const preview = await workbook.render({
  sheetName: "Page Type Heatmap",
  range: `A1:${pageLastCol}${pageLastRow}`,
  scale: 1,
  format: "png",
});
await fs.writeFile(
  `${outputDir}/page_type_heatmap_preview.png`,
  new Uint8Array(await preview.arrayBuffer()),
);

for (const [sheetName, range, fileName] of [
  ["URL Keyword Summary", `A1:${urlLastCol}24`, "url_keyword_summary_preview.png"],
  ["Scoring Notes", "A1:B15", "scoring_notes_preview.png"],
]) {
  const rendered = await workbook.render({
    sheetName,
    range: sheetName === "Scoring Notes" ? `A1:B${notesRows.length}` : range,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    `${outputDir}/${fileName}`,
    new Uint8Array(await rendered.arrayBuffer()),
  );
}

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errorScan.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputXlsx);
console.log(outputXlsx);
