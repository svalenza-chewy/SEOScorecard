const DATA_URL = "../outputs/page_type_heatmap_with_ranking_keywords.csv";
const ISSUE_URL = "../outputs/page_type_issue_summary.csv";

const scoreColumns = [
  ["meta", "Meta"],
  ["schema", "Schema"],
  ["url", "URL"],
  ["h1", "H1"],
  ["index_readiness", "Index"],
  ["links", "Links"],
  ["speed", "Speed"],
  ["ranking_potential", "Ranking Potential"],
  ["overall_proxy_excl_ranking_keywords", "Overall Tech"],
];

const healthBarColumns = scoreColumns.filter(([key]) => key !== "ranking_potential");

const state = {
  rows: [],
  issues: [],
  urlDetails: [],
  activeIssue: null,
  selected: null,
  view: "heatmap",
  pageType: "all",
  band: "all",
  urlSearch: "",
};

const els = {
  refreshButton: document.querySelector("#refreshButton"),
  csvInput: document.querySelector("#csvInput"),
  pageTypeFilter: document.querySelector("#pageTypeFilter"),
  bandFilter: document.querySelector("#bandFilter"),
  statusBanner: document.querySelector("#statusBanner"),
  scoreTableBody: document.querySelector("#scoreTable tbody"),
  siteOverall: document.querySelector("#siteOverall"),
  siteOverallBand: document.querySelector("#siteOverallBand"),
  scorableUrls: document.querySelector("#scorableUrls"),
  rankingKeywordsAvg: document.querySelector("#rankingKeywordsAvg"),
  top10ShareAvg: document.querySelector("#top10ShareAvg"),
  opportunityPageTypes: document.querySelector("#opportunityPageTypes"),
  rankingBars: document.querySelector("#rankingBars"),
  detailTitle: document.querySelector("#detailTitle"),
  detailMetrics: document.querySelector("#detailMetrics"),
  issueHeading: document.querySelector("#issueHeading"),
  issueGrid: document.querySelector("#issueGrid"),
  urlListTitle: document.querySelector("#urlListTitle"),
  urlListMeta: document.querySelector("#urlListMeta"),
  urlListStatus: document.querySelector("#urlListStatus"),
  urlSearchInput: document.querySelector("#urlSearchInput"),
  urlDownloadLink: document.querySelector("#urlDownloadLink"),
  urlDetailTable: document.querySelector("#urlDetailTable"),
  urlEmptyState: document.querySelector("#urlEmptyState"),
  backToOpportunityButton: document.querySelector("#backToOpportunityButton"),
  weightsGrid: document.querySelector("#weightsGrid"),
  inputGrid: document.querySelector("#inputGrid"),
};

const detailPreviewLimit = 500;

function dataUrl(path) {
  if (!path) return "";
  const value = path.trim();
  if (/^(https?:|blob:|data:)/i.test(value)) return value;
  if (value.startsWith("/outputs/")) return `..${value}`;
  return value;
}

function cacheBust(path) {
  const url = dataUrl(path);
  return `${url}${url.includes("?") ? "&" : "?"}ts=${Date.now()}`;
}

const categoryOrder = [
  "Meta",
  "Schema",
  "URL",
  "H1",
  "Index",
  "Links",
  "Speed",
  "Ranking Keywords",
  "Ranking Potential",
];

const weightModels = [
  {
    name: "Meta",
    contribution: "14.3%",
    decision: "Measures title and meta-description quality, with Conductor keyword alignment layered in when a main keyword exists.",
    rules: [
      "Title exists: 25 pts; title length: 15 pts ideal or 8 pts acceptable; uniqueness: 10 pts unique or 5 pts mild duplicate.",
      "Meta description exists: 25 pts; description length: 15 pts ideal or 8 pts acceptable; uniqueness: 10 pts unique or 5 pts mild duplicate.",
      "When a main keyword exists, the base meta score is weighted to 75%, then title keyword alignment can add up to 15 pts and meta-description alignment can add up to 10 pts.",
    ],
    caveat: "Main keyword comes from the Conductor upload and is selected per URL before scoring.",
  },
  {
    name: "Schema",
    contribution: "14.3%",
    decision: "Uses page-type-specific schema expectations instead of one generic structured-data rule.",
    rules: [
      "PDP: Product 35 pts, Offer 30 pts, Breadcrumb 25 pts, Rating 5 pts, FAQ 5 pts.",
      "Browse and HVF: Breadcrumb 80 pts, FAQ 20 pts.",
      "Editorial is fixed at 90 for now because schema exists, but article schema placement is not implemented correctly in the head.",
      "HVSP, Top Picks, Brand, and Deals: Breadcrumb present scores 100; missing breadcrumb scores 0.",
    ],
    caveat: "Future exports should include richer schema validation details, not just existence signals.",
  },
  {
    name: "URL",
    contribution: "14.3%",
    decision: "Rewards clean, stable, readable URLs that carry the page topic or main keyword.",
    rules: [
      "Recognized page type: 25 pts.",
      "Clean normalized URL with lowercase path and no query or fragment: 15 pts.",
      "Stable Chewy template ID pattern: 10 pts.",
      "Readable slug: up to 20 pts.",
      "Slug-to-main-keyword or topic match: up to 25 pts.",
      "Reasonable path length and depth: 5 pts.",
    ],
    caveat: "Browse and Brand are stricter: species or brand modifiers in the target keyword should appear in the slug.",
  },
  {
    name: "H1",
    contribution: "14.3%",
    decision: "Scores whether the page has a clear primary heading and whether it aligns to the target topic.",
    rules: [
      "H1 present: 25 pts.",
      "Single H1: 10 pts; multiple H1s receive 5 pts.",
      "Descriptive H1 token count: up to 10 pts.",
      "H1-to-main-keyword or title-topic alignment: up to 25 pts.",
      "Title/H1 consistency: up to 10 pts.",
      "Page-type support signals such as product, review, rating, FAQ, offer, or Botify audience signal: up to 15 pts.",
      "Low H1 duplication: up to 5 pts.",
    ],
    caveat: "On-page body content is intentionally not scored yet until body-copy data is available.",
  },
  {
    name: "Index",
    contribution: "14.3%",
    decision: "Scores index-readiness from Botify proxy fields rather than confirmed Google indexation.",
    rules: [
      "Indexable URL baseline: 20 pts.",
      "No nofollow: 10 pts.",
      "Self-referencing canonical: 40 pts.",
      "Sitemap compliance: 30 pts.",
      "Caps for indexable URLs: missing self-canonical caps at 60, missing sitemap compliance caps at 70, nofollow caps at 85.",
      "Noindex URLs are expected out of the sitemap and are scored on whether exclusion is implemented cleanly.",
    ],
    caveat: "This upload does not include direct indexed/not-indexed status, robots.txt allow/block, HTTP status, or canonical target URL.",
  },
  {
    name: "Links",
    contribution: "14.3%",
    decision: "Uses page-type-specific discovery thresholds because a PDP, Browse page, Brand page, and Deal page should not have the same link expectations.",
    rules: [
      "Unique inlinks: up to 45 pts.",
      "Crawl depth: up to 25 pts.",
      "Unique internal outlinks: up to 20 pts.",
      "Breadcrumb/navigation signal: 10 pts.",
      "Caps: zero unique inlinks caps Links at 60; zero unique outlinks caps Links at 75.",
    ],
    caveat: "Current outlink counts are quantity-only and should eventually split contextual links, product-grid links, navigation links, broken links, redirects, and orphan state.",
  },
  {
    name: "Speed",
    contribution: "14.3%",
    decision: "Uses measured Core Web Vitals only. URLs with no speed measurement are excluded from the Speed score instead of counted as 0.",
    rules: [
      "Each available metric scores Good as 100 pts, Needs Improvement as 70 pts, and Poor as 0 pts.",
      "LCP: good under 2.5s, needs improvement 2.5s to 4.0s, poor over 4.0s.",
      "INP: good under 200ms, needs improvement 200ms to 500ms, poor over 500ms.",
      "CLS: good under 0.1, needs improvement 0.1 to 0.25, poor over 0.25.",
      "The URL Speed score is the average of available LCP, CLS, and INP metrics.",
    ],
    caveat: "Current Botify data has LCP and CLS coverage, but INP is not present until a future CrUX, PageSpeed, or Botify upload includes it.",
  },
  {
    name: "Ranking Potential",
    contribution: "0%",
    decision: "Shows prioritization upside, not health. Higher means more opportunity, not a better page.",
    rules: [
      "Top-10 gap: 30% of the opportunity score.",
      "Quick-win keywords in positions 11-20: 25%.",
      "Striking-distance keywords in positions 21-50: 15%.",
      "Average Botify audience-keyword potential: 15%.",
      "Total Botify audience-keyword potential: 15%.",
    ],
    caveat: "Ranking Potential is intentionally excluded from Overall Tech and should guide what to fix first.",
  },
];

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && quoted && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(value);
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      value = "";
    } else {
      value += char;
    }
  }

  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }

  const [headers, ...records] = rows;
  return records.map((record) =>
    Object.fromEntries(headers.map((header, index) => [header, record[index] ?? ""])),
  );
}

async function fetchCsvPreview(url, maxRows = detailPreviewLimit) {
  const response = await fetch(cacheBust(url));
  if (!response.ok) throw new Error(`Could not load ${url}`);
  if (!response.body) {
    return parseCsv(await response.text()).slice(0, maxRows);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let text = "";
  let newlineCount = 0;

  while (newlineCount <= maxRows) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = decoder.decode(value, { stream: true });
    text += chunk;
    newlineCount += (chunk.match(/\n/g) || []).length;
  }

  await reader.cancel().catch(() => {});
  text += decoder.decode();
  const lastNewline = text.lastIndexOf("\n");
  const previewText = lastNewline > -1 ? text.slice(0, lastNewline + 1) : text;
  return parseCsv(previewText).slice(0, maxRows);
}

function numberValue(row, key) {
  const value = row[key];
  if (value === "" || value == null) return null;
  const parsed = Number(String(value).replaceAll(",", ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function formatScore(value) {
  if (value == null) return "--";
  return Math.round(value).toString();
}

function formatDecimal(value) {
  if (value == null) return "--";
  return Number(value).toLocaleString(undefined, {
    maximumFractionDigits: 1,
  });
}

function formatInteger(value) {
  if (value == null) return "--";
  return Math.round(value).toLocaleString();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function average(rows, key) {
  const values = rows
    .map((row) => numberValue(row, key))
    .filter((value) => value != null);
  return values.length
    ? values.reduce((sum, value) => sum + value, 0) / values.length
    : null;
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (numberValue(row, key) ?? 0), 0);
}

function bandForScore(value) {
  if (value == null) return "empty";
  const displayValue = Math.round(value);
  if (displayValue >= 85) return "green";
  if (displayValue >= 70) return "amber";
  return "red";
}

function bandLabel(value) {
  const band = bandForScore(value);
  if (band === "green") return "Strong";
  if (band === "amber") return "Needs attention";
  if (band === "red") return "At risk";
  return "No score";
}

function filteredRows() {
  return state.rows.filter((row) => {
    if (state.pageType !== "all" && row.page_type !== state.pageType) return false;
    const overall = numberValue(row, "overall_proxy_excl_ranking_keywords");
    if (state.band !== "all" && bandForScore(overall) !== state.band) return false;
    return true;
  });
}

function showStatus(message) {
  els.statusBanner.textContent = message;
  els.statusBanner.classList.toggle("hidden", !message);
}

function renderSummary(rows) {
  const overall = average(rows, "overall_proxy_excl_ranking_keywords");
  const ranking = average(rows, "ranking_keywords");
  const top10 = average(rows, "top10_share_pct");
  const scorable = sum(rows, "scorable_urls");

  els.siteOverall.textContent = formatScore(overall);
  els.siteOverallBand.textContent = bandLabel(overall);
  els.scorableUrls.textContent = formatInteger(scorable);
  els.rankingKeywordsAvg.textContent = formatScore(ranking);
  els.top10ShareAvg.textContent = `${formatDecimal(top10)}% top 10 share`;
}

function scoreCell(row, key) {
  const value = numberValue(row, key);
  const span = document.createElement("span");
  span.className = `score-cell ${key === "ranking_potential" ? "opportunity" : bandForScore(value)}`;
  span.textContent = formatScore(value);
  return span;
}

function renderHeatmap(rows) {
  els.scoreTableBody.replaceChildren();
  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.dataset.pageType = row.page_type;
    const pageType = document.createElement("td");
    pageType.textContent = row.page_type;
    tr.append(pageType);

    for (const [key] of scoreColumns) {
      const td = document.createElement("td");
      td.append(scoreCell(row, key));
      tr.append(td);
    }

    tr.addEventListener("click", () => {
      selectPageType(row);
    });
    els.scoreTableBody.append(tr);
  }
}

function selectedPageType() {
  return state.selected?.page_type || filteredRows()[0]?.page_type || state.rows[0]?.page_type || "";
}

function renderOpportunityPageTypes(rows) {
  const activePageType = selectedPageType();
  els.opportunityPageTypes.replaceChildren(
    ...rows.map((row) => {
      const overall = numberValue(row, "overall_proxy_excl_ranking_keywords");
      const button = document.createElement("button");
      button.className = `page-type-button ${bandForScore(overall)} ${row.page_type === activePageType ? "active" : ""}`;
      button.type = "button";
      button.setAttribute("aria-pressed", row.page_type === activePageType ? "true" : "false");
      button.innerHTML = `
        <span>${escapeHtml(row.page_type)}</span>
        <strong>${formatScore(overall)}</strong>
      `;
      button.addEventListener("click", () => selectPageType(row));
      return button;
    }),
  );
}

function metricStatus(key, value) {
  if (value == null) return "No data";
  if (key === "ranking_potential") return "Upside";
  return bandLabel(value);
}

function metricTone(key, value) {
  if (key === "ranking_potential") return "opportunity";
  return bandForScore(value);
}

function renderMetricBars(row) {
  const target = row || state.selected || filteredRows()[0] || state.rows[0];
  els.rankingBars.replaceChildren();
  if (!target) return;

  const heading = document.createElement("div");
  heading.className = "metric-bar-heading";
  heading.innerHTML = `
    <span>Metric</span>
    <span>Health</span>
    <span>Score</span>
    <span>Status</span>
  `;
  els.rankingBars.append(heading);

  for (const [key, label] of healthBarColumns) {
    const value = numberValue(target, key);
    const width = value == null ? 0 : Math.max(0, Math.min(100, value));
    const tone = metricTone(key, value);
    const item = document.createElement("div");
    item.className = `bar-row metric-health-row ${tone}`;
    item.innerHTML = `
      <span class="bar-label">${escapeHtml(label)}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${width}%"></span></span>
      <span class="bar-score">${formatScore(value)}</span>
      <span class="bar-status">${escapeHtml(metricStatus(key, value))}</span>
    `;
    els.rankingBars.append(item);
  }
}

function renderDetail(row) {
  const target = row || filteredRows()[0] || state.rows[0];
  if (!target) return;
  state.selected = target;
  els.detailTitle.textContent = target.page_type;
  const metricGroups = [
    {
      title: "Health Snapshot",
      items: [
        ["Overall Technical", formatScore(numberValue(target, "overall_proxy_excl_ranking_keywords"))],
        ["Ranking Potential", formatScore(numberValue(target, "ranking_potential"))],
        ["Scorable URLs", formatInteger(numberValue(target, "scorable_urls"))],
        ["Bad Crawls Excluded", formatInteger(numberValue(target, "excluded_bad_crawl_urls"))],
      ],
    },
    {
      title: "Keyword Data",
      items: [
        ["Ranking Keywords", formatScore(numberValue(target, "ranking_keywords"))],
        ["Distinct Keywords", formatInteger(numberValue(target, "distinct_keywords"))],
        ["Keyword Rows", formatInteger(numberValue(target, "keyword_rows"))],
        ["Keyword URLs", formatInteger(numberValue(target, "keyword_urls"))],
        ["Matched Botify URLs", formatInteger(numberValue(target, "keyword_urls_matched_botify"))],
        ["Inventory Coverage", `${formatDecimal(numberValue(target, "keyword_inventory_coverage_pct"))}%`],
        ["Avg Current Rank", formatDecimal(numberValue(target, "avg_current_rank"))],
        ["Top 3 Share", `${formatDecimal(numberValue(target, "top3_share_pct"))}%`],
        ["Top 10 Share", `${formatDecimal(numberValue(target, "top10_share_pct"))}%`],
        ["Top 20 Share", `${formatDecimal(numberValue(target, "top20_share_pct"))}%`],
        ["Lost / Unranked", `${formatDecimal(numberValue(target, "lost_or_unranked_share_pct"))}%`],
        ["Quick Win 11-20", `${formatDecimal(numberValue(target, "quick_win_11_20_share_pct"))}%`],
        ["Striking Distance 21-50", `${formatDecimal(numberValue(target, "striking_distance_21_50_share_pct"))}%`],
        ["Audience Potential", formatInteger(numberValue(target, "total_botify_audience_potential"))],
      ],
    },
    {
      title: "Measurement Coverage",
      items: [
        ["Speed Coverage", `${formatDecimal(numberValue(target, "speed_coverage_pct"))}%`],
        ["Speed Scored URLs", formatInteger(numberValue(target, "speed_scored_urls"))],
        ["Main Keyword Coverage", `${formatDecimal(numberValue(target, "main_keyword_target_coverage_pct"))}%`],
      ],
    },
  ];
  const nodes = [];
  for (const group of metricGroups) {
    const heading = document.createElement("h4");
    heading.className = "detail-section-title";
    heading.textContent = group.title;
    nodes.push(heading);
    for (const [label, value] of group.items) {
      const item = document.createElement("div");
      item.className = "detail-item";
      item.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
      nodes.push(item);
    }
  }
  els.detailMetrics.replaceChildren(...nodes);
}

function issueTableColumns(issue) {
  const base = [["normalized_url", "URL"]];
  const category = issue?.category || "";
  if (category === "Meta") {
    return [
      ...base,
      ["title", "Title"],
      ["title_length", "Title Len"],
      ["meta_description", "Meta Description"],
      ["meta_description_length", "Meta Len"],
      ["main_keyword", "Main Keyword"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "Schema") {
    return [
      ...base,
      ["breadcrumb_exists", "Breadcrumb"],
      ["product_exists", "Product"],
      ["offer_exists", "Offer"],
      ["faq_exists", "FAQ"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "URL") {
    return [
      ...base,
      ["url_slug", "Slug"],
      ["main_keyword", "Main Keyword"],
      ["title", "Title"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "H1") {
    return [
      ...base,
      ["h1", "H1"],
      ["h1_count", "H1 Count"],
      ["title", "Title"],
      ["main_keyword", "Main Keyword"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "Index") {
    return [
      ...base,
      ["noindex", "Noindex"],
      ["nofollow", "Nofollow"],
      ["canonical_points_to_self", "Self Canonical"],
      ["in_sitemaps", "In Sitemap"],
    ];
  }
  if (category === "Links") {
    return [
      ...base,
      ["unique_inlinks", "Inlinks"],
      ["unique_outlinks", "Outlinks"],
      ["depth", "Depth"],
      ["breadcrumb_exists", "Breadcrumb"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "Speed") {
    return [
      ...base,
      ["lcp_status", "LCP"],
      ["lcp_value", "LCP Value"],
      ["cls_status", "CLS"],
      ["cls_value", "CLS Value"],
      ["inp_status", "INP"],
      ["inp_value", "INP Value"],
      ["evidence_value", "Evidence"],
    ];
  }
  if (category === "Ranking Keywords" || category === "Ranking Potential") {
    return [
      ...base,
      ["matched_botify", "In Botify"],
      ["main_keyword", "Main Keyword"],
      ["distinct_keywords", "Keywords"],
      ["avg_current_rank", "Avg Rank"],
      ["top10_keywords", "Top 10"],
      ["quick_win_11_20_keywords", "Quick Win"],
      ["striking_distance_21_50_keywords", "Distance"],
      ["lost_or_unranked_keywords", "Lost"],
      ["botify_audience_keywords_to_90pct", "Audience Potential"],
    ];
  }
  return [...base, ["title", "Title"], ["meta_description", "Meta Description"], ["evidence_value", "Evidence"]];
}

function filteredUrlDetails() {
  const query = state.urlSearch.trim().toLowerCase();
  if (!query) return state.urlDetails;
  return state.urlDetails.filter((row) =>
    [
      row.normalized_url,
      row.title,
      row.meta_description,
      row.main_keyword,
      row.h1,
      row.evidence_value,
    ]
      .join(" ")
      .toLowerCase()
      .includes(query),
  );
}

function renderUrlDetailTable() {
  const issue = state.activeIssue;
  const rows = filteredUrlDetails();
  const columns = issueTableColumns(issue);
  const thead = els.urlDetailTable.querySelector("thead");
  const tbody = els.urlDetailTable.querySelector("tbody");

  thead.replaceChildren();
  tbody.replaceChildren();

  if (!issue || !rows.length) {
    els.urlEmptyState.classList.remove("hidden");
    els.urlEmptyState.textContent = issue
      ? "No loaded rows match the current search."
      : "Choose a recommended fix to see URL-level evidence.";
    return;
  }

  els.urlEmptyState.classList.add("hidden");
  const headerRow = document.createElement("tr");
  for (const [, label] of columns) {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.append(th);
  }
  thead.append(headerRow);

  for (const row of rows.slice(0, detailPreviewLimit)) {
    const tr = document.createElement("tr");
    for (const [key] of columns) {
      const td = document.createElement("td");
      if (key === "normalized_url" && row[key]) {
        const link = document.createElement("a");
        link.href = row[key];
        link.target = "_blank";
        link.rel = "noreferrer";
        link.textContent = row[key];
        td.append(link);
      } else {
        td.textContent = row[key] || "";
      }
      tr.append(td);
    }
    tbody.append(tr);
  }
}

function renderUrlDetailHeader(message = "") {
  const issue = state.activeIssue;
  if (!issue) {
    els.urlListTitle.textContent = "Select A Recommended Fix";
    els.urlListMeta.textContent =
      "Click any callout on the Opportunity tab to inspect the affected URLs and supporting data.";
    els.urlListStatus.textContent = "No issue selected";
    els.urlDownloadLink.classList.add("hidden");
    return;
  }

  const affected = formatInteger(numberValue(issue, "affected_urls"));
  const eligible = formatInteger(numberValue(issue, "eligible_urls"));
  els.urlListTitle.textContent = `${issue.page_type}: ${issue.issue}`;
  els.urlListMeta.textContent = `${issue.category} issue affecting ${affected} of ${eligible} eligible URLs. ${issue.recommendation}`;
  els.urlListStatus.textContent =
    message ||
    `Showing up to ${formatInteger(Math.min(state.urlDetails.length, detailPreviewLimit))} loaded rows. CSV download uses the available detail file.`;
  if (issue.detail_file) {
    els.urlDownloadLink.href = dataUrl(issue.detail_file);
    els.urlDownloadLink.classList.remove("hidden");
  } else {
    els.urlDownloadLink.classList.add("hidden");
  }
}

async function openIssueDetail(issue) {
  state.activeIssue = issue;
  state.urlDetails = [];
  state.urlSearch = "";
  els.urlSearchInput.value = "";
  setView("urls");
  renderUrlDetailHeader("Loading URL-level evidence...");
  renderUrlDetailTable();

  if (!issue.detail_file) {
    renderUrlDetailHeader("No detail file is available for this uploaded issue summary.");
    return;
  }

  try {
    state.urlDetails = await fetchCsvPreview(issue.detail_file);
    renderUrlDetailHeader();
    renderUrlDetailTable();
  } catch (error) {
    state.urlDetails = [];
    renderUrlDetailHeader(error.message);
    renderUrlDetailTable();
  }
}

function renderIssueDrivers(row) {
  const target = row || state.selected || filteredRows()[0] || state.rows[0];
  if (!target) return;
  els.issueHeading.textContent = `Why ${target.page_type} Is Not At 100`;
  const issues = state.issues
    .filter((issue) => issue.page_type === target.page_type)
    .sort((a, b) => {
      const categoryDelta =
        categoryOrder.indexOf(a.category) - categoryOrder.indexOf(b.category);
      if (categoryDelta) return categoryDelta;
      return (numberValue(b, "affected_urls") ?? 0) - (numberValue(a, "affected_urls") ?? 0);
    });

  if (!issues.length) {
    els.issueGrid.replaceChildren();
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No aggregated issue drivers are available for this page type.";
    els.issueGrid.append(empty);
    return;
  }

  const grouped = new Map();
  for (const issue of issues) {
    if (!grouped.has(issue.category)) grouped.set(issue.category, []);
    grouped.get(issue.category).push(issue);
  }

  els.issueGrid.replaceChildren(
    ...categoryOrder
      .filter((category) => grouped.has(category))
      .map((category) => {
        const categoryIssues = grouped.get(category);
        const impacted = categoryIssues.reduce(
          (max, issue) => Math.max(max, numberValue(issue, "affected_urls") ?? 0),
          0,
        );
        const panel = document.createElement("article");
        panel.className = "issue-category";
        panel.innerHTML = `
          <div class="issue-category-header">
            <h3>${escapeHtml(category)}</h3>
            <span>${formatInteger(impacted)} max impacted</span>
          </div>
          <div class="issue-list">
            ${categoryIssues
              .map((issue) => {
                const affected = numberValue(issue, "affected_urls");
                const pct = numberValue(issue, "affected_pct");
                const severity = String(issue.severity || "low").toLowerCase();
                const detailFile = issue.detail_file || "";
                return `
                  <button class="issue-row issue-action-row" type="button"
                    data-page-type="${escapeHtml(issue.page_type)}"
                    data-category="${escapeHtml(issue.category)}"
                    data-issue="${escapeHtml(issue.issue)}"
                    data-detail-file="${escapeHtml(detailFile)}">
                    <div>
                      <div class="issue-title">${escapeHtml(issue.issue)}</div>
                      <p class="issue-copy">${escapeHtml(issue.evidence)}</p>
                      <p class="issue-copy">${escapeHtml(issue.recommendation)}</p>
                    </div>
                    <div class="issue-count">
                      <strong>${formatInteger(affected)}</strong>
                      <span>${formatDecimal(pct)}%</span>
                    </div>
                    <span class="issue-action-stack">
                      <span class="severity-pill ${escapeHtml(severity)}">${escapeHtml(severity)}</span>
                      <span class="drill-link">${detailFile ? "View URLs" : "No detail"}</span>
                    </span>
                  </button>
                `;
              })
              .join("")}
          </div>
        `;
        return panel;
      }),
  );
}

function renderWeights() {
  els.weightsGrid.replaceChildren(
    ...weightModels.map((model) => {
      const article = document.createElement("article");
      article.className = `weight-card ${model.contribution === "0%" ? "supporting" : ""}`;
      article.innerHTML = `
        <div class="weight-card-header">
          <div>
            <span class="metric-label">Top-level metric</span>
            <h3>${escapeHtml(model.name)}</h3>
          </div>
          <span class="weight-pill">${escapeHtml(model.contribution)} of Overall Tech</span>
        </div>
        <p class="weight-decision">${escapeHtml(model.decision)}</p>
        <ul class="weight-rule-list">
          ${model.rules.map((rule) => `<li>${escapeHtml(rule)}</li>`).join("")}
        </ul>
        <p class="weight-caveat">${escapeHtml(model.caveat)}</p>
      `;
      return article;
    }),
  );
}

function renderInputs(rows) {
  const inputItems = [
    {
      title: "Botify Crawl",
      status: "Connected",
      tone: "green",
      body: `${formatInteger(sum(rows, "botify_urls"))} crawled URLs, ${formatInteger(sum(rows, "scorable_urls"))} scorable URLs.`,
    },
    {
      title: "Conductor Keywords",
      status: "Connected",
      tone: "green",
      body: `${formatInteger(sum(rows, "keyword_rows"))} ranking rows across ${formatInteger(sum(rows, "distinct_keywords"))} distinct keyword-page-type signals.`,
    },
    {
      title: "CrUX / PageSpeed",
      status: "Partial",
      tone: "amber",
      body: `${formatDecimal(average(rows, "speed_coverage_pct"))}% average speed coverage. INP is not present in the current upload.`,
    },
    {
      title: "Botify Audience Potential",
      status: "Connected",
      tone: "green",
      body: `${formatInteger(sum(rows, "total_botify_audience_potential"))} total audience-keyword potential signals.`,
    },
    {
      title: "On-Page Content",
      status: "Pending",
      tone: "red",
      body: "Body-copy depth, uniqueness, freshness, and duplicate-content fields are not included yet.",
    },
    {
      title: "Botify ZIP Upload",
      status: "Planned",
      tone: "blue",
      body: "Current UI reads the generated scorecard CSV. Direct ZIP parsing belongs in the next local backend pass.",
    },
  ];

  els.inputGrid.replaceChildren(
    ...inputItems.map((item) => {
      const article = document.createElement("article");
      article.className = "input-item";
      article.innerHTML = `
        <span class="input-status ${item.tone}">${item.status}</span>
        <h3>${item.title}</h3>
        <p>${item.body}</p>
      `;
      return article;
    }),
  );
}

function populatePageTypes(rows) {
  const options = ["all", ...rows.map((row) => row.page_type)];
  els.pageTypeFilter.replaceChildren(
    ...options.map((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value === "all" ? "All" : value;
      return option;
    }),
  );
}

function renderAll() {
  const rows = filteredRows();
  const selected = state.selected || rows[0] || state.rows[0];
  state.selected = selected || null;
  renderSummary(rows);
  renderHeatmap(rows);
  renderOpportunityPageTypes(state.rows);
  renderMetricBars(selected);
  renderDetail(selected);
  renderIssueDrivers(selected);
  renderWeights();
  renderInputs(state.rows);
}

async function loadCurrentCsv() {
  const response = await fetch(cacheBust(DATA_URL));
  if (!response.ok) throw new Error(`Could not load ${DATA_URL}`);
  const rows = parseCsv(await response.text());
  const issueResponse = await fetch(cacheBust(ISSUE_URL));
  const issues = issueResponse.ok ? parseCsv(await issueResponse.text()) : [];
  state.rows = rows;
  state.issues = issues;
  state.activeIssue = null;
  state.urlDetails = [];
  state.urlSearch = "";
  state.selected = rows[0] || null;
  populatePageTypes(rows);
  renderAll();
  showStatus("Loaded current scorecard CSV.");
  setTimeout(() => showStatus(""), 2400);
}

function setView(view) {
  state.view = view;
  document.querySelectorAll(".segment").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === view);
  });
  document.querySelectorAll(".view").forEach((panel) => {
    panel.classList.toggle("active", panel.id === `${view}View`);
  });
}

function selectPageType(row) {
  state.selected = row;
  setView("opportunity");
  renderAll();
}

document.querySelectorAll(".segment").forEach((button) => {
  button.addEventListener("click", () => setView(button.dataset.view));
});

els.pageTypeFilter.addEventListener("change", (event) => {
  state.pageType = event.target.value;
  state.selected = filteredRows()[0] || state.rows[0] || null;
  renderAll();
});

els.bandFilter.addEventListener("change", (event) => {
  state.band = event.target.value;
  renderAll();
});

els.issueGrid.addEventListener("click", (event) => {
  const target = event.target.closest(".issue-action-row");
  if (!target) return;
  const issue = state.issues.find(
    (candidate) =>
      candidate.page_type === target.dataset.pageType &&
      candidate.category === target.dataset.category &&
      candidate.issue === target.dataset.issue,
  );
  if (issue) openIssueDetail(issue);
});

els.urlSearchInput.addEventListener("input", (event) => {
  state.urlSearch = event.target.value;
  renderUrlDetailHeader();
  renderUrlDetailTable();
});

els.backToOpportunityButton.addEventListener("click", () => {
  setView("opportunity");
});

els.refreshButton.addEventListener("click", () => {
  loadCurrentCsv().catch((error) => showStatus(error.message));
});

els.csvInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  const rows = parseCsv(await file.text());
  state.rows = rows;
  state.issues = [];
  state.activeIssue = null;
  state.urlDetails = [];
  state.selected = rows[0] || null;
  populatePageTypes(rows);
  renderAll();
  showStatus(`Loaded ${file.name}. Issue drilldown requires the generated issue summary CSV.`);
});

loadCurrentCsv().catch((error) => showStatus(error.message));
