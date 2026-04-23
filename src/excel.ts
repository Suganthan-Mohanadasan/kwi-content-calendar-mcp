import ExcelJS from "exceljs";
import type { CalendarItem } from "./types.js";

const TITLE_BG = "FF4472C4";
const TITLE_FG = "FFFFFFFF";
const SUBTITLE_BG = "FFDCE6F1";
const HEADER_BG = "FF1B3A5C";
const HEADER_FG = "FFFFFFFF";
const SECTION_BG = "FFD9E1F2";
const ALT_ROW_BG = "FFF4F7FB";
const TOTAL_BG = "FFFFF2CC";

const ACTION_COLORS: Record<string, string> = {
  CREATE: "FF2E7D32",
  UPDATE: "FF1565C0",
  OPTIMISE: "FFEF6C00",
  REWRITE: "FFC62828"
};

export interface ExportMetadata {
  totalClusters: number;
  domain: string;
  piecesPerWeek: number;
  weeks: number;
}

function setTitleRow(sheet: ExcelJS.Worksheet, row: number, text: string, colCount: number): void {
  sheet.mergeCells(row, 1, row, colCount);
  const cell = sheet.getCell(row, 1);
  cell.value = text;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TITLE_BG } };
  cell.font = { bold: true, size: 14, color: { argb: TITLE_FG } };
  cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(row).height = 30;
}

function setSubtitleRow(sheet: ExcelJS.Worksheet, row: number, text: string, colCount: number): void {
  sheet.mergeCells(row, 1, row, colCount);
  const cell = sheet.getCell(row, 1);
  cell.value = text;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTITLE_BG } };
  cell.font = { italic: true, color: { argb: "FF333333" } };
  cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(row).height = 22;
}

function setSectionRow(sheet: ExcelJS.Worksheet, row: number, text: string, colCount: number): void {
  sheet.mergeCells(row, 1, row, colCount);
  const cell = sheet.getCell(row, 1);
  cell.value = text;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SECTION_BG } };
  cell.font = { bold: true, size: 12, color: { argb: "FF1B3A5C" } };
  cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(row).height = 24;
}

function styleHeaderRow(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_BG } };
    cell.font = { bold: true, color: { argb: HEADER_FG } };
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    cell.border = { bottom: { style: "thin", color: { argb: HEADER_BG } } };
  });
  row.height = 40;
}

function applyDataRowStyle(row: ExcelJS.Row, index: number, actionCol: number, action: string): void {
  if (index % 2 === 1) {
    row.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ALT_ROW_BG } };
    });
  }
  row.eachCell((cell) => {
    cell.alignment = { vertical: "top", wrapText: true };
  });
  const actionCell = row.getCell(actionCol);
  const colour = ACTION_COLORS[action];
  if (colour) {
    actionCell.font = { bold: true, color: { argb: colour } };
  }
}

function whatToLookFor(item: CalendarItem): string {
  const rank = Number.isFinite(item.cluster_rank) ? item.cluster_rank.toFixed(1) : "n/a";
  const kwCount = item.cluster_n_kwd ?? 0;

  if (item.action === "OPTIMISE") {
    return `Near top 3 (rank ${rank}). Identify depth gaps vs competitors, add FAQ schema for PAA, strengthen entity coverage, improve internal linking.`;
  }
  if (item.action === "UPDATE") {
    return `Striking distance (rank ${rank}). Find missing subtopics, expand thin sections, add ${kwCount} cluster keywords naturally. Check competitor word count.`;
  }
  if (item.action === "REWRITE") {
    return `Rank ${rank} with low content score (${item.content_score || "n/a"}). Foundation too weak to polish. Full content rewrite against top 3 competitors.`;
  }
  return item.reason;
}

function buildCalendarSheet(wb: ExcelJS.Workbook, items: CalendarItem[], meta: ExportMetadata): void {
  const sheet = wb.addWorksheet("Content Calendar", {
    views: [{ state: "frozen", ySplit: 4 }]
  });

  const columns = [
    { header: "Week", width: 7 },
    { header: "Score", width: 8 },
    { header: "Action", width: 12 },
    { header: "Target Keyword\n(Recommended for Content)", width: 34 },
    { header: "Content Type", width: 26 },
    { header: "Cluster\nSearch Vol", width: 11 },
    { header: "Cluster\nOpportunity", width: 12 },
    { header: "Cluster\nDifficulty", width: 11 },
    { header: "Cluster\nAvg Rank", width: 11 },
    { header: "Keywords\nin Cluster", width: 11 },
    { header: "Cluster URL\n(Page to Update/Analyse)", width: 46 },
    { header: "Needs\nScraping?", width: 11 },
    { header: "Intent", width: 14 },
    { header: "Journey", width: 10 },
    { header: "Hub Topic", width: 28 },
    { header: "Supporting Keywords (from cluster)", width: 50 },
    { header: "SERP Features", width: 22 },
    { header: "AI Overview", width: 12 },
    { header: "Action Rationale", width: 55 }
  ];
  const colCount = columns.length;

  setTitleRow(sheet, 1, `Content Calendar: ${meta.weeks} Week Plan (${meta.piecesPerWeek} pieces/week) | Cluster-Level Analysis`, colCount);
  setSubtitleRow(sheet, 2, `Generated from ${meta.totalClusters.toLocaleString()} keyword clusters | ${meta.domain} | Cluster-level ranking, opportunity, and intent data`, colCount);
  sheet.getRow(3).height = 8;

  const headerRow = sheet.getRow(4);
  columns.forEach((col, i) => {
    headerRow.getCell(i + 1).value = col.header;
    sheet.getColumn(i + 1).width = col.width;
  });
  styleHeaderRow(headerRow);

  const sortedItems = [...items].sort((a, b) => {
    if ((a.week ?? 0) !== (b.week ?? 0)) return (a.week ?? 0) - (b.week ?? 0);
    return (b.priority_score ?? 0) - (a.priority_score ?? 0);
  });

  const firstDataRow = 5;
  sortedItems.forEach((item, i) => {
    const row = sheet.getRow(firstDataRow + i);
    row.getCell(1).value = item.week;
    row.getCell(2).value = item.priority_score;
    row.getCell(3).value = item.action;
    row.getCell(4).value = item.keyword;
    row.getCell(5).value = item.content_type;
    row.getCell(6).value = item.cluster_search_volume;
    row.getCell(7).value = item.cluster_opportunity;
    row.getCell(8).value = item.cluster_difficulty;
    row.getCell(9).value = item.cluster_rank;
    row.getCell(10).value = item.cluster_n_kwd;
    row.getCell(11).value = item.cluster_url;
    row.getCell(12).value = item.needs_scraping ? "Yes" : "No";
    row.getCell(13).value = item.intent;
    row.getCell(14).value = item.journey;
    row.getCell(15).value = item.hub;
    row.getCell(16).value = item.supporting_kws;
    row.getCell(17).value = item.serp_features;
    row.getCell(18).value = item.ai_overview;
    row.getCell(19).value = item.reason;
    applyDataRowStyle(row, i, 3, item.action);
  });

  const totalsRow = firstDataRow + sortedItems.length + 1;
  sheet.getCell(totalsRow, 4).value = "TOTALS";
  sheet.getCell(totalsRow, 6).value = { formula: `SUM(F${firstDataRow}:F${firstDataRow + sortedItems.length - 1})` };
  sheet.getCell(totalsRow, 7).value = { formula: `SUM(G${firstDataRow}:G${firstDataRow + sortedItems.length - 1})` };
  sheet.getCell(totalsRow, 10).value = { formula: `SUM(J${firstDataRow}:J${firstDataRow + sortedItems.length - 1})` };
  for (let c = 1; c <= colCount; c++) {
    const cell = sheet.getCell(totalsRow, c);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TOTAL_BG } };
    cell.font = { bold: true };
    cell.border = { top: { style: "thin", color: { argb: HEADER_BG } } };
  }

  const lastDataRow = firstDataRow + sortedItems.length - 1;
  sheet.addConditionalFormatting({
    ref: `B${firstDataRow}:B${lastDataRow}`,
    rules: [{
      type: "colorScale",
      priority: 1,
      cfvo: [
        { type: "num", value: 0 },
        { type: "num", value: 50 },
        { type: "num", value: 100 }
      ],
      color: [
        { argb: "FFFFCDD2" },
        { argb: "FFFFF59D" },
        { argb: "FFC8E6C9" }
      ]
    }]
  });
}

function buildGapAnalysisSheet(wb: ExcelJS.Workbook): void {
  const sheet = wb.addWorksheet("Gap Analysis Logic");
  const colCount = 4;
  sheet.getColumn(1).width = 34;
  sheet.getColumn(2).width = 60;
  sheet.getColumn(3).width = 36;
  sheet.getColumn(4).width = 60;

  setTitleRow(sheet, 1, "Page Scraping and Gap Analysis: Logic for UPDATE / OPTIMISE / REWRITE Actions", colCount);
  sheet.getRow(2).height = 8;

  const steps: { title: string; rows: [string, string, string, string][] }[] = [
    {
      title: "STEP 1: Scrape Existing Page",
      rows: [
        ["Page title and meta description", "Extract the <title> tag and meta description to check alignment with target keyword and intent.", "Current title, meta desc, character counts", "Identifies if title/meta need rewriting to match cluster target keyword"],
        ["Heading structure (H1 to H6)", "Extract all headings to map the content structure and identify what subtopics the page covers.", "List of all headings with hierarchy", "AI writer knows what sections exist and what to add/restructure"],
        ["Word count", "Count total words on the page to compare against SERP competitors.", "Total word count", "Sets target word count: if competitors average 2,500 words and page has 800, the gap is clear"],
        ["Key entities and topics covered", "NLP extraction of main entities, concepts, and topics discussed on the page.", "List of entities/topics with frequency", "Gap detection: compare against cluster keywords and competitor entities"],
        ["Internal and external links", "Extract all outbound links to check internal linking health and external references.", "Link count, internal vs external, anchor text", "Identifies missing internal links to related hub/spoke content"],
        ["Images and media", "Count and catalogue images, videos, infographics on the page.", "Media count, alt text presence", "Flags missing visual content that competitors use"],
        ["Structured data / Schema", "Check for FAQ schema, HowTo schema, Article schema etc.", "Schema types present or missing", "If SERP shows PAA or AI Overview, recommend adding FAQ schema"]
      ]
    },
    {
      title: "STEP 2: Scrape Top 3 SERP Competitors",
      rows: [
        ["Competitor heading structure", "Extract H2/H3 headings from top 3 competitors to find common subtopics.", "Union of all competitor headings", "Identifies subtopics that ALL competitors cover but the client page misses"],
        ["Competitor word count", "Average word count across top 3 competitors.", "Average competitor word count", "Sets the target word count for the updated content"],
        ["Competitor entities", "NLP extraction of key entities from competitor content.", "List of entities competitors cover", "Direct gap list: entities competitors mention that the client page does not"],
        ["Competitor FAQ/questions answered", "Extract questions answered (from headings, FAQ sections, PAA responses).", "List of questions", "Provides specific questions the updated content must answer"]
      ]
    },
    {
      title: "STEP 3: Gap Analysis Calculation",
      rows: [
        ["Missing subtopics", "Headings present in 2+ competitors but absent from client page.", "List of missing H2/H3 topics", "Required sections to add to the content"],
        ["Missing entities", "Entities mentioned by 2+ competitors but not on client page.", "List of missing entities with context", "Keywords and concepts the writer must naturally include"],
        ["Content depth gap", "Difference between client word count and competitor average.", "Word count delta and percentage", "Target expansion amount (e.g. add 1,200 words)"],
        ["Questions not answered", "Questions in competitor FAQs or PAA that the client does not address.", "List of unanswered questions", "Specific Q&A pairs to add, especially for AI Overview eligibility"],
        ["Structural issues", "Missing schema, poor heading hierarchy, no images where competitors have them.", "List of structural improvements", "Technical optimisation requirements alongside content changes"],
        ["Internal linking opportunities", "Other hub/spoke pages in the calendar that should link to/from this page.", "Suggested internal link targets with anchor text", "Specific pages to link to, building topical authority for the hub"]
      ]
    },
    {
      title: "STEP 4: Generate Enhanced Brief",
      rows: [
        ["Target keyword and cluster context", "Lead keyword + all supporting keywords + intent + journey stage", "Brief header", "Sets the writing direction and keyword coverage requirements"],
        ["What to keep", "Sections/entities on the existing page that are already strong (match competitor coverage).", "Preservation list", "Tells the writer what NOT to change"],
        ["What to add", "Missing subtopics, entities, questions, and structural elements from gap analysis.", "Addition requirements with priority order", "The core of the brief: specific content to create"],
        ["What to restructure", "Heading hierarchy fixes, content reordering, CTA placement based on journey stage.", "Restructuring instructions", "Guides the writer on information architecture"],
        ["Target metrics", "Word count target, keyword density guidance, schema to implement.", "Quantitative targets", "Measurable goals the writer must hit"]
      ]
    }
  ];

  let r = 3;
  for (const step of steps) {
    setSectionRow(sheet, r, step.title, colCount);
    r++;
    const sub = ["What we extract", "How and why", "Data extracted", "How it feeds into writing agent"];
    const subHeader = sheet.getRow(r);
    sub.forEach((h, i) => { subHeader.getCell(i + 1).value = h; });
    styleHeaderRow(subHeader);
    r++;
    step.rows.forEach(([a, b, c, d], i) => {
      const row = sheet.getRow(r);
      row.getCell(1).value = a;
      row.getCell(2).value = b;
      row.getCell(3).value = c;
      row.getCell(4).value = d;
      row.height = 40;
      row.eachCell((cell) => {
        cell.alignment = { vertical: "top", wrapText: true };
      });
      if (i % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ALT_ROW_BG } };
        });
      }
      row.getCell(1).font = { bold: true };
      r++;
    });
    r++;
  }
}

function buildScrapeSheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Pages to Scrape", {
    views: [{ state: "frozen", ySplit: 3 }]
  });

  const columns = [
    { header: "Week", width: 7 },
    { header: "Action", width: 12 },
    { header: "Target Keyword", width: 34 },
    { header: "URL to Scrape", width: 60 },
    { header: "Cluster Avg Rank", width: 14 },
    { header: "Content Score", width: 13 },
    { header: "Keywords in Cluster", width: 16 },
    { header: "What to Look For", width: 70 }
  ];
  const colCount = columns.length;

  setTitleRow(sheet, 1, "Pages Requiring Scraping and Gap Analysis (from Content Calendar)", colCount);
  sheet.getRow(2).height = 8;

  const headerRow = sheet.getRow(3);
  columns.forEach((col, i) => {
    headerRow.getCell(i + 1).value = col.header;
    sheet.getColumn(i + 1).width = col.width;
  });
  styleHeaderRow(headerRow);

  const scrapeItems = items
    .filter((i) => i.needs_scraping)
    .sort((a, b) => {
      if ((a.week ?? 0) !== (b.week ?? 0)) return (a.week ?? 0) - (b.week ?? 0);
      return (b.priority_score ?? 0) - (a.priority_score ?? 0);
    });

  scrapeItems.forEach((item, i) => {
    const row = sheet.getRow(4 + i);
    row.getCell(1).value = item.week;
    row.getCell(2).value = item.action;
    row.getCell(3).value = item.keyword;
    row.getCell(4).value = item.cluster_url;
    row.getCell(5).value = item.cluster_rank;
    row.getCell(6).value = item.content_score;
    row.getCell(7).value = item.cluster_n_kwd;
    row.getCell(8).value = whatToLookFor(item);
    applyDataRowStyle(row, i, 2, item.action);
  });
}

function buildBacklogSheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Priority Backlog", {
    views: [{ state: "frozen", ySplit: 3 }]
  });

  const columns = [
    { header: "Score", width: 8 },
    { header: "Action", width: 12 },
    { header: "Target Keyword", width: 40 },
    { header: "Content Type", width: 28 },
    { header: "Cluster SV", width: 11 },
    { header: "Cluster Opp", width: 12 },
    { header: "Cluster Diff", width: 12 },
    { header: "Cluster Rank", width: 13 },
    { header: "KWs", width: 7 },
    { header: "Intent", width: 14 },
    { header: "Journey", width: 10 },
    { header: "Hub", width: 30 }
  ];
  const colCount = columns.length;

  setTitleRow(sheet, 1, "Priority Backlog: Next 30 Items (Week 5+)", colCount);
  sheet.getRow(2).height = 8;

  const headerRow = sheet.getRow(3);
  columns.forEach((col, i) => {
    headerRow.getCell(i + 1).value = col.header;
    sheet.getColumn(i + 1).width = col.width;
  });
  styleHeaderRow(headerRow);

  const sorted = [...items]
    .sort((a, b) => (b.priority_score ?? 0) - (a.priority_score ?? 0))
    .slice(0, 30);

  sorted.forEach((item, i) => {
    const row = sheet.getRow(4 + i);
    row.getCell(1).value = item.priority_score;
    row.getCell(2).value = item.action;
    row.getCell(3).value = item.keyword;
    row.getCell(4).value = item.content_type;
    row.getCell(5).value = item.cluster_search_volume;
    row.getCell(6).value = item.cluster_opportunity;
    row.getCell(7).value = item.cluster_difficulty;
    row.getCell(8).value = item.cluster_rank;
    row.getCell(9).value = item.cluster_n_kwd;
    row.getCell(10).value = item.intent;
    row.getCell(11).value = item.journey;
    row.getCell(12).value = item.hub;
    applyDataRowStyle(row, i, 2, item.action);
  });

  if (sorted.length > 0) {
    sheet.addConditionalFormatting({
      ref: `A4:A${3 + sorted.length}`,
      rules: [{
        type: "colorScale",
        priority: 1,
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: 50 },
          { type: "num", value: 100 }
        ],
        color: [
          { argb: "FFFFCDD2" },
          { argb: "FFFFF59D" },
          { argb: "FFC8E6C9" }
        ]
      }]
    });
  }
}

function buildSummarySheet(wb: ExcelJS.Workbook, calendar: CalendarItem[], meta: ExportMetadata): void {
  const sheet = wb.addWorksheet("Summary");
  const colCount = 2;
  sheet.getColumn(1).width = 44;
  sheet.getColumn(2).width = 70;

  setTitleRow(sheet, 1, "Content Calendar Summary", colCount);
  sheet.getRow(2).height = 8;

  setSectionRow(sheet, 3, "Cluster-Level Analysis Approach", colCount);

  const principlesHeader = sheet.getRow(4);
  principlesHeader.getCell(1).value = "Principle";
  principlesHeader.getCell(2).value = "Implementation";
  styleHeaderRow(principlesHeader);

  const scrapeCount = calendar.filter((i) => i.needs_scraping).length;
  const principles: [string, string][] = [
    ["All decisions made at CLUSTER level, not keyword level", "Using cluster_url, cluster_rank, cluster_search_volume, cluster_opportunity, cluster_difficulty, cluster_content_score, cluster_n_kwd"],
    ["Lead keyword = recommended keyword for content", "The is_lead_keyword=True row provides the target keyword. All other cluster keywords become supporting keywords."],
    ["Cluster URL = the page to update or scrape", "cluster_url represents the primary ranking page across the cluster. This may differ from the lead keyword individual URL."],
    ["Intent mismatch checked across ALL ranking keywords", "Not just the lead keyword. If >60% of ranking keywords in the cluster have is_intent_match=False, the whole cluster flags as mismatch."],
    ["Cluster rank is the AVERAGE across ranking keywords", "A cluster with avg rank 17.1 means some keywords are page one, others are page two. The average drives the action classification."],
    ["Page scraping triggered for UPDATE/OPTIMISE/REWRITE", `${scrapeCount} of ${calendar.length} calendar items need gap analysis. The scraping extracts content structure, entities, and competitor comparison.`]
  ];

  principles.forEach(([p, impl], i) => {
    const row = sheet.getRow(5 + i);
    row.getCell(1).value = p;
    row.getCell(2).value = impl;
    row.getCell(1).font = { bold: true };
    row.eachCell((cell) => {
      cell.alignment = { vertical: "top", wrapText: true };
    });
    if (i % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ALT_ROW_BG } };
      });
    }
    row.height = 38;
  });

  let r = 5 + principles.length + 1;
  setSectionRow(sheet, r, "Calendar Metrics", colCount);
  r++;

  const metricsHeader = sheet.getRow(r);
  metricsHeader.getCell(1).value = "Metric";
  metricsHeader.getCell(2).value = "Value";
  styleHeaderRow(metricsHeader);
  r++;

  const totalSv = calendar.reduce((s, i) => s + (i.cluster_search_volume || 0), 0);
  const totalOpp = calendar.reduce((s, i) => s + (i.cluster_opportunity || 0), 0);
  const totalKws = calendar.reduce((s, i) => s + (i.cluster_n_kwd || 0), 0);
  const avgScore = calendar.length > 0
    ? Math.round(calendar.reduce((s, i) => s + (i.priority_score || 0), 0) / calendar.length)
    : 0;

  const metrics: [string, number | string][] = [
    ["Total clusters analysed", meta.totalClusters],
    [`Scheduled items (${meta.weeks} weeks)`, calendar.length],
    ["Items needing page scraping", scrapeCount],
    ["Total cluster search volume (scheduled)", totalSv],
    ["Total cluster opportunity (scheduled)", Math.round(totalOpp)],
    ["Total keywords covered across clusters", totalKws],
    ["Average priority score", avgScore]
  ];

  metrics.forEach(([m, v], i) => {
    const row = sheet.getRow(r + i);
    row.getCell(1).value = m;
    row.getCell(2).value = v;
    row.getCell(1).font = { bold: true };
    row.getCell(2).alignment = { horizontal: "right" };
    if (typeof v === "number" && v >= 1000) {
      row.getCell(2).numFmt = "#,##0";
    }
    if (i % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ALT_ROW_BG } };
      });
    }
  });
}

export async function buildExcel(
  calendar: CalendarItem[],
  backlog: CalendarItem[],
  meta: ExportMetadata
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Keyword Insights";
  wb.created = new Date();
  buildCalendarSheet(wb, calendar, meta);
  buildGapAnalysisSheet(wb);
  buildScrapeSheet(wb, calendar);
  buildBacklogSheet(wb, backlog);
  buildSummarySheet(wb, calendar, meta);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

export function buildCsv(calendar: CalendarItem[]): string {
  const headers = [
    "week", "score", "action", "target_keyword", "content_type",
    "cluster_search_volume", "cluster_opportunity", "cluster_difficulty", "cluster_avg_rank",
    "keywords_in_cluster", "cluster_url", "needs_scraping", "intent", "journey", "hub",
    "supporting_keywords", "serp_features", "ai_overview", "action_rationale"
  ];
  const rows = calendar.map((item) => [
    item.week, item.priority_score, item.action, item.keyword, item.content_type,
    item.cluster_search_volume, item.cluster_opportunity, item.cluster_difficulty, item.cluster_rank,
    item.cluster_n_kwd, item.cluster_url, item.needs_scraping ? "Yes" : "No",
    item.intent, item.journey, item.hub, item.supporting_kws,
    item.serp_features, item.ai_overview, item.reason
  ]);

  const escape = (v: unknown): string => {
    const s = String(v ?? "");
    return s.includes(",") || s.includes('"') || s.includes("\n")
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  };

  return [headers.join(","), ...rows.map((r) => r.map(escape).join(","))].join("\n");
}
