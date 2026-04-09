import ExcelJS from "exceljs";
import type { CalendarItem } from "./types.js";

const NAVY = "FF1B3A5C";
const WHITE = "FFFFFFFF";
const ALT_ROW = "FFF4F7FB";

const ACTION_COLORS: Record<string, string> = {
  CREATE: "FF2E7D32",
  UPDATE: "FF1565C0",
  OPTIMISE: "FFEF6C00",
  REWRITE: "FFC62828"
};

const COLUMNS: { header: string; key: string; width: number }[] = [
  { header: "Week", key: "week", width: 7 },
  { header: "Action", key: "action", width: 12 },
  { header: "Keyword", key: "keyword", width: 36 },
  { header: "Intent", key: "intent", width: 14 },
  { header: "Journey", key: "journey", width: 10 },
  { header: "Priority", key: "priority_score", width: 10 },
  { header: "Cluster Rank", key: "cluster_rank", width: 12 },
  { header: "Cluster URL", key: "cluster_url", width: 50 },
  { header: "Search Volume", key: "cluster_search_volume", width: 14 },
  { header: "Opportunity", key: "cluster_opportunity", width: 12 },
  { header: "Difficulty", key: "cluster_difficulty", width: 11 },
  { header: "Content Type", key: "content_type", width: 28 },
  { header: "Needs Scraping", key: "needs_scraping", width: 14 },
  { header: "Reason", key: "reason", width: 60 },
  { header: "Supporting Keywords", key: "supporting_kws", width: 60 }
];

function applyHeader(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    cell.font = { color: { argb: WHITE }, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.border = { bottom: { style: "thin", color: { argb: NAVY } } };
  });
}

function applyDataRow(row: ExcelJS.Row, index: number, item: CalendarItem): void {
  if (index % 2 === 0) {
    row.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ALT_ROW } };
    });
  }
  const actionCell = row.getCell(2);
  const colour = ACTION_COLORS[item.action];
  if (colour) {
    actionCell.font = { bold: true, color: { argb: colour } };
  }
}

function buildCalendarSheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Content Calendar", {
    views: [{ state: "frozen", ySplit: 1 }]
  });
  sheet.columns = COLUMNS.map((c) => ({ header: c.header, key: c.key, width: c.width }));
  applyHeader(sheet.getRow(1));

  items.forEach((item, i) => {
    const row = sheet.addRow({
      week: item.week,
      action: item.action,
      keyword: item.keyword,
      intent: item.intent,
      journey: item.journey,
      priority_score: item.priority_score,
      cluster_rank: item.cluster_rank,
      cluster_url: item.cluster_url,
      cluster_search_volume: item.cluster_search_volume,
      cluster_opportunity: item.cluster_opportunity,
      cluster_difficulty: item.cluster_difficulty,
      content_type: item.content_type,
      needs_scraping: item.needs_scraping ? "Yes" : "No",
      reason: item.reason,
      supporting_kws: item.supporting_kws
    });
    applyDataRow(row, i, item);
  });

  sheet.addConditionalFormatting({
    ref: `F2:F${items.length + 1}`,
    rules: [
      {
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
      }
    ]
  });
}

function buildScrapeSheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Pages to Scrape", {
    views: [{ state: "frozen", ySplit: 1 }]
  });
  sheet.columns = [
    { header: "Action", key: "action", width: 12 },
    { header: "Keyword", key: "keyword", width: 36 },
    { header: "Cluster URL", key: "cluster_url", width: 60 },
    { header: "Cluster Rank", key: "cluster_rank", width: 12 },
    { header: "Reason", key: "reason", width: 60 }
  ];
  applyHeader(sheet.getRow(1));
  items
    .filter((i) => i.needs_scraping)
    .forEach((item, i) => {
      const row = sheet.addRow({
        action: item.action,
        keyword: item.keyword,
        cluster_url: item.cluster_url,
        cluster_rank: item.cluster_rank,
        reason: item.reason
      });
      applyDataRow(row, i, item);
    });
}

function buildBacklogSheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Priority Backlog", {
    views: [{ state: "frozen", ySplit: 1 }]
  });
  sheet.columns = COLUMNS.filter((c) => c.key !== "week").map((c) => ({
    header: c.header,
    key: c.key,
    width: c.width
  }));
  applyHeader(sheet.getRow(1));
  items.forEach((item, i) => {
    const row = sheet.addRow({
      action: item.action,
      keyword: item.keyword,
      intent: item.intent,
      journey: item.journey,
      priority_score: item.priority_score,
      cluster_rank: item.cluster_rank,
      cluster_url: item.cluster_url,
      cluster_search_volume: item.cluster_search_volume,
      cluster_opportunity: item.cluster_opportunity,
      cluster_difficulty: item.cluster_difficulty,
      content_type: item.content_type,
      needs_scraping: item.needs_scraping ? "Yes" : "No",
      reason: item.reason,
      supporting_kws: item.supporting_kws
    });
    applyDataRow(row, i, item);
  });
}

function buildSummarySheet(wb: ExcelJS.Workbook, items: CalendarItem[]): void {
  const sheet = wb.addWorksheet("Summary");
  sheet.columns = [
    { header: "Metric", key: "metric", width: 30 },
    { header: "Value", key: "value", width: 30 }
  ];
  applyHeader(sheet.getRow(1));

  const actionCounts: Record<string, number> = {};
  const journeyCounts: Record<string, number> = {};
  for (const item of items) {
    actionCounts[item.action] = (actionCounts[item.action] || 0) + 1;
    journeyCounts[item.journey] = (journeyCounts[item.journey] || 0) + 1;
  }

  sheet.addRow({ metric: "Total Calendar Items", value: items.length });
  sheet.addRow({ metric: "Total Weeks", value: Math.max(...items.map((i) => i.week ?? 0), 0) });
  sheet.addRow({ metric: "", value: "" });
  sheet.addRow({ metric: "Action Distribution", value: "" });
  for (const [a, c] of Object.entries(actionCounts)) {
    sheet.addRow({ metric: `  ${a}`, value: c });
  }
  sheet.addRow({ metric: "", value: "" });
  sheet.addRow({ metric: "Journey Stage Distribution", value: "" });
  for (const [j, c] of Object.entries(journeyCounts)) {
    sheet.addRow({ metric: `  ${j}`, value: c });
  }
}

export async function buildExcel(
  calendar: CalendarItem[],
  backlog: CalendarItem[]
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Keyword Insights";
  wb.created = new Date();
  buildCalendarSheet(wb, calendar);
  buildScrapeSheet(wb, calendar);
  buildBacklogSheet(wb, backlog);
  buildSummarySheet(wb, calendar);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

export function buildCsv(calendar: CalendarItem[]): string {
  const headers = [
    "week", "action", "keyword", "intent", "journey", "priority_score",
    "cluster_rank", "cluster_url", "cluster_search_volume", "cluster_opportunity",
    "cluster_difficulty", "content_type", "needs_scraping", "reason", "supporting_keywords"
  ];
  const rows = calendar.map((item) => [
    item.week, item.action, item.keyword, item.intent, item.journey,
    item.priority_score, item.cluster_rank, item.cluster_url,
    item.cluster_search_volume, item.cluster_opportunity, item.cluster_difficulty,
    item.content_type, item.needs_scraping ? "Yes" : "No", item.reason, item.supporting_kws
  ]);

  const escape = (v: unknown): string => {
    const s = String(v ?? "");
    return s.includes(",") || s.includes('"') || s.includes("\n")
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  };

  return [headers.join(","), ...rows.map((r) => r.map(escape).join(","))].join("\n");
}
