#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { readFileSync, writeFileSync } from "fs";
import path from "path";

import { parseCsv, validateCsv, compactForLlm } from "./csv.js";
import { buildExcel, buildCsv } from "./excel.js";
import { ANALYSIS_PROMPT } from "./prompt.js";
import type { CalendarItem } from "./types.js";

const server = new McpServer({
  name: "kwi-content-calendar",
  version: "1.0.0"
});

// ─── Tool 1: Parse a KWI clustering CSV ───

server.tool(
  "parse_clustering_csv",
  "Parses a Keyword Insights clustering CSV file, validates it, compacts it to lead-keyword rows with derived columns, and returns the data ready for content calendar analysis. The response includes full analysis instructions for the LLM.",
  {
    csv_path: z.string().describe("Absolute path to the KWI clustering CSV file"),
    domain: z.string().describe("The client's domain (e.g. example.com)"),
    pieces_per_week: z.number().min(1).max(10).default(3).describe("Content pieces to schedule per week"),
    weeks: z.number().min(1).max(52).default(4).describe("Number of weeks for the calendar"),
    tofu_pct: z.number().min(0).max(100).default(45).describe("TOFU (awareness) percentage target"),
    mofu_pct: z.number().min(0).max(100).default(30).describe("MOFU (consideration) percentage target"),
    bofu_pct: z.number().min(0).max(100).default(25).describe("BOFU (conversion) percentage target"),
    max_clusters: z.number().min(50).max(5000).default(500).describe("Maximum clusters to send for analysis (top N by opportunity)")
  },
  async ({ csv_path, domain, pieces_per_week, weeks, tofu_pct, mofu_pct, bofu_pct, max_clusters }) => {
    try {
      const raw = readFileSync(csv_path, "utf8");
      const parsed = parseCsv(raw);
      const validation = validateCsv(parsed);

      if (!validation.ok) {
        return { content: [{ type: "text" as const, text: `CSV validation failed: ${validation.error}` }] };
      }

      const stats = validation.stats;
      const compact = compactForLlm(parsed, max_clusters);
      const actualClusters = compact.split("\n").length - 1; // minus header

      const userConfig = `
USER CONFIGURATION:
- Client domain: ${domain}
- Pieces per week: ${pieces_per_week}
- Number of weeks: ${weeks}
- Total calendar items needed: ${pieces_per_week * weeks}
- Target journey mix: TOFU ${tofu_pct}%, MOFU ${mofu_pct}%, BOFU ${bofu_pct}%
`;

      const output = `## KWI Clustering CSV Parsed Successfully

**Stats:**
- Total rows: ${stats.totalRows.toLocaleString()}
- Unique clusters: ${stats.uniqueClusters.toLocaleString()}
- Unique hubs: ${stats.uniqueHubs.toLocaleString()}
- Detected domain: ${stats.detectedDomain || "none"}
- Clusters sent for analysis: ${actualClusters} (top by opportunity)

${userConfig}

## Analysis Instructions

${ANALYSIS_PROMPT}

## Cluster Data (CSV format, ${actualClusters} lead-keyword rows)

${compact}

Now analyse this data following the instructions above and return the JSON object with "calendar" and "backlog" arrays.`;

      return { content: [{ type: "text" as const, text: output }] };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return { content: [{ type: "text" as const, text: `Error reading CSV: ${msg}` }] };
    }
  }
);

// ─── Tool 2: Export calendar to Excel + CSV files ───

server.tool(
  "export_content_calendar",
  "Takes a content calendar JSON (with 'calendar' and 'backlog' arrays of CalendarItem objects) and exports it as a formatted Excel workbook (4 sheets: Calendar, Pages to Scrape, Priority Backlog, Summary) and a flat CSV file.",
  {
    calendar_json: z.string().describe("JSON string with 'calendar' and 'backlog' arrays"),
    output_dir: z.string().describe("Directory to save the output files"),
    filename_prefix: z.string().default("content_calendar").describe("Prefix for output filenames")
  },
  async ({ calendar_json, output_dir, filename_prefix }) => {
    try {
      let data: { calendar?: unknown; backlog?: unknown };
      try {
        data = JSON.parse(calendar_json);
      } catch {
        // Try extracting JSON from markdown code fences
        const fenced = calendar_json.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
        if (fenced) {
          data = JSON.parse(fenced[1].trim());
        } else {
          const start = calendar_json.indexOf("{");
          const end = calendar_json.lastIndexOf("}");
          if (start !== -1 && end > start) {
            data = JSON.parse(calendar_json.slice(start, end + 1));
          } else {
            throw new Error("Could not find JSON object in input");
          }
        }
      }

      const calendar = Array.isArray(data.calendar) ? (data.calendar as CalendarItem[]) : [];
      const backlog = Array.isArray(data.backlog) ? (data.backlog as CalendarItem[]) : [];

      if (calendar.length === 0) {
        return { content: [{ type: "text" as const, text: "Error: calendar array is empty. The JSON must have a 'calendar' key with an array of items." }] };
      }

      const today = new Date().toISOString().slice(0, 10);
      const safePrefix = filename_prefix.replace(/[^a-z0-9_-]/gi, "_").toLowerCase();
      const baseName = `${safePrefix}_${today}`;

      // Build Excel
      const excelBuf = await buildExcel(calendar, backlog);
      const excelPath = path.join(output_dir, `${baseName}.xlsx`);
      writeFileSync(excelPath, excelBuf);

      // Build CSV
      const csvContent = buildCsv(calendar);
      const csvPath = path.join(output_dir, `${baseName}.csv`);
      writeFileSync(csvPath, csvContent, "utf8");

      // Summary stats
      const actions: Record<string, number> = {};
      const journeys: Record<string, number> = {};
      for (const item of calendar) {
        actions[item.action] = (actions[item.action] || 0) + 1;
        journeys[item.journey] = (journeys[item.journey] || 0) + 1;
      }

      const actionSummary = Object.entries(actions).map(([a, c]) => `  ${a}: ${c}`).join("\n");
      const journeySummary = Object.entries(journeys).map(([j, c]) => `  ${j}: ${c}`).join("\n");

      return {
        content: [{
          type: "text" as const,
          text: `Content calendar exported successfully.

**Files created:**
- Excel: ${excelPath}
- CSV: ${csvPath}

**Summary:**
- Calendar items: ${calendar.length}
- Backlog items: ${backlog.length}
- Weeks: ${Math.max(...calendar.map((i) => i.week ?? 0), 0)}

**Action distribution:**
${actionSummary}

**Journey stage distribution:**
${journeySummary}`
        }]
      };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return { content: [{ type: "text" as const, text: `Export error: ${msg}` }] };
    }
  }
);

// ─── Start ───

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});
