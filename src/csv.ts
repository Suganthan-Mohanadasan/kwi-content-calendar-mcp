import Papa from "papaparse";

export interface CsvStats {
  totalRows: number;
  uniqueClusters: number;
  uniqueHubs: number;
  hasLeadKeywords: boolean;
  detectedDomain: string;
}

export interface ParsedCsv {
  headers: string[];
  rows: Record<string, string>[];
}

export const REQUIRED_COLUMNS = [
  "keyword",
  "cluster",
  "cluster_url",
  "cluster_rank",
  "cluster_search_volume",
  "cluster_opportunity",
  "cluster_difficulty",
  "cluster_n_kwd",
  "cluster_dominant_intent",
  "cluster_content_score",
  "is_lead_keyword",
  "dominant_intent",
  "journey_classifications",
  "is_intent_match",
  "hub",
  "spoke",
  "url",
  "rank"
];

const LLM_COLUMNS = [
  "keyword",
  "cluster",
  "cluster_search_volume",
  "cluster_opportunity",
  "cluster_difficulty",
  "cluster_n_kwd",
  "cluster_dominant_intent",
  "cluster_content_score",
  "dominant_intent",
  "journey_classifications",
  "hub",
  "spoke",
  "url",
  "rank",
  "cluster_url",
  "cluster_rank",
  "cluster_blog_pages_avg",
  "cluster_product_pages_avg",
  "cluster_other_pages_avg",
  "cluster_serp_features",
  "cluster_social_sources"
];

export function parseCsv(text: string): ParsedCsv {
  const result = Papa.parse<Record<string, string>>(text, {
    header: true,
    skipEmptyLines: true
  });
  return {
    headers: result.meta.fields || [],
    rows: result.data || []
  };
}

export function validateCsv(parsed: ParsedCsv): { ok: true; stats: CsvStats } | { ok: false; error: string } {
  const missing = REQUIRED_COLUMNS.filter((c) => !parsed.headers.includes(c));
  if (missing.length > 0) {
    return { ok: false, error: `Missing required columns: ${missing.join(", ")}` };
  }
  const clusters = new Set<string>();
  const hubs = new Set<string>();
  const hostCounts = new Map<string, number>();
  let hasLead = false;

  for (const row of parsed.rows) {
    if (row.cluster) clusters.add(row.cluster);
    if (row.hub) hubs.add(row.hub);
    if (row.is_lead_keyword === "True") hasLead = true;
    const url = row.url;
    if (url) {
      try {
        const u = new URL(url);
        hostCounts.set(u.hostname, (hostCounts.get(u.hostname) || 0) + 1);
      } catch {
        // ignore
      }
    }
  }

  let detectedDomain = "";
  let max = 0;
  for (const [host, count] of hostCounts) {
    if (count > max) {
      max = count;
      detectedDomain = host.replace(/^www\./, "");
    }
  }

  return {
    ok: true,
    stats: {
      totalRows: parsed.rows.length,
      uniqueClusters: clusters.size,
      uniqueHubs: hubs.size,
      hasLeadKeywords: hasLead,
      detectedDomain
    }
  };
}

export function compactForLlm(parsed: ParsedCsv, maxClusters = 500): string {
  const byCluster = new Map<string, Record<string, string>[]>();
  for (const row of parsed.rows) {
    const c = row.cluster;
    if (!c) continue;
    if (!byCluster.has(c)) byCluster.set(c, []);
    byCluster.get(c)!.push(row);
  }

  const hubClusterCounts = new Map<string, number>();
  for (const row of parsed.rows) {
    if (row.is_lead_keyword === "True" && row.hub) {
      hubClusterCounts.set(row.hub, (hubClusterCounts.get(row.hub) || 0) + 1);
    }
  }

  const leadRows = parsed.rows.filter((r) => r.is_lead_keyword === "True");

  leadRows.sort((a, b) => {
    const ao = parseFloat(a.cluster_opportunity) || 0;
    const bo = parseFloat(b.cluster_opportunity) || 0;
    return bo - ao;
  });

  const capped = leadRows.slice(0, maxClusters);

  const enriched = capped.map((lead) => {
    const cluster = lead.cluster;
    const all = byCluster.get(cluster) || [];
    const ranking = all.filter((r) => r.url && r.url !== "" && Number(r.rank) < 100);
    const mismatch = ranking.filter((r) => r.is_intent_match === "False");
    const supporting = all
      .filter((r) => r.is_lead_keyword !== "True")
      .map((r) => r.keyword)
      .slice(0, 5)
      .join("|");

    const slim: Record<string, string> = {};
    for (const col of LLM_COLUMNS) {
      if (lead[col] !== undefined) {
        slim[col] = lead[col];
      }
    }
    slim._ranking_kw_count = String(ranking.length);
    slim._mismatch_count = String(mismatch.length);
    slim._supporting_kws = supporting;
    slim._hub_cluster_count = String(hubClusterCounts.get(lead.hub) || 0);

    return slim;
  });

  return Papa.unparse(enriched);
}
