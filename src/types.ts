export interface CalendarItem {
  keyword: string;
  action: "CREATE" | "UPDATE" | "OPTIMISE" | "REWRITE";
  reason: string;
  priority_score: number;
  content_type: string;
  supporting_kws: string;
  n_supporting: number;
  hub: string;
  spoke: string;
  cluster_search_volume: number;
  cluster_opportunity: number;
  cluster_difficulty: number;
  cluster_rank: number;
  cluster_url: string;
  cluster_n_kwd: number;
  intent: string;
  journey: "TOFU" | "MOFU" | "BOFU";
  content_score: string;
  serp_features: string;
  social_sources: string;
  ai_overview: string;
  needs_scraping: boolean;
  week: number;
}
