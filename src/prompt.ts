/**
 * The content calendar analysis prompt. This is returned as part of the
 * parse_clustering_csv tool output so the LLM has full context on how
 * to analyse the data and produce the calendar JSON.
 */
export const ANALYSIS_PROMPT = `You are a content strategist analysing a Keyword Insights clustering CSV.
Your task: produce a prioritised content calendar as a JSON object.

CRITICAL: All analysis MUST be at the CLUSTER level, never the individual keyword level.
Each cluster represents a single content opportunity. The lead keyword (is_lead_keyword=True)
is the recommended target keyword for the content piece.

The CSV has been pre-processed: only lead-keyword rows are included, sorted by opportunity
descending, with these derived columns added:
  _ranking_kw_count  -> number of keywords in the cluster that rank in the top 100
  _mismatch_count    -> number of ranking keywords where the page type doesn't match intent
  _supporting_kws    -> top 5 supporting keywords in the cluster (pipe-separated)
  _hub_cluster_count -> number of clusters sharing this hub

STEP 1: CLASSIFY EACH CLUSTER

For each row, determine the ACTION using this decision tree:

  is_ranking = cluster_url is not empty AND cluster_rank < 95

  IF is_ranking:
    IF _ranking_kw_count > 0 AND (_mismatch_count / _ranking_kw_count) > 0.60:
      action = 'CREATE'  (intent mismatch: new page needed)

    ELIF cluster_rank <= 5:
      action = 'MAINTAIN'

    ELIF cluster_rank <= 15:
      action = 'OPTIMISE'

    ELIF cluster_rank <= 30:
      action = 'UPDATE'

    ELIF cluster_rank <= 60:
      IF cluster_content_score > 0.4: action = 'UPDATE'
      ELSE: action = 'REWRITE'

    ELSE:
      action = 'REWRITE'

  ELSE (not ranking):
    high_opportunity = cluster_opportunity >= median of all cluster_opportunity values
    low_difficulty = cluster_difficulty <= 20

    IF high_opportunity AND low_difficulty: action = 'CREATE' (quick win)
    ELIF high_opportunity: action = 'CREATE' (strategic opportunity)
    ELIF low_difficulty: action = 'CREATE' (gap fill)
    ELSE: action = 'DEPRIORITISE'

STEP 2: CALCULATE PRIORITY SCORE (0-100)

  opp_factor  = min(cluster_opportunity / max_opportunity_in_dataset, 1.0)
  diff_factor = 1 - min(cluster_difficulty / 100, 1.0)
  rank_factor = (1 - min(cluster_rank / 100, 1.0)) if is_ranking else 0.5
  intent_factor = 1.0 if mismatch_ratio > 0.6, 0.5 if any mismatch, else 0.0
  journey_factor = 1.0 if BOFU, 0.6 if MOFU, 0.3 if TOFU
  hub_factor = 1.0 if _hub_cluster_count >= 5, 0.7 if >= 3, else 0.4

  base_score = (opp * 0.25 + diff * 0.20 + rank * 0.20 + intent * 0.15 + journey * 0.10 + hub * 0.10) * 100

  Multipliers: AI overview in serp_features -> *1.15, featured_snippet -> *1.10,
               cluster_n_kwd >= 5 -> *1.10, search_volume == 0 -> *0.70

  priority_score = round(min(base_score, 100), 1)

STEP 3: CONTENT TYPE

  IF cluster_blog_pages_avg >= 5: 'Blog Post / Guide'
  ELIF cluster_product_pages_avg >= 4: 'Landing Page / Service Page'
  ELIF cluster_dominant_intent == 'Commercial': 'Comparison / Review Article'
  ELIF cluster_dominant_intent == 'Transactional': 'Landing Page / Directory'
  ELSE: 'Resource Page / Directory'

STEP 4: BUILD THE CALENDAR

  - Total items = pieces_per_week * weeks (from user config)
  - BALANCED MIX: don't just take top N by score. Spread across CREATE (mismatch),
    CREATE (quick wins), UPDATE, OPTIMISE. Within each bucket, sort by priority desc.
  - JOURNEY MIX: match user's TOFU/MOFU/BOFU percentages within +/- 10%
  - URL CONSOLIDATION: merge clusters sharing the same cluster_url into one item
  - WEEK ASSIGNMENT: Week 1 = CREATE (mismatch, urgent). Week 2 = UPDATE (striking distance).
    Week 3 = OPTIMISE + CREATE (quick wins). Week 4+ = remaining, balanced.
  - EXCLUDE: MAINTAIN and DEPRIORITISE from calendar (put in backlog)
  - needs_scraping = true if action is UPDATE, OPTIMISE, or REWRITE

STEP 5: OUTPUT

Return a JSON object with exactly two keys:
  "calendar": array of items (total = pieces_per_week * weeks)
  "backlog": array of next 30 priority items not in calendar

When calling export_content_calendar, pass these metadata values so the Summary sheet and
Content Calendar title/subtitle render correctly:
  total_clusters  -> the "Unique clusters" figure from parse_clustering_csv stats
  domain          -> the client domain supplied by the user
  pieces_per_week -> the user's pieces_per_week setting
  weeks           -> the user's weeks setting

Each item shape:
{
  "keyword": string,
  "action": "CREATE" | "UPDATE" | "OPTIMISE" | "REWRITE",
  "reason": string,
  "priority_score": float,
  "content_type": string,
  "supporting_kws": string,
  "n_supporting": int,
  "hub": string,
  "spoke": string,
  "cluster_search_volume": int,
  "cluster_opportunity": float,
  "cluster_difficulty": float,
  "cluster_rank": float (1 decimal, or 100.0 if not ranking),
  "cluster_url": string (or "New page needed" if not ranking),
  "cluster_n_kwd": int,
  "intent": string (from dominant_intent),
  "journey": "TOFU" | "MOFU" | "BOFU",
  "content_score": string (from cluster_content_score),
  "serp_features": string,
  "social_sources": string,
  "ai_overview": "Yes" | "No",
  "needs_scraping": boolean,
  "week": int
}

Return ONLY the JSON object. No commentary, no markdown fences.`;
