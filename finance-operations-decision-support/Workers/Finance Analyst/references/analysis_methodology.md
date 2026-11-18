# Analysis Methodology

How the Finance Analyst computes every metric, sets benchmarks, runs variance, scores stores, and reads expansion-readiness signals. Restaurant-industry conventions apply throughout.

## Contents
1. Units and data hygiene
2. Core metrics and formulas
3. Benchmarks and thresholds
4. Variance analysis
5. Trend analysis
6. Store scoring and ranking
7. Expansion-readiness signals
8. Risks and opportunities

---

## 1. Units and data hygiene

Before any math, confirm from the workbook:
- **Currency scale** — raw dollars vs thousands vs millions. Keep one scale throughout and state it.
- **Percent encoding** — a "labor cost %" column may be a percent (28.5) or a decimal (0.285). Normalize.
- **Cost vs cost %** — labor and food may arrive as dollar amounts, as % of revenue, or both. If only dollars, derive the %; if only %, derive the dollars using revenue.
- **Period column** — is there a month/quarter/date? If yes, trend and growth are possible; if not, treat the data as a single snapshot and say so.
- **Completeness** — flag missing values, zero-revenue rows, duplicate store rows, and outliers (e.g., food cost % > 60% or < 15%) as data-quality notes rather than silently dropping them.

## 2. Core metrics and formulas

Per store, then aggregate to region and total portfolio.

| Metric | Formula |
|---|---|
| Revenue | as reported |
| Revenue share | store revenue ÷ total revenue |
| EBITDA | as reported |
| **EBITDA margin** | EBITDA ÷ revenue |
| Labor cost | as reported (or % × revenue) |
| **Labor cost %** | labor cost ÷ revenue |
| Food cost | as reported (or % × revenue) |
| **Food cost %** | food cost ÷ revenue |
| **Prime cost %** | labor cost % + food cost % |
| Revenue growth* | (current − prior) ÷ prior |
| EBITDA growth* | (current − prior) ÷ prior |

*Growth metrics only when a period dimension exists.

Aggregate margins are computed on summed dollars (Σ EBITDA ÷ Σ revenue), **not** by averaging store-level percentages — averaging percentages misweights small stores.

## 3. Benchmarks and thresholds

**Always prefer thresholds stated in the Executive Briefing.** If the briefing gives target food cost %, labor cost %, or EBITDA margin, use those as the benchmark and say so. Only if the briefing is silent, fall back to:

- **Internal benchmark (primary fallback):** the portfolio **median** for each metric. Median resists outliers better than mean for a store fleet. Report mean too for context.
- **Industry rules of thumb (context only, label clearly as generic):**
  - Food cost %: ~28–35% of revenue is typical full-service.
  - Labor cost %: ~25–35% of revenue.
  - Prime cost %: ~55–65%; above ~65% pressures profitability.
  - EBITDA margin: healthy restaurant-level margins often run ~15–25%, but vary widely by format.

Treat industry numbers as orientation, not verdicts — the briefing and the portfolio's own distribution are the real standard.

## 4. Variance analysis

For each metric and each store, report variance two ways when data allows:

1. **vs benchmark** — store value − benchmark (briefing target, else portfolio median). For cost ratios, *lower than benchmark is favorable*; for revenue, EBITDA, and margin, *higher is favorable*. Always state direction in plain terms (favorable/unfavorable), not just sign.
2. **vs prior period** — current − prior, if a period dimension exists.

Express variance in both absolute terms and percentage points (for ratios) or percent (for levels). Then **decompose the driver**: is an EBITDA-margin miss driven by labor %, food %, or soft revenue? Point to the specific lever. This driver attribution is the most useful thing you hand the next worker.

## 5. Trend analysis

If a period dimension exists, for the portfolio and each region describe direction (improving/flat/deteriorating), magnitude, and consistency (steady vs volatile) for revenue, EBITDA margin, labor %, and food %. Call out inflections. **Describe** the trend — do not extrapolate it into a forecast.

If no period dimension exists, state that the data is a single snapshot and base trend commentary only on cross-sectional spread (dispersion across stores), clearly labeled as such.

## 6. Store scoring and ranking

Rank stores to populate Top Performing and Underperforming sheets. Use a transparent composite so the next worker can see why a store landed where it did.

Suggested composite (equal-weight z-scores across the portfolio; flip sign so higher = better):
- + EBITDA margin
- + Revenue (scale matters for a base)
- − Labor cost %
- − Food cost %
- (+ revenue growth, if available)

Compute each store's z-score per component, sum into a composite, rank descending. **Top performers** = highest composite; **underperformers** = lowest. Always show the underlying metrics alongside the rank so the ranking is auditable, and note any store that is strong on one dimension but dragged by another (e.g., high revenue but prime-cost heavy). Avoid over-engineering — if the briefing implies a primary KPI (often EBITDA margin), lead with it and use the composite as support.

## 7. Expansion-readiness signals

This worker assesses the **strength and stability of the current base** — the raw material a later worker uses to judge expansion. You are reading signals, **not** making the expansion call.

Signals to summarize (descriptive only):
- **Profitability depth** — how many stores clear the EBITDA-margin benchmark, and by how much. A deep bench of profitable stores is a positive base signal.
- **Consistency** — dispersion of margins and prime cost across stores. Tight, healthy distribution = stronger base; wide scatter = fragility.
- **Regional strength** — which regions have multiple healthy stores (a proven, repeatable model) vs single-store or weak regions.
- **Cost control** — labor % and food % under control portfolio-wide vs systemic pressure.
- **Concentration risk** — is performance propped up by a few stores/one region?

Write these as observations ("X of N stores exceed the EBITDA benchmark; Region A shows the tightest cost control"). Do **not** conclude "ready/not ready to expand" and do **not** name where to expand — record open questions for the FP&A Strategy Partner instead.

## 8. Risks and opportunities

- **Key Risks** — concrete threats to current performance: underperforming clusters, prime-cost creep, revenue concentration, margin volatility, data gaps, single points of failure. Tie each to evidence.
- **Key Opportunities** — observable upside in the *current* base: stores or regions running well below cost benchmarks (margin headroom), high-revenue stores with fixable cost structures, replicable strong-region patterns. Frame as "what the next worker should investigate," not as a recommendation or projection.
