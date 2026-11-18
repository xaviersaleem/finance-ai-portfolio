# Report Structure — Store_Analysis_Report.xlsx

Exact layout for the output workbook. The builder script writes the data tables and KPI rows; the analyst fills the narrative cells. Save to `Worker Outputs/Store_Analysis_Report.xlsx`.

## General formatting
- One worksheet per section, in the order below, tab names matching the section.
- Top of each sheet: section title (bold, larger), then a one-line purpose, then content.
- Currency formatted with thousands separators at the stated scale; percentages to one decimal.
- Highlight favorable variances green, unfavorable red (cost ratios: lower = favorable; revenue/EBITDA/margin: higher = favorable).
- Freeze header rows on table sheets. No placeholder text in the final file.
- First sheet is **Executive Summary**; final sheet is **Handoff Notes**.

## Sheet order
1. Executive Summary
2. Revenue Analysis
3. EBITDA Analysis
4. Labor Cost Analysis
5. Food Cost Analysis
6. Regional Analysis
7. Top Performing Stores
8. Underperforming Stores
9. Expansion Readiness Assessment
10. Key Risks
11. Key Opportunities
12. Handoff Notes

---

### 1. Executive Summary
Stands alone — the most important sheet for the FP&A Strategy Partner.
- KPI band: total revenue, blended EBITDA margin, portfolio labor %, food %, prime cost %, store count, region count.
- 4–7 bullet headline findings (best/worst, biggest driver, concentration, cost posture).
- One line each on overall risk posture and base strength.
- Note benchmarks used (briefing targets vs portfolio median) and any data caveats.

### 2. Revenue Analysis
- Table: store, region, revenue, revenue share %, (growth % if period exists), variance vs benchmark.
- Totals row. Narrative: distribution, concentration, leaders/laggards.

### 3. EBITDA Analysis
- Table: store, region, EBITDA, EBITDA margin %, variance vs benchmark, (growth if available).
- Blended margin row (Σ EBITDA ÷ Σ revenue). Narrative: margin spread, driver attribution.

### 4. Labor Cost Analysis
- Table: store, region, labor cost, labor cost %, variance vs benchmark.
- Portfolio labor % row. Narrative: where labor is in/out of control.

### 5. Food Cost Analysis
- Table: store, region, food cost, food cost %, variance vs benchmark.
- Portfolio food % row. Narrative: food-cost pressure points; outliers.

### 6. Regional Analysis
- Table by region: store count, revenue, revenue share, blended EBITDA margin, labor %, food %, prime cost %.
- Narrative: strongest/weakest regions, consistency within region, concentration.

### 7. Top Performing Stores
- Ranked table (best first): rank, store, region, composite score, EBITDA margin %, revenue, labor %, food %.
- Show the metrics behind the rank so it is auditable. Narrative: what makes them strong.

### 8. Underperforming Stores
- Ranked table (worst first): same columns. Narrative: the specific drag (revenue, labor, food, or margin) per store.

### 9. Expansion Readiness Assessment
Descriptive signals only — NOT a recommendation, NOT a forecast.
- Profitability depth (count/% of stores above EBITDA benchmark), consistency (margin & prime-cost dispersion), regional strength, cost control, concentration risk.
- Close with open questions for the FP&A Strategy Partner. Do not state "ready/not ready" or where to expand.

### 10. Key Risks
- Table: risk, evidence (stores/regions/figures), affected area, severity (High/Med/Low).

### 11. Key Opportunities
- Table: opportunity, evidence, where, potential area to investigate. Frame for the next worker; no projections.

### 12. Handoff Notes
For the FP&A Strategy Partner.
- Scope completed; benchmarks/thresholds used; data-quality caveats and assumptions.
- Consolidated risks and opportunities.
- Open questions an expansion evaluation must answer.
- Explicit statement: forecasting, investment analysis, and expansion recommendations are out of scope for this worker and belong to the FP&A Strategy Partner and CFO Advisor.
