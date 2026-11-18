# Report Structure — Expansion_Business_Case.xlsx

Exact layout for the output workbook. The builder script writes data tables and KPI rows; the analyst writes the narrative and the scenario call. Save to `Worker Outputs/Expansion_Business_Case.xlsx`.

## General formatting
- One worksheet per section, in order, tab names matching the section.
- Top of each sheet: title (bold, larger), one-line purpose, then content.
- Currency with thousands separators at the stated scale; percentages to one decimal; years to one decimal.
- Pass/fail vs policy thresholds: favorable green, unfavorable red.
- Freeze header rows on table sheets. No placeholder text in the final file.
- First sheet is **Executive Summary**; final sheet is **Handoff Notes**.

## Sheet order
1. Executive Summary
2. Expansion Forecast
3. Scenario Analysis
4. ROI Analysis
5. Payback Analysis
6. Risk Assessment
7. Opportunity Assessment
8. Recommended Scenario
9. Key Assumptions
10. Handoff Notes

---

### 1. Executive Summary
Stands alone — the key read for the CFO Advisor.
- KPI band per scenario: total capex, NPV, ROI, IRR, payback (years), pass/fail vs policy.
- Headline verdict: is expansion financially attractive, and which scenario leads.
- 4–7 bullets: the trade-off across scenarios, the binding constraint (capital/payback/return), the key sensitivity, the leading risk.
- Thresholds used (hurdle rate, min ROI, max payback, capital ceiling) and their source.

### 2. Expansion Forecast
- Per scenario, a year-by-year table over the horizon: stores open, revenue, EBITDA margin, operating cash flow, net cash flow, cumulative cash flow.
- Narrative: how the forecast was built (base economics, ramp, growth/inflation applied).

### 3. Scenario Analysis
- Side-by-side comparison table: 5 / 10 / 15-store on capex, NPV, ROI, IRR, payback, profitability index.
- Sensitivity block: NPV and payback under ±band on revenue/store, EBITDA margin, capex/store, ramp.
- Narrative: trade-offs, where each scenario wins/loses, most sensitive driver.

### 4. ROI Analysis
- Per scenario: total capex, cumulative net cash flow, ROI %, vs minimum-ROI threshold (pass/fail).
- Narrative: which scenarios clear the return bar and by how much.

### 5. Payback Analysis
- Per scenario: cumulative cash flow path, payback period (years, interpolated), vs max-payback threshold (pass/fail).
- Narrative: how ramp and capex intensity drive payback differences.

### 6. Risk Assessment
- Table: risk, driver/evidence, likelihood (H/M/L), impact (H/M/L), affected scenario(s), mitigation.
- Tie to sensitivity results and worker 1's risks.

### 7. Opportunity Assessment
- Table: opportunity, evidence (strong regions, margin headroom, market growth), which scenario captures it, upside area.

### 8. Recommended Scenario
- State the financially most attractive scenario and the metrics behind it.
- Show it clears each policy threshold; explain the capital-efficiency/risk trade-off vs the alternatives.
- Frame as a recommendation **for the CFO Advisor to decide on** — NOT a final decision, NOT a board resolution.

### 9. Key Assumptions
- Table: assumption, value, source (Capital Investment Policy / Market_Assumptions / Store_Analysis_Report / analyst default), notes.
- Include hurdle rate, min ROI, max payback, capital ceiling, inflation, labor inflation, market growth, ramp curve, revenue/store, EBITDA margin, capex/store, horizon.

### 10. Handoff Notes
For the CFO Advisor.
- Verdict and leading scenario; thresholds applied; key sensitivities and risks.
- All assumptions and sources; data-quality caveats.
- Open questions for the executive decision (strategic fit, financing, board appetite, phasing).
- Explicit statement: board recommendations and the final investment decision are out of scope for this worker and belong to the CFO Advisor.
