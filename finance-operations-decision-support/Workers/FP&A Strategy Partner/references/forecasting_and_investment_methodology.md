# Forecasting & Investment Methodology

How the FP&A Strategy Partner builds the expansion forecast and computes the investment metrics that decide whether expansion creates shareholder value. Standard corporate-finance and restaurant-unit-economics conventions apply.

## Contents
1. Principles and units
2. New-store economics from the current base
3. Applying market assumptions
4. Ramp-up modeling
5. Building the cash flow per scenario
6. Investment metrics (NPV, ROI, IRR, payback)
7. Thresholds from the Capital Investment Policy
8. Scenario & sensitivity analysis
9. Risk and capital-allocation efficiency
10. Choosing the recommended scenario (and the handoff line)

---

## 1. Principles and units

- **Ground the forecast in worker 1's reality.** New-store revenue and margins should start from the current portfolio's proven economics (especially strong regions), not optimistic round numbers. State where each input comes from.
- Confirm currency scale and the **forecast horizon** (usually set by the policy's payback/return window — commonly 5 years). Keep one horizon across all scenarios so they compare fairly.
- Confirm whether scenario figures are per-store or total, annual or cumulative, and whether costs are dollars or % of revenue. Normalize before modeling.
- Every assumption gets recorded on the Key Assumptions sheet with its source.

## 2. New-store economics from the current base

For a new store, derive a baseline annual P&L:
- **Revenue/store** — use the scenario's stated figure if given; otherwise anchor to the current portfolio median (or proven-region median) revenue per store from `Store_Analysis_Report.xlsx`.
- **EBITDA margin** — anchor to current portfolio blended margin, adjusted for assumptions (section 3). New stores often run below mature margin during ramp.
- **Mature store EBITDA $** = mature revenue × mature EBITDA margin.

## 3. Applying market assumptions

From `Market_Assumptions.xlsx` (Notion): inflation, labor inflation, market growth, industry outlook.
- **Revenue** grows with market-growth (and/or menu-price inflation) each year: `rev_t = rev_(t-1) × (1 + market_growth)`.
- **Costs** inflate: food cost with general inflation, labor cost with labor inflation. If margins are modeled directly, compress/expand margin by the net of cost inflation vs price inflation.
- **Industry outlook** is qualitative context — use it to temper or support growth assumptions and to inform the Risk and Opportunity sheets; do not double-count it as a number.
- If an assumption is missing, choose a clearly-stated conservative default and flag it.

## 4. Ramp-up modeling

New stores rarely hit mature volume in year 1. Apply a ramp curve to revenue (and often margin) unless the scenario specifies otherwise:
- Default ramp (share of mature revenue): Year 1 ~70%, Year 2 ~90%, Year 3+ 100%. Margins ramp similarly as fixed costs get absorbed.
- Stagger openings if the scenario implies phased rollout; otherwise assume stores open at the start of year 1.
- Document the ramp used — it materially moves payback.

## 5. Building the cash flow per scenario

For each scenario (5/10/15 stores) over the horizon:
1. **Capex (Year 0)** = stores × capital per store (or total capex if given). Negative cash flow.
2. **Annual operating cash flow_t** = Σ across stores of (revenue_t × EBITDA margin_t), applying ramp and assumption growth. Use EBITDA as the operating cash-flow proxy unless the policy/data specify maintenance capex or taxes — if they do, subtract them.
3. **Net cash flow_t** = operating cash flow_t (− any ongoing capex). Year 0 is the capex outflow.
4. Build the cumulative cash flow series for payback.

## 6. Investment metrics

Compute per scenario:

- **NPV** = Σ (net cash flow_t ÷ (1 + r)^t) − capex, where r = hurdle rate from policy.
  NPV > 0 → creates value at the hurdle rate.
- **ROI** = total undiscounted net operating cash flow over horizon ÷ total capex (or, simple ROI = (cumulative profit − capex) ÷ capex). State which definition you use; be consistent. Compare to the policy minimum ROI.
- **IRR** = discount rate where NPV = 0. Solve numerically; if no sign change (all-positive/all-negative flows) IRR is undefined — say so and rely on NPV/ROI.
- **Payback period** = time for cumulative net cash flow to turn positive (cross zero). Interpolate within the year for a fractional figure. Compare to the policy maximum payback.
- **Profitability index** = (NPV + capex) ÷ capex — useful for capital-efficiency ranking when capital is constrained.

## 7. Thresholds from the Capital Investment Policy

Extract and apply, do not invent:
- **Hurdle / discount rate** (for NPV and IRR comparison).
- **Minimum acceptable ROI**.
- **Maximum acceptable payback period**.
- **Capital ceiling / availability** — a scenario that exceeds available capital fails regardless of returns; note it.

A scenario is **financially attractive** when NPV > 0 at the hurdle rate, ROI ≥ minimum, payback ≤ maximum, and capex ≤ ceiling. Show pass/fail per criterion.

## 8. Scenario & sensitivity analysis

- **Across scenarios:** compare 5 vs 10 vs 15 on NPV, ROI, IRR, payback, PI, and total capex. Larger scenarios usually have higher absolute NPV but may have worse capital efficiency or longer payback — surface the trade-off.
- **Sensitivity (within scenarios):** flex the high-impact drivers — revenue per store, EBITDA margin, capex per store, and ramp speed — by ± a reasonable band (e.g., ±10–15%) and show how NPV/payback move. Identify the variable the case is most sensitive to; that is the key risk to manage.
- Optionally frame downside/base/upside cases using market-assumption ranges.

## 9. Risk and capital-allocation efficiency

- **Risk Assessment:** execution risk (opening many stores at once), market/cost risk (inflation, labor), demand risk (revenue ramp slower than assumed), concentration, and financing risk. Tie each to the sensitivity findings and worker 1's risks. Rate likelihood/impact.
- **Capital-allocation efficiency:** rank scenarios by NPV-per-dollar-of-capex (PI) and payback, not just absolute NPV. The "best" scenario maximizes risk-adjusted value per capital dollar within the policy constraints — this is the heart of the recommendation.

## 10. Choosing the recommended scenario

Select the scenario that best clears the policy thresholds with the strongest **risk-adjusted capital efficiency** — typically the one with positive NPV, ROI above minimum, payback within limit, and the best PI given acceptable risk. Explain the trade-off you made (e.g., "10-store has lower absolute NPV than 15-store but materially better payback and PI, and stays within the capital ceiling").

**Hold the handoff line:** present this as the financially most attractive option **for the CFO Advisor to decide on** — recommend on merit, but do not declare execution or write the board resolution. That is worker 3's role.
