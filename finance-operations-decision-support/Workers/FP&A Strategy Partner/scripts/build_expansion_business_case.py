#!/usr/bin/env python3
"""
build_expansion_business_case.py
Crave Kitchen Finance Team - FP&A Strategy Partner (worker 2).

Reads the Finance Analyst's Store_Analysis_Report.xlsx plus Expansion_Scenarios.xlsx
(Google Sheets: 5/10/15-store scenarios) and Market_Assumptions.xlsx (Notion:
inflation, labor, market growth, industry outlook), forecasts each scenario, computes
NPV / ROI / IRR / payback / profitability index, runs a sensitivity pass, scores each
scenario against the Capital Investment Policy thresholds, and writes a formatted
Expansion_Business_Case.xlsx.

This worker EVALUATES and recommends a scenario on financial merit. It does NOT write
board recommendations or make the final investment decision - that is the CFO Advisor.

The script is schema-flexible: it fuzzy-matches columns and prints the mapping and the
assumptions it used, plus data-quality warnings. Override guesses with --*-col flags and
set policy thresholds with --hurdle-rate / --min-roi / --max-payback-years.

Usage:
  python3 build_expansion_business_case.py \
      --analysis    "Worker Outputs/Store_Analysis_Report.xlsx" \
      --scenarios   "Source Data/Expansion_Scenarios.xlsx" \
      --assumptions "Source Data/Market_Assumptions.xlsx" \
      --output      "Worker Outputs/Expansion_Business_Case.xlsx" \
      --hurdle-rate 0.12 --max-payback-years 4 --min-roi 0.20

Requires: pandas, numpy, openpyxl.
  pip install pandas numpy openpyxl --break-system-packages
"""
import argparse
import sys
import re
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    sys.exit("Missing pandas/numpy. Install: pip install pandas numpy openpyxl --break-system-packages")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing openpyxl. Install: pip install openpyxl --break-system-packages")


# ---------------- helpers: detection ----------------

def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def find_col(columns, keys, override=None):
    if override:
        for c in columns:
            if _norm(c) == _norm(override):
                return c
        for c in columns:
            if _norm(override) in _norm(c):
                return c
        print(f"  ! override '{override}' not found; auto-detecting")
    ncols = {c: _norm(c) for c in columns}
    for key in keys:
        for c, n in ncols.items():
            if key == n or re.search(rf"\b{re.escape(key)}\b", n):
                return c
    for key in keys:
        for c, n in ncols.items():
            if key in n:
                return c
    return None


def num(v):
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        return float(str(v).replace(",", "").replace("$", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return np.nan


def first_numeric(df, col):
    s = df[col].apply(num).dropna() if col in df.columns else pd.Series([], dtype=float)
    return s


def scan_assumptions(df):
    """Pull key rates from a 2-column-ish Notion table (label, value) or wide table."""
    found = {}
    targets = {
        "inflation": ["inflation"],
        "labor_inflation": ["labor inflation", "labour inflation", "wage", "labor growth", "labor"],
        "market_growth": ["market growth", "growth rate", "industry growth", "same store", "comp growth", "market"],
    }
    # try label/value pairs across all cells
    rows = df.values.tolist()
    for row in rows:
        cells = [str(c) for c in row]
        for i, cell in enumerate(cells):
            n = _norm(cell)
            for key, kws in targets.items():
                if key in found:
                    continue
                if any(kw in n for kw in kws):
                    # look right for a number
                    for j in range(i + 1, len(row)):
                        v = num(row[j])
                        if not np.isnan(v):
                            found[key] = v / 100.0 if v > 1.5 else v
                            break
    return found


# ---------------- finance math ----------------

def npv(rate, flows):
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(flows))


def irr(flows, lo=-0.9, hi=2.0):
    f = [num(x) for x in flows]
    if not any(c > 0 for c in f) or not any(c < 0 for c in f):
        return np.nan
    a, b = lo, hi
    fa, fb = npv(a, f), npv(b, f)
    if fa * fb > 0:
        return np.nan
    for _ in range(200):
        m = (a + b) / 2
        fm = npv(m, f)
        if abs(fm) < 1e-6:
            return m
        if fa * fm < 0:
            b, fb = m, fm
        else:
            a, fa = m, fm
    return (a + b) / 2


def payback_years(flows):
    """flows[0] is the capex outflow (negative). Returns interpolated years or nan."""
    cum = 0.0
    prev = 0.0
    for t, cf in enumerate(flows):
        prev = cum
        cum += cf
        if cum >= 0 and t > 0:
            need = -prev
            return (t - 1) + (need / cf) if cf else float(t)
    return np.nan


def build_scenario_flows(n_stores, rev_per_store, ebitda_margin, capex_per_store,
                         years, ramp, market_growth, margin_drift=0.0):
    """Return net cash flow list with flows[0] = -capex."""
    capex = n_stores * capex_per_store
    flows = [-capex]
    for t in range(1, years + 1):
        ramp_factor = ramp[t - 1] if t - 1 < len(ramp) else 1.0
        rev = rev_per_store * ((1 + market_growth) ** (t - 1)) * ramp_factor
        margin = ebitda_margin * ramp_factor + margin_drift * (t - 1)
        margin = max(0.0, min(margin, 0.6))
        ocf = n_stores * rev * margin
        flows.append(ocf)
    return capex, flows


# ---------------- styling ----------------

TITLE = Font(bold=True, size=16, color="1F3864")
H = Font(bold=True, size=11, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F3864")
SUB = Font(bold=True, size=11, color="1F3864")
TOTAL = Font(bold=True)
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NOTE = Font(italic=True, color="808080")


def ws_title(ws, title, purpose):
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = purpose
    ws["A2"].font = NOTE
    ws.freeze_panes = "A5"


def write_table(ws, start_row, headers, rows, pct_cols=None, money_cols=None, yr_cols=None, totals=None):
    pct_cols, money_cols, yr_cols = pct_cols or set(), money_cols or set(), yr_cols or set()
    r = start_row
    for j, head in enumerate(headers, 1):
        c = ws.cell(row=r, column=j, value=head)
        c.font = H; c.fill = HFILL
        c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = BORDER
    for row in rows:
        r += 1
        for j, val in enumerate(row, 1):
            c = ws.cell(row=r, column=j, value=val); c.border = BORDER
            if isinstance(val, (int, float)) and not (isinstance(val, float) and np.isnan(val)):
                if j - 1 in pct_cols: c.number_format = "0.0%"
                elif j - 1 in money_cols: c.number_format = "#,##0"
                elif j - 1 in yr_cols: c.number_format = "0.0"
    if totals:
        r += 1
        for j, val in enumerate(totals, 1):
            c = ws.cell(row=r, column=j, value=val); c.font = TOTAL; c.border = BORDER
            if isinstance(val, (int, float)) and not (isinstance(val, float) and np.isnan(val)):
                if j - 1 in pct_cols: c.number_format = "0.0%"
                elif j - 1 in money_cols: c.number_format = "#,##0"
                elif j - 1 in yr_cols: c.number_format = "0.0"
    for j, head in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(30, len(str(head)) + 3))
    return r + 1


def narrative(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 3, end_column=span)
    return row + 5


def clean(v):
    return "" if (v is None or (isinstance(v, float) and np.isnan(v))) else v


# ---------------- main ----------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--analysis", required=False, help="Store_Analysis_Report.xlsx from worker 1")
    ap.add_argument("--scenarios", required=True, help="Expansion_Scenarios.xlsx")
    ap.add_argument("--assumptions", required=False, help="Market_Assumptions.xlsx")
    ap.add_argument("--output", required=True)
    # scenario workbook column overrides
    ap.add_argument("--scenario-col", default=None)
    ap.add_argument("--stores-col", default=None)
    ap.add_argument("--capex-col", default=None)
    ap.add_argument("--revenue-col", default=None)
    ap.add_argument("--margin-col", default=None)
    # policy thresholds + model knobs
    ap.add_argument("--hurdle-rate", type=float, default=0.12)
    ap.add_argument("--min-roi", type=float, default=0.20)
    ap.add_argument("--max-payback-years", type=float, default=4.0)
    ap.add_argument("--capital-ceiling", type=float, default=None)
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--ramp", default="0.70,0.90,1.0", help="comma ramp factors yr1,yr2,yr3+")
    ap.add_argument("--market-growth", type=float, default=None, help="override; else from assumptions")
    args = ap.parse_args()

    ramp = [float(x) for x in args.ramp.split(",")]

    # ---- assumptions ----
    asm = {"inflation": np.nan, "labor_inflation": np.nan, "market_growth": np.nan}
    if args.assumptions and Path(args.assumptions).exists():
        adf = pd.ExcelFile(args.assumptions)
        merged = pd.concat([adf.parse(s, header=None) for s in adf.sheet_names], ignore_index=True)
        asm.update(scan_assumptions(merged))
        print("Assumptions scanned from Market_Assumptions.xlsx:", {k: v for k, v in asm.items() if not (isinstance(v, float) and np.isnan(v))})
    market_growth = args.market_growth if args.market_growth is not None else (
        asm.get("market_growth") if not np.isnan(asm.get("market_growth", np.nan)) else 0.03)

    # ---- base economics from worker 1 ----
    base_rev = np.nan
    base_margin = np.nan
    if args.analysis and Path(args.analysis).exists():
        try:
            axl = pd.ExcelFile(args.analysis)
            # try the EBITDA / Revenue sheets to estimate medians
            for sh in axl.sheet_names:
                d = axl.parse(sh)
                d.columns = [str(c) for c in d.columns]
                rc = find_col(d.columns, ["revenue", "sales"])
                mc = find_col(d.columns, ["ebitda margin", "margin"])
                if rc and np.isnan(base_rev):
                    s = first_numeric(d, rc)
                    if len(s): base_rev = float(s.median())
                if mc and np.isnan(base_margin):
                    s = first_numeric(d, mc)
                    if len(s):
                        m = float(s.median()); base_margin = m if m <= 1.5 else m / 100.0
            print(f"Base economics from Store_Analysis_Report: median rev/store={clean(base_rev)}, median EBITDA margin={clean(base_margin)}")
        except Exception as e:
            print(f"  ! could not parse analysis report ({e}); will rely on scenario file")

    # ---- scenarios ----
    sxl = pd.ExcelFile(args.scenarios)
    sdf = sxl.parse(sxl.sheet_names[0])
    sdf.columns = [str(c).strip() for c in sdf.columns]
    cols = list(sdf.columns)
    print(f"\nLoaded scenarios '{Path(args.scenarios).name}' sheet '{sxl.sheet_names[0]}': {len(sdf)} rows")
    print(f"Columns: {cols}")

    cmap = {
        "scenario": find_col(cols, ["scenario", "case", "option", "plan", "name"], args.scenario_col),
        "stores": find_col(cols, ["stores", "number of stores", "store count", "units", "new stores", "count"], args.stores_col),
        "capex": find_col(cols, ["capex", "capital", "investment", "cost per store", "capital per store", "build cost", "total capex"], args.capex_col),
        "revenue": find_col(cols, ["revenue per store", "revenue", "sales per store", "sales", "aus", "unit volume"], args.revenue_col),
        "margin": find_col(cols, ["ebitda margin", "margin", "ebitda"], args.margin_col),
    }
    print("Scenario column mapping:")
    for k, v in cmap.items():
        print(f"  {k:9s} -> {v}")
    print("Review mapping; re-run with --<field>-col to override.\n")

    if not cmap["stores"]:
        sys.exit("Could not find a store-count column. Pass --stores-col.")
    if not cmap["capex"]:
        sys.exit("Could not find a capex/capital column. Pass --capex-col.")

    warnings = []
    scenarios = []
    for idx, row in sdf.iterrows():
        n = num(row[cmap["stores"]])
        if np.isnan(n) or n <= 0:
            continue
        n = int(round(n))
        name = str(row[cmap["scenario"]]) if cmap["scenario"] else f"{n}-store"
        capex_raw = num(row[cmap["capex"]])
        # decide per-store vs total capex: if value is large relative to plausible per-store, treat as total
        if cmap["capex"] and not np.isnan(capex_raw):
            capex_per_store = capex_raw / n if capex_raw > 5_000_000 else capex_raw
            # if column name says total
            if "total" in _norm(cmap["capex"]):
                capex_per_store = capex_raw / n
        else:
            capex_per_store = np.nan
        rev_ps = num(row[cmap["revenue"]]) if cmap["revenue"] else np.nan
        if np.isnan(rev_ps):
            rev_ps = base_rev
        marg = np.nan
        if cmap["margin"]:
            mraw = num(row[cmap["margin"]])
            marg = mraw if (not np.isnan(mraw) and mraw <= 1.5) else (mraw / 100.0 if not np.isnan(mraw) else np.nan)
        if np.isnan(marg):
            marg = base_margin if not np.isnan(base_margin) else 0.18
        if np.isnan(rev_ps):
            warnings.append(f"{name}: no revenue/store available; cannot model. Provide --revenue-col or worker-1 report.")
            continue
        if np.isnan(capex_per_store):
            warnings.append(f"{name}: missing capex; cannot model.")
            continue
        scenarios.append(dict(name=name, n=n, rev_ps=rev_ps, margin=marg, capex_ps=capex_per_store))

    if not scenarios:
        sys.exit("No modelable scenarios found. Check column mapping / inputs.")

    # ---- compute metrics ----
    results = []
    for s in scenarios:
        capex, flows = build_scenario_flows(
            s["n"], s["rev_ps"], s["margin"], s["capex_ps"],
            args.years, ramp, market_growth)
        scen_npv = npv(args.hurdle_rate, flows)
        scen_irr = irr(flows)
        pb = payback_years(flows)
        total_ocf = sum(flows[1:])
        roi = (total_ocf - capex) / capex if capex else np.nan
        pi = (scen_npv + capex) / capex if capex else np.nan
        passes = {
            "npv": scen_npv > 0,
            "roi": (not np.isnan(roi)) and roi >= args.min_roi,
            "payback": (not np.isnan(pb)) and pb <= args.max_payback_years,
            "capital": (args.capital_ceiling is None) or (capex <= args.capital_ceiling),
        }
        results.append(dict(**s, capex=capex, flows=flows, npv=scen_npv, irr=scen_irr,
                            payback=pb, roi=roi, pi=pi, total_ocf=total_ocf, passes=passes,
                            all_pass=all(passes.values())))

    # pick recommended: prefer all-pass, then best PI, tie-break payback
    eligible = [r for r in results if r["all_pass"]] or results
    recommended = sorted(eligible, key=lambda r: (-(r["pi"] if not np.isnan(r["pi"]) else -9),
                                                  r["payback"] if not np.isnan(r["payback"]) else 99))[0]

    if warnings:
        print("DATA-QUALITY WARNINGS:")
        for w in warnings:
            print("  ! " + w)
        print()

    # ---------------- workbook ----------------
    wb = Workbook()

    def passfill(cell, ok):
        cell.fill = GREEN if ok else RED

    # 1 Executive Summary
    ws = wb.active
    ws.title = "Executive Summary"
    ws_title(ws, "Expansion Business Case - Executive Summary",
             "Crave Kitchen Group | FP&A Strategy Partner (worker 2) | for the CFO Advisor")
    hdr = ["Scenario", "Stores", "Total capex", "NPV @ hurdle", "ROI", "IRR", "Payback (yrs)", "PI", "Meets policy?"]
    rows = []
    for r in results:
        rows.append([r["name"], r["n"], r["capex"], r["npv"], r["roi"], clean(r["irr"]),
                     clean(r["payback"]), r["pi"], "PASS" if r["all_pass"] else "FAIL"])
    end = write_table(ws, 5, hdr, rows, pct_cols={4, 5}, money_cols={2, 3}, yr_cols={6, 7})
    # color the meets-policy col
    for i, r in enumerate(results):
        passfill(ws.cell(row=6 + i, column=9), r["all_pass"])
    ws.cell(row=end, column=1,
            value=f"Policy thresholds applied - hurdle rate {args.hurdle_rate:.1%}, min ROI {args.min_roi:.1%}, "
                  f"max payback {args.max_payback_years:.1f} yrs"
                  + (f", capital ceiling {args.capital_ceiling:,.0f}" if args.capital_ceiling else "")
                  + f". Market growth {market_growth:.1%}; horizon {args.years} yrs; ramp {ramp}.").font = NOTE
    end += 2
    narrative(ws, end,
        f"[ANALYST: state whether expansion is financially attractive and that the analysis points to "
        f"'{recommended['name']}' as the most attractive on merit. Give 4-7 bullets on the scenario trade-off, "
        f"the binding constraint, the key sensitivity, and the leading risk. Recommend for the CFO Advisor to "
        f"decide - do NOT declare execution or write a board resolution.]")

    # 2 Expansion Forecast
    ws = wb.create_sheet("Expansion Forecast")
    ws_title(ws, "Expansion Forecast", "Year-by-year projected economics per scenario.")
    r0 = 5
    for r in results:
        ws.cell(row=r0, column=1, value=r["name"]).font = SUB
        r0 += 1
        yhdr = ["Year"] + [f"Y{t}" for t in range(0, args.years + 1)]
        # rows: net cash flow + cumulative
        cum, cumrow = 0.0, []
        for cf in r["flows"]:
            cum += cf; cumrow.append(cum)
        for label, series in [("Net cash flow", r["flows"]), ("Cumulative", cumrow)]:
            rr = [label] + [round(x) for x in series]
            r0 = write_table(ws, r0, yhdr, [rr], money_cols=set(range(1, args.years + 2))) if False else r0
        # simpler: write two rows manually
        for j, h in enumerate(yhdr, 1):
            c = ws.cell(row=r0, column=j, value=h); c.font = H; c.fill = HFILL; c.border = BORDER
        r0 += 1
        ws.cell(row=r0, column=1, value="Net cash flow").border = BORDER
        for t, cf in enumerate(r["flows"]):
            c = ws.cell(row=r0, column=2 + t, value=round(cf)); c.number_format = "#,##0"; c.border = BORDER
        r0 += 1
        ws.cell(row=r0, column=1, value="Cumulative").border = BORDER
        for t, cf in enumerate(cumrow):
            c = ws.cell(row=r0, column=2 + t, value=round(cf)); c.number_format = "#,##0"; c.border = BORDER
        r0 += 2
    narrative(ws, r0,
        "[ANALYST: explain how the forecast was built - base per-store economics from worker 1, ramp curve, "
        "and the inflation/market-growth assumptions applied. Note Year 0 = capex outflow.]")

    # 3 Scenario Analysis (comparison + sensitivity)
    ws = wb.create_sheet("Scenario Analysis")
    ws_title(ws, "Scenario Analysis", "Side-by-side comparison and sensitivity.")
    comp = [[r["name"], r["n"], r["capex"], r["npv"], r["roi"], clean(r["irr"]), clean(r["payback"]), r["pi"]] for r in results]
    end = write_table(ws, 5, ["Scenario", "Stores", "Capex", "NPV", "ROI", "IRR", "Payback (yrs)", "PI"],
                      comp, pct_cols={4, 5}, money_cols={2, 3}, yr_cols={6})
    end += 1
    ws.cell(row=end, column=1, value="Sensitivity of NPV (recommended scenario)").font = SUB
    end += 1
    rec = recommended
    sens_rows = []
    for drv, lo_hi in [("Revenue/store", "rev_ps"), ("EBITDA margin", "margin"), ("Capex/store", "capex_ps")]:
        line = [drv]
        for delta in (-0.10, 0.0, 0.10):
            kw = dict(n_stores=rec["n"], rev_per_store=rec["rev_ps"], ebitda_margin=rec["margin"],
                      capex_per_store=rec["capex_ps"], years=args.years, ramp=ramp, market_growth=market_growth)
            if lo_hi == "rev_ps": kw["rev_per_store"] *= (1 + delta)
            elif lo_hi == "margin": kw["ebitda_margin"] *= (1 + delta)
            elif lo_hi == "capex_ps": kw["capex_per_store"] *= (1 + delta)
            _, fl = build_scenario_flows(**kw)
            line.append(round(npv(args.hurdle_rate, fl)))
        sens_rows.append(line)
    end = write_table(ws, end, ["Driver (-10% / base / +10%)", "-10%", "Base", "+10%"],
                      sens_rows, money_cols={1, 2, 3})
    narrative(ws, end + 1,
        "[ANALYST: discuss the scenario trade-off (absolute NPV vs capital efficiency vs payback) and identify "
        "the driver the case is most sensitive to - that is the key risk to manage.]")

    # 4 ROI Analysis
    ws = wb.create_sheet("ROI Analysis")
    ws_title(ws, "ROI Analysis", f"Return on invested capital vs policy minimum ({args.min_roi:.0%}).")
    rows = [[r["name"], r["capex"], r["total_ocf"], r["roi"], "PASS" if r["passes"]["roi"] else "FAIL"] for r in results]
    end = write_table(ws, 5, ["Scenario", "Total capex", "Cumulative net cash flow", "ROI", f"vs min {args.min_roi:.0%}"],
                      rows, pct_cols={3}, money_cols={1, 2})
    for i, r in enumerate(results): passfill(ws.cell(row=6 + i, column=5), r["passes"]["roi"])
    narrative(ws, end + 1, "[ANALYST: which scenarios clear the return bar and by how much.]")

    # 5 Payback Analysis
    ws = wb.create_sheet("Payback Analysis")
    ws_title(ws, "Payback Analysis", f"Years to recover capital vs policy max ({args.max_payback_years:.1f}).")
    rows = [[r["name"], r["capex"], clean(r["payback"]), "PASS" if r["passes"]["payback"] else "FAIL"] for r in results]
    end = write_table(ws, 5, ["Scenario", "Total capex", "Payback (yrs)", f"vs max {args.max_payback_years:.1f} yrs"],
                      rows, money_cols={1}, yr_cols={2})
    for i, r in enumerate(results): passfill(ws.cell(row=6 + i, column=4), r["passes"]["payback"])
    narrative(ws, end + 1, "[ANALYST: how ramp and capex intensity drive the payback differences.]")

    # 6 Risk Assessment
    ws = wb.create_sheet("Risk Assessment")
    ws_title(ws, "Risk Assessment", "Threats to the expansion case, tied to sensitivity and worker-1 risks.")
    write_table(ws, 5, ["Risk", "Driver / evidence", "Likelihood (H/M/L)", "Impact (H/M/L)", "Scenario(s)", "Mitigation"],
                [["", "", "", "", "", ""] for _ in range(8)])
    ws.cell(row=4, column=1, value="[ANALYST: fill from forecasting_and_investment_methodology.md section 9 - execution, market/cost, demand-ramp, concentration, financing.]").font = NOTE

    # 7 Opportunity Assessment
    ws = wb.create_sheet("Opportunity Assessment")
    ws_title(ws, "Opportunity Assessment", "Upside the expansion could capture.")
    write_table(ws, 5, ["Opportunity", "Evidence", "Scenario that captures it", "Upside area"],
                [["", "", "", ""] for _ in range(8)])
    ws.cell(row=4, column=1, value="[ANALYST: strong regions, margin headroom, market growth, first-mover - tie to worker-1 opportunities.]").font = NOTE

    # 8 Recommended Scenario
    ws = wb.create_sheet("Recommended Scenario")
    ws_title(ws, "Recommended Scenario",
             "Most financially attractive option - a recommendation for the CFO Advisor, NOT a final decision.")
    rec = recommended
    kvs = [
        ("Recommended (on financial merit)", rec["name"]),
        ("Stores", rec["n"]),
        ("Total capex", f"{rec['capex']:,.0f}"),
        ("NPV @ hurdle", f"{rec['npv']:,.0f}"),
        ("ROI", f"{rec['roi']:.1%}" if not np.isnan(rec['roi']) else "n/a"),
        ("IRR", f"{rec['irr']:.1%}" if not np.isnan(rec['irr']) else "undefined"),
        ("Payback", f"{rec['payback']:.1f} yrs" if not np.isnan(rec['payback']) else "beyond horizon"),
        ("Profitability index", f"{rec['pi']:.2f}" if not np.isnan(rec['pi']) else "n/a"),
        ("Meets all policy thresholds", "Yes" if rec["all_pass"] else "No - see caveats"),
    ]
    r = 5
    for k, v in kvs:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=v); r += 1
    ws.column_dimensions["A"].width = 32; ws.column_dimensions["B"].width = 28
    r += 1
    narrative(ws, r,
        "[ANALYST: explain WHY this scenario leads - it clears each policy threshold and offers the best "
        "risk-adjusted capital efficiency vs the alternatives. State the trade-off accepted. Frame explicitly as "
        "a recommendation for the CFO Advisor to decide on; do NOT make the final call or write the board paper.]")

    # 9 Key Assumptions
    ws = wb.create_sheet("Key Assumptions")
    ws_title(ws, "Key Assumptions", "Every assumption and its source.")
    a_rows = [
        ["Hurdle / discount rate", f"{args.hurdle_rate:.1%}", "Capital Investment Policy"],
        ["Minimum ROI", f"{args.min_roi:.1%}", "Capital Investment Policy"],
        ["Maximum payback", f"{args.max_payback_years:.1f} yrs", "Capital Investment Policy"],
        ["Capital ceiling", f"{args.capital_ceiling:,.0f}" if args.capital_ceiling else "not set", "Capital Investment Policy"],
        ["Forecast horizon", f"{args.years} yrs", "Policy / analyst"],
        ["Ramp curve", str(ramp), "Methodology default (override if scenario specifies)"],
        ["Market growth", f"{market_growth:.1%}", "Market_Assumptions / override"],
        ["Inflation", f"{asm.get('inflation'):.1%}" if not np.isnan(asm.get('inflation', np.nan)) else "n/a", "Market_Assumptions"],
        ["Labor inflation", f"{asm.get('labor_inflation'):.1%}" if not np.isnan(asm.get('labor_inflation', np.nan)) else "n/a", "Market_Assumptions"],
        ["Base revenue/store", f"{base_rev:,.0f}" if not np.isnan(base_rev) else "from scenarios", "Store_Analysis_Report"],
        ["Base EBITDA margin", f"{base_margin:.1%}" if not np.isnan(base_margin) else "from scenarios", "Store_Analysis_Report"],
    ]
    write_table(ws, 5, ["Assumption", "Value", "Source"], a_rows)
    ws.cell(row=4, column=1, value="[ANALYST: add per-scenario revenue/store, margin, capex/store actually used, and any analyst defaults.]").font = NOTE

    # 10 Handoff Notes
    ws = wb.create_sheet("Handoff Notes")
    ws_title(ws, "Handoff Notes - for the CFO Advisor", "Makes this case decision-ready for worker 3.")
    lines = [
        ("Verdict", f"Analysis indicates expansion is {'financially attractive' if recommended['all_pass'] else 'marginal - see caveats'}; "
                     f"'{recommended['name']}' is most attractive on financial merit."),
        ("Thresholds applied", f"Hurdle {args.hurdle_rate:.1%}, min ROI {args.min_roi:.1%}, max payback {args.max_payback_years:.1f} yrs"
                               + (f", ceiling {args.capital_ceiling:,.0f}" if args.capital_ceiling else "") + "."),
        ("Key sensitivities", "[ANALYST: the driver(s) the case is most sensitive to, from Scenario Analysis.]"),
        ("Leading risks", "[ANALYST: top risks from Risk Assessment.]"),
        ("Data-quality caveats", " | ".join(warnings) if warnings else "None material."),
        ("Open questions for the decision", "[ANALYST: strategic fit, financing/capital availability, board appetite, phasing of openings.]"),
        ("Out of scope (by design)", "Board recommendations and the final investment decision belong to the CFO Advisor - not this worker."),
    ]
    r = 5
    for k, v in lines:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=v); c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 2
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 95

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}  ({len(wb.sheetnames)} sheets: {wb.sheetnames})")
    print(f"Recommended on merit: {recommended['name']} (NPV {recommended['npv']:,.0f}, "
          f"payback {clean(recommended['payback'])}, PI {clean(recommended['pi'])}).")
    print("Next: author the [ANALYST: ...] narrative cells, then verify before handoff to the CFO Advisor.")


if __name__ == "__main__":
    main()
