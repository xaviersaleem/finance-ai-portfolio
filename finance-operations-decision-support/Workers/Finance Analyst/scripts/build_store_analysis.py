#!/usr/bin/env python3
"""
build_store_analysis.py
Crave Kitchen Finance Team - Finance Analyst (worker 1).

Reads Current_Store_Performance.xlsx (Airtable export: store revenue, EBITDA,
labor cost, food cost, regional performance) and writes a formatted, multi-sheet
Store_Analysis_Report.xlsx containing the required analysis sections.

The script is schema-flexible: it fuzzy-matches columns and prints the mapping it
chose plus data-quality warnings. Override any guess with the --*-col flags.

It computes the data tables and KPI rows. The analyst then authors the narrative
cells. This worker DESCRIBES current performance only - no forecasting, no
investment analysis, no expansion recommendation.

Usage:
  python3 build_store_analysis.py \
      --input "Source Data/Current_Store_Performance.xlsx" \
      --output "Worker Outputs/Store_Analysis_Report.xlsx"

Optional column overrides (1-based header names as they appear in the sheet):
  --store-col --region-col --revenue-col --ebitda-col --labor-col --food-col --period-col

Benchmark overrides (from the Executive Briefing, if stated):
  --target-ebitda-margin 0.18  --target-labor-pct 0.30  --target-food-pct 0.32

Requires: pandas, openpyxl.  pip install pandas openpyxl --break-system-packages
"""
import argparse
import sys
import re
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    sys.exit("Missing pandas/numpy. Install: pip install pandas openpyxl --break-system-packages")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing openpyxl. Install: pip install openpyxl --break-system-packages")


# ---------- column detection ----------

CANDIDATES = {
    "store":   ["store", "location", "site", "restaurant", "unit", "name", "branch", "outlet"],
    "region":  ["region", "area", "market", "district", "zone", "territory", "state"],
    "revenue": ["revenue", "sales", "net sales", "total revenue", "turnover", "gross sales"],
    "ebitda":  ["ebitda", "operating profit", "operating income", "ebit"],
    "labor":   ["labor", "labour", "payroll", "wages", "staff cost", "labor cost"],
    "food":    ["food", "cogs", "cost of goods", "food cost", "ingredient"],
    "period":  ["period", "month", "date", "quarter", "year", "week"],
}


def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def detect_column(columns, keys, override=None):
    if override:
        for c in columns:
            if _norm(c) == _norm(override):
                return c
        # allow partial match on override
        for c in columns:
            if _norm(override) in _norm(c):
                return c
        print(f"  ! override '{override}' not found; falling back to auto-detect")
    ncols = {c: _norm(c) for c in columns}
    # exact-ish: candidate is a whole word in the column
    for key in keys:
        for c, n in ncols.items():
            if key == n or re.search(rf"\b{re.escape(key)}\b", n):
                return c
    # loose: candidate substring
    for key in keys:
        for c, n in ncols.items():
            if key in n:
                return c
    return None


def is_percentish(series):
    """Heuristic: are these values percentages-as-decimals already (<=1) vs whole pct (e.g., 28.5)?"""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0:
        return None
    return s.median() <= 1.5  # decimals like 0.28


def to_ratio(series):
    """Normalize a percent column to a decimal ratio."""
    s = pd.to_numeric(series, errors="coerce")
    if is_percentish(s):
        return s
    return s / 100.0


# ---------- styling ----------

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


def write_table(ws, start_row, headers, rows, pct_cols=None, money_cols=None, totals=None):
    pct_cols = pct_cols or set()
    money_cols = money_cols or set()
    r = start_row
    for j, head in enumerate(headers, 1):
        c = ws.cell(row=r, column=j, value=head)
        c.font = H
        c.fill = HFILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER
    for row in rows:
        r += 1
        for j, val in enumerate(row, 1):
            c = ws.cell(row=r, column=j, value=val)
            c.border = BORDER
            if j - 1 in pct_cols and isinstance(val, (int, float)):
                c.number_format = "0.0%"
            elif j - 1 in money_cols and isinstance(val, (int, float)):
                c.number_format = "#,##0"
    if totals:
        r += 1
        for j, val in enumerate(totals, 1):
            c = ws.cell(row=r, column=j, value=val)
            c.font = TOTAL
            c.border = BORDER
            if j - 1 in pct_cols and isinstance(val, (int, float)):
                c.number_format = "0.0%"
            elif j - 1 in money_cols and isinstance(val, (int, float)):
                c.number_format = "#,##0"
    # width
    for j, head in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(30, len(str(head)) + 4))
    return r + 1


def narrative_prompt(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = NOTE
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 3, end_column=8)
    return row + 5


# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--sheet", default=None, help="worksheet name in the input workbook")
    for k in ["store", "region", "revenue", "ebitda", "labor", "food", "period"]:
        ap.add_argument(f"--{k}-col", default=None)
    ap.add_argument("--target-ebitda-margin", type=float, default=None)
    ap.add_argument("--target-labor-pct", type=float, default=None)
    ap.add_argument("--target-food-pct", type=float, default=None)
    args = ap.parse_args()

    inp = Path(args.input)
    if not inp.exists():
        sys.exit(f"Input not found: {inp}")

    xl = pd.ExcelFile(inp)
    sheet = args.sheet or xl.sheet_names[0]
    df = xl.parse(sheet)
    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)
    print(f"Loaded '{inp.name}' sheet '{sheet}': {len(df)} rows, {len(cols)} cols")
    print(f"Columns: {cols}")

    m = {}
    for key, keys in CANDIDATES.items():
        m[key] = detect_column(cols, keys, getattr(args, f"{key}_col"))
    print("\nColumn mapping chosen:")
    for k, v in m.items():
        print(f"  {k:8s} -> {v}")
    print("Review this mapping. Re-run with --<field>-col to override any wrong guess.\n")

    warnings = []
    if not m["revenue"]:
        sys.exit("Could not find a revenue column. Pass --revenue-col explicitly.")
    if not m["store"]:
        warnings.append("No store column detected; using row index as store id.")
        df["_store"] = [f"Store {i+1}" for i in range(len(df))]
        m["store"] = "_store"
    if not m["region"]:
        warnings.append("No region column detected; regional analysis will be limited.")
        df["_region"] = "Unspecified"
        m["region"] = "_region"

    # numeric coercion
    df["revenue"] = pd.to_numeric(df[m["revenue"]], errors="coerce")
    df["ebitda"] = pd.to_numeric(df[m["ebitda"]], errors="coerce") if m["ebitda"] else np.nan

    # labor / food may be dollars or percents
    def cost_dollars_pct(colkey):
        if not m[colkey]:
            return np.nan, np.nan
        raw = pd.to_numeric(df[m[colkey]], errors="coerce")
        looks_pct = "%" in str(m[colkey]) or "pct" in _norm(m[colkey]) or "percent" in _norm(m[colkey])
        med = raw.dropna().median() if raw.notna().any() else np.nan
        if looks_pct or (pd.notna(med) and med <= 1.5):  # ratio
            pct = raw if (pd.notna(med) and med <= 1.5) else raw / 100.0
            dollars = pct * df["revenue"]
        elif pd.notna(med) and med < 100 and df["revenue"].median() > 1000:
            # whole-number percent like 28.5
            pct = raw / 100.0
            dollars = pct * df["revenue"]
        else:
            dollars = raw
            pct = raw / df["revenue"]
        return dollars, pct

    df["labor"], df["labor_pct"] = cost_dollars_pct("labor")
    df["food"], df["food_pct"] = cost_dollars_pct("food")

    df["ebitda_margin"] = df["ebitda"] / df["revenue"]
    df["prime_pct"] = df["labor_pct"] + df["food_pct"]
    df["region"] = df[m["region"]].astype(str)
    df["store"] = df[m["store"]].astype(str)

    # data quality
    if df["revenue"].isna().any():
        warnings.append(f"{int(df['revenue'].isna().sum())} rows with non-numeric/missing revenue.")
    if (df["revenue"] <= 0).any():
        warnings.append(f"{int((df['revenue'] <= 0).sum())} rows with zero/negative revenue.")
    if df["store"].duplicated().any():
        warnings.append("Duplicate store names present - verify the export.")
    for label, s in [("food %", df["food_pct"]), ("labor %", df["labor_pct"])]:
        bad = ((s > 0.6) | (s < 0.05)).sum()
        if bad:
            warnings.append(f"{int(bad)} stores with implausible {label} (check units).")

    # benchmarks: briefing target else portfolio median
    bm = {
        "ebitda_margin": args.target_ebitda_margin if args.target_ebitda_margin is not None else df["ebitda_margin"].median(),
        "labor_pct": args.target_labor_pct if args.target_labor_pct is not None else df["labor_pct"].median(),
        "food_pct": args.target_food_pct if args.target_food_pct is not None else df["food_pct"].median(),
        "revenue": df["revenue"].median(),
    }
    bm_source = {
        "ebitda_margin": "briefing target" if args.target_ebitda_margin is not None else "portfolio median",
        "labor_pct": "briefing target" if args.target_labor_pct is not None else "portfolio median",
        "food_pct": "briefing target" if args.target_food_pct is not None else "portfolio median",
    }

    total_rev = df["revenue"].sum()
    df["rev_share"] = df["revenue"] / total_rev
    blended_margin = df["ebitda"].sum() / total_rev if df["ebitda"].notna().any() else np.nan
    port_labor = df["labor"].sum() / total_rev if df["labor"].notna().any() else np.nan
    port_food = df["food"].sum() / total_rev if df["food"].notna().any() else np.nan

    # composite score (z-scores; higher = better)
    def z(s):
        s = s.astype(float)
        sd = s.std(ddof=0)
        return (s - s.mean()) / sd if sd and not np.isnan(sd) else s * 0
    df["score"] = (
        z(df["ebitda_margin"]).fillna(0)
        + z(df["revenue"]).fillna(0)
        - z(df["labor_pct"]).fillna(0)
        - z(df["food_pct"]).fillna(0)
    )
    ranked = df.sort_values("score", ascending=False).reset_index(drop=True)

    if warnings:
        print("DATA-QUALITY WARNINGS:")
        for w in warnings:
            print("  ! " + w)
        print()

    # ---------- build workbook ----------
    wb = Workbook()

    def fav_fill(cell, value, lower_is_better):
        if not isinstance(value, (int, float)) or np.isnan(value):
            return
        if (value > 0) != lower_is_better:
            cell.fill = GREEN
        else:
            cell.fill = RED

    # 1. Executive Summary
    ws = wb.active
    ws.title = "Executive Summary"
    ws_title(ws, "Store Analysis Report - Executive Summary",
             "Crave Kitchen Group | Finance Analyst (worker 1) | current-performance read for the FP&A Strategy Partner")
    kpis = [
        ["Total revenue", total_rev, "#,##0"],
        ["Blended EBITDA margin", blended_margin, "0.0%"],
        ["Portfolio labor %", port_labor, "0.0%"],
        ["Portfolio food %", port_food, "0.0%"],
        ["Portfolio prime cost %", (port_labor or 0) + (port_food or 0), "0.0%"],
        ["Store count", int(df["store"].nunique()), "0"],
        ["Region count", int(df["region"].nunique()), "0"],
    ]
    r = 5
    ws.cell(row=r, column=1, value="Key metrics").font = SUB
    r += 1
    for label, val, fmt in kpis:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=val)
        c.number_format = fmt
        r += 1
    r += 1
    ws.cell(row=r, column=1,
            value=f"Benchmarks used - EBITDA margin: {bm_source['ebitda_margin']} "
                  f"({bm['ebitda_margin']:.1%}); labor %: {bm_source['labor_pct']} ({bm['labor_pct']:.1%}); "
                  f"food %: {bm_source['food_pct']} ({bm['food_pct']:.1%}).").font = NOTE
    r += 2
    r = narrative_prompt(ws, r,
        "[ANALYST: write 4-7 headline findings - best/worst stores and regions, the biggest "
        "margin driver, any revenue/region concentration, and overall cost posture. One line each "
        "on risk posture and strength of the current base. Cite figures. Describe only - no "
        "forecasts or expansion recommendations.]")
    if warnings:
        ws.cell(row=r, column=1, value="Data-quality caveats:").font = SUB
        r += 1
        for w in warnings:
            ws.cell(row=r, column=1, value="- " + w).font = NOTE
            r += 1
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18

    # helper to add an analysis sheet
    def metric_sheet(name, purpose, value_col, pct_col, benchmark_key, lower_is_better,
                     value_label, pct_label, port_value):
        ws = wb.create_sheet(name)
        ws_title(ws, name, purpose)
        headers = ["Store", "Region", value_label, pct_label, f"Var vs benchmark ({pct_label})"]
        rows, fills = [], []
        for _, x in ranked.sort_values("revenue", ascending=False).iterrows():
            var = (x[pct_col] - bm[benchmark_key]) if pd.notna(x[pct_col]) else np.nan
            rows.append([x["store"], x["region"],
                         _num(x[value_col]), _num(x[pct_col]), _num(var)])
            fills.append(var)
        end = write_table(ws, 5, headers, rows,
                          pct_cols={3, 4}, money_cols={2},
                          totals=["TOTAL / PORTFOLIO", "",
                                  _num(df[value_col].sum()), _num(port_value), ""])
        # color variance col (index 4)
        for i, var in enumerate(fills):
            cell = ws.cell(row=6 + i, column=5)
            fav_fill(cell, var, lower_is_better)
        narrative_prompt(ws, end + 1,
            f"[ANALYST: interpret {name.lower()} - distribution, leaders/laggards, and the driver "
            f"behind the variance. Name stores and regions. Benchmark = {bm_source.get(benchmark_key, 'portfolio median')}.]")
        return ws

    metric_sheet("Revenue Analysis", "Revenue level, mix, and variance by store.",
                 "revenue", "rev_share", "revenue", False, "Revenue", "Revenue share %", df["rev_share"].sum())
    # revenue benchmark/variance differs; fix var col semantics handled generically above

    metric_sheet("EBITDA Analysis", "EBITDA and EBITDA margin by store.",
                 "ebitda", "ebitda_margin", "ebitda_margin", False, "EBITDA", "EBITDA margin %", blended_margin)
    metric_sheet("Labor Cost Analysis", "Labor cost and labor % of revenue by store.",
                 "labor", "labor_pct", "labor_pct", True, "Labor cost", "Labor cost %", port_labor)
    metric_sheet("Food Cost Analysis", "Food cost and food % of revenue by store.",
                 "food", "food_pct", "food_pct", True, "Food cost", "Food cost %", port_food)

    # 6. Regional Analysis
    ws = wb.create_sheet("Regional Analysis")
    ws_title(ws, "Regional Analysis", "Performance aggregated by region.")
    grp = df.groupby("region")
    reg_rows = []
    for region, g in grp:
        rev = g["revenue"].sum()
        reg_rows.append([
            region, int(g["store"].nunique()), _num(rev), _num(rev / total_rev),
            _num(g["ebitda"].sum() / rev if rev else np.nan),
            _num(g["labor"].sum() / rev if rev else np.nan),
            _num(g["food"].sum() / rev if rev else np.nan),
            _num((g["labor"].sum() + g["food"].sum()) / rev if rev else np.nan),
        ])
    reg_rows.sort(key=lambda x: (x[2] if isinstance(x[2], (int, float)) else 0), reverse=True)
    end = write_table(ws, 5,
        ["Region", "Stores", "Revenue", "Revenue share", "EBITDA margin", "Labor %", "Food %", "Prime cost %"],
        reg_rows, pct_cols={3, 4, 5, 6, 7}, money_cols={2},
        totals=["TOTAL", int(df["store"].nunique()), _num(total_rev), _num(1.0),
                _num(blended_margin), _num(port_labor), _num(port_food),
                _num((port_labor or 0) + (port_food or 0))])
    narrative_prompt(ws, end + 1,
        "[ANALYST: strongest vs weakest regions, within-region consistency, and any concentration. "
        "Note regions with multiple healthy stores (a proven model) vs single/weak regions. Descriptive only.]")

    # 7 & 8 Top / Underperforming
    def rank_sheet(name, purpose, frame, worst_first):
        ws = wb.create_sheet(name)
        ws_title(ws, name, purpose)
        rows = []
        for i, (_, x) in enumerate(frame.iterrows(), 1):
            rows.append([i, x["store"], x["region"], round(float(x["score"]), 2),
                         _num(x["ebitda_margin"]), _num(x["revenue"]),
                         _num(x["labor_pct"]), _num(x["food_pct"])])
        end = write_table(ws, 5,
            ["Rank", "Store", "Region", "Composite score", "EBITDA margin", "Revenue", "Labor %", "Food %"],
            rows, pct_cols={4, 6, 7}, money_cols={5})
        narrative_prompt(ws, end + 1,
            "[ANALYST: explain what drives each store's position - the specific lever "
            "(revenue scale, EBITDA margin, labor %, food %). Flag stores strong on one axis but dragged by another.]")
    n = max(3, min(10, len(ranked) // 3 or 3))
    rank_sheet("Top Performing Stores", "Highest composite performers (audit metrics shown).",
               ranked.head(n), False)
    rank_sheet("Underperforming Stores", "Lowest composite performers (audit metrics shown).",
               ranked.tail(n).iloc[::-1], True)

    # 9 Expansion Readiness (signals only)
    ws = wb.create_sheet("Expansion Readiness Assessment")
    ws_title(ws, "Expansion Readiness Assessment",
             "Strength/stability signals of the CURRENT base. Descriptive only - NOT a recommendation or forecast.")
    above = int((df["ebitda_margin"] > bm["ebitda_margin"]).sum())
    sig = [
        ["Profitability depth", f"{above} of {len(df)} stores above the EBITDA-margin benchmark ({bm['ebitda_margin']:.1%})"],
        ["Margin consistency", f"EBITDA margin spread: {df['ebitda_margin'].min():.1%} to {df['ebitda_margin'].max():.1%} (std {df['ebitda_margin'].std():.1%})"],
        ["Prime-cost consistency", f"Prime cost spread: {df['prime_pct'].min():.1%} to {df['prime_pct'].max():.1%}"],
        ["Cost control", f"Portfolio labor {port_labor:.1%}, food {port_food:.1%}, prime {(port_labor or 0)+(port_food or 0):.1%}"],
        ["Regional strength", f"{df['region'].nunique()} regions; revenue concentration in top region: {df.groupby('region')['revenue'].sum().max()/total_rev:.1%}"],
    ]
    r = 5
    for label, val in sig:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=2, value=val)
        r += 1
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 70
    r += 1
    narrative_prompt(ws, r,
        "[ANALYST: summarize base strength from the signals above and list OPEN QUESTIONS for the "
        "FP&A Strategy Partner. Do NOT state ready/not-ready and do NOT name where to expand - that is the next worker's job.]")

    # 10 & 11 Risks / Opportunities (templates)
    for name, purpose, c4 in [
        ("Key Risks", "Threats to current performance, evidence-based.", "Severity (High/Med/Low)"),
        ("Key Opportunities", "Upside visible in the current base - for the next worker to investigate.", "Area to investigate"),
    ]:
        ws = wb.create_sheet(name)
        ws_title(ws, name, purpose)
        write_table(ws, 5,
            [name[4:-1] if False else ("Risk" if "Risk" in name else "Opportunity"),
             "Evidence (stores/regions/figures)", "Affected area", c4],
            [["", "", "", ""] for _ in range(8)])
        ws.cell(row=4, column=1,
                value="[ANALYST: fill rows from analysis_methodology.md section 8. Frame opportunities for the next worker - no projections.]").font = NOTE

    # 12 Handoff Notes
    ws = wb.create_sheet("Handoff Notes")
    ws_title(ws, "Handoff Notes - for the FP&A Strategy Partner",
             "Makes this report consumable by worker 2.")
    lines = [
        ("Scope completed", "Current store & business performance: revenue, EBITDA, labor, food, regional, variance, trends, rankings, base-strength signals, risks, opportunities."),
        ("Benchmarks / thresholds used", f"EBITDA margin {bm['ebitda_margin']:.1%} ({bm_source['ebitda_margin']}); labor {bm['labor_pct']:.1%} ({bm_source['labor_pct']}); food {bm['food_pct']:.1%} ({bm_source['food_pct']}); revenue median {bm['revenue']:,.0f}."),
        ("Data-quality caveats", " | ".join(warnings) if warnings else "None material."),
        ("Open questions for expansion eval", "[ANALYST: list what worker 2 must answer - e.g., demand drivers behind strong regions, fixability of high-cost stores, capacity headroom.]"),
        ("Out of scope (by design)", "Forecasting, investment analysis (NPV/IRR/payback), and expansion recommendations belong to the FP&A Strategy Partner and CFO Advisor - not this worker."),
    ]
    r = 5
    for label, val in lines:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=val)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 2
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 90

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}  ({len(wb.sheetnames)} sheets: {wb.sheetnames})")
    print("Next: open the report and author the [ANALYST: ...] narrative cells, then verify before handoff.")


def _num(v):
    """Return a clean float or empty string for NaN."""
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        return float(v)
    except (TypeError, ValueError):
        return ""


if __name__ == "__main__":
    main()
