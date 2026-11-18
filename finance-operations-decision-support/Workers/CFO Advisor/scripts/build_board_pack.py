#!/usr/bin/env python3
"""
build_board_pack.py
Crave Kitchen Finance Team - CFO Advisor (worker 3, final).

Reads the Finance Analyst's Store_Analysis_Report.xlsx and the FP&A Strategy
Partner's Expansion_Business_Case.xlsx, extracts the headline evidence (especially
the 10-store scenario), and assembles a board-ready Board_Recommendation_Pack.pptx
with all eight required sections, pre-filled with the prior workers' figures and
[CFO: ...] prompts where executive judgment and wording go.

This worker OWNS THE DECISION. The script surfaces whether the 10-store scenario
passes policy as a starting signal, but the CFO sets the actual recommendation and
final verdict by replacing the [CFO: ...] prompts.

The deck answers one question: Should Crave Kitchen Group open 10 new locations
next year?

Usage:
  python3 build_board_pack.py \
      --analysis      "Worker Outputs/Store_Analysis_Report.xlsx" \
      --business-case "Worker Outputs/Expansion_Business_Case.xlsx" \
      --output        "Final Deliverables/Board_Recommendation_Pack.pptx" \
      --decision-stores 10

Requires: pandas, openpyxl, python-pptx.
  pip install pandas openpyxl python-pptx --break-system-packages
"""
import argparse
import sys
import re
from datetime import date
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    sys.exit("Missing pandas. Install: pip install pandas openpyxl python-pptx --break-system-packages")

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
except ImportError:
    sys.exit("Missing python-pptx. Install: pip install python-pptx --break-system-packages")


# ---------------- palette ----------------
NAVY = RGBColor(0x1F, 0x38, 0x64)
SLATE = RGBColor(0x3B, 0x3B, 0x3B)
GREY = RGBColor(0x80, 0x80, 0x80)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xEE, 0xF1, 0xF7)


# ---------------- xlsx helpers ----------------

def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def num(v):
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        return float(str(v).replace(",", "").replace("$", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return np.nan


def find_col(cols, keys):
    nc = {c: _norm(c) for c in cols}
    for k in keys:
        for c, n in nc.items():
            if k == n or re.search(rf"\b{re.escape(k)}\b", n):
                return c
    for k in keys:
        for c, n in nc.items():
            if k in n:
                return c
    return None


def fmt_money(v):
    return "n/a" if (v is None or (isinstance(v, float) and np.isnan(v))) else f"${v:,.0f}"


def fmt_pct(v):
    return "n/a" if (v is None or (isinstance(v, float) and np.isnan(v))) else f"{v*100:.1f}%"


def fmt_yrs(v):
    return "n/a" if (v is None or (isinstance(v, float) and np.isnan(v))) else f"{v:.1f} yrs"


def extract_business_case(path, decision_stores):
    """Pull the decision-scenario metrics from Expansion_Business_Case.xlsx."""
    out = {"capex": np.nan, "npv": np.nan, "roi": np.nan, "irr": np.nan,
           "payback": np.nan, "pi": np.nan, "passes": None, "scenario_label": f"{decision_stores}-store",
           "all_rows": [], "recommended": None}
    if not path or not Path(path).exists():
        return out
    xl = pd.ExcelFile(path)
    sheet = None
    for cand in ["Executive Summary", "Scenario Analysis"]:
        if cand in xl.sheet_names:
            sheet = cand
            break
    sheet = sheet or xl.sheet_names[0]
    # find header row containing 'Scenario'
    raw = xl.parse(sheet, header=None)
    hdr_row = None
    for i in range(min(12, len(raw))):
        rowvals = [_norm(x) for x in raw.iloc[i].tolist()]
        if any("scenario" == v or v.startswith("scenario") for v in rowvals):
            hdr_row = i
            break
    if hdr_row is None:
        return out
    df = xl.parse(sheet, header=hdr_row)
    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)
    c = {
        "scenario": find_col(cols, ["scenario", "case", "option"]),
        "stores": find_col(cols, ["stores", "store count", "units"]),
        "capex": find_col(cols, ["total capex", "capex", "capital"]),
        "npv": find_col(cols, ["npv"]),
        "roi": find_col(cols, ["roi"]),
        "irr": find_col(cols, ["irr"]),
        "payback": find_col(cols, ["payback"]),
        "pi": find_col(cols, ["pi", "profitability index"]),
        "meets": find_col(cols, ["meets policy", "meets", "pass"]),
    }

    def val(row, key, pct=False):
        if not c[key]:
            return np.nan
        v = num(row[c[key]])
        if pct and not np.isnan(v) and abs(v) > 1.5:
            v = v / 100.0
        return v

    rows = []
    target = None
    for _, row in df.iterrows():
        label = str(row[c["scenario"]]) if c["scenario"] else ""
        if not label or _norm(label) in ("", "nan", "total"):
            continue
        st = val(row, "stores")
        rec = dict(
            label=label, stores=st,
            capex=val(row, "capex"), npv=val(row, "npv"),
            roi=val(row, "roi", pct=True), irr=val(row, "irr", pct=True),
            payback=val(row, "payback"), pi=val(row, "pi"),
            meets=str(row[c["meets"]]).strip() if c["meets"] else "",
        )
        rows.append(rec)
        # match the decision scenario by store count or label
        if (not np.isnan(rec["stores"]) and int(rec["stores"]) == decision_stores) or \
           (str(decision_stores) in _norm(label)):
            target = rec
    out["all_rows"] = rows
    if target:
        out.update({k: target[k] for k in ["capex", "npv", "roi", "irr", "payback", "pi"]})
        out["scenario_label"] = target["label"]
        out["passes"] = ("pass" in _norm(target.get("meets", "")))
    return out


def extract_analysis(path):
    """Pull a few base-health KPIs from Store_Analysis_Report.xlsx Executive Summary."""
    out = {"total_revenue": np.nan, "blended_margin": np.nan, "labor_pct": np.nan,
           "food_pct": np.nan, "prime_pct": np.nan, "stores": np.nan, "regions": np.nan}
    if not path or not Path(path).exists():
        return out
    xl = pd.ExcelFile(path)
    sheet = "Executive Summary" if "Executive Summary" in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)
    keymap = {
        "total_revenue": ["total revenue"],
        "blended_margin": ["blended ebitda margin", "ebitda margin"],
        "labor_pct": ["labor", "labour"],
        "food_pct": ["food"],
        "prime_pct": ["prime"],
        "stores": ["store count", "stores"],
        "regions": ["region count", "regions"],
    }
    for _, row in raw.iterrows():
        cells = row.tolist()
        for i, cell in enumerate(cells):
            n = _norm(cell)
            for key, kws in keymap.items():
                if not np.isnan(out[key]):
                    continue
                if any(kw in n for kw in kws):
                    for j in range(i + 1, len(cells)):
                        v = num(cells[j])
                        if not np.isnan(v):
                            if key.endswith("pct") or key == "blended_margin":
                                out[key] = v / 100.0 if abs(v) > 1.5 else v
                            else:
                                out[key] = v
                            break
    return out


# ---------------- slide helpers ----------------

def add_title_slide(prs, decision_stores):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _band(s, prs, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _text(s, Inches(0.7), Inches(2.2), prs.slide_width - Inches(1.4), Inches(1.4),
          "Board Recommendation", 40, WHITE, bold=True)
    _text(s, Inches(0.7), Inches(3.5), prs.slide_width - Inches(1.4), Inches(0.8),
          "Crave Kitchen Group — Expansion Decision", 22, WHITE)
    _text(s, Inches(0.7), Inches(4.3), prs.slide_width - Inches(1.4), Inches(0.7),
          f"Should we open {decision_stores} new locations next year?", 18, RGBColor(0xC9, 0xD4, 0xEA), italic=True)
    _text(s, Inches(0.7), prs.slide_height - Inches(1.0), prs.slide_width - Inches(1.4), Inches(0.5),
          f"Prepared by the CFO Advisor  ·  {date.today():%B %d, %Y}  ·  Confidential — Board Use Only",
          12, RGBColor(0xC9, 0xD4, 0xEA))
    return s


def add_content_slide(prs, header, kicker=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _band(s, prs, 0, 0, prs.slide_width, Inches(1.15), NAVY)
    _text(s, Inches(0.6), Inches(0.18), prs.slide_width - Inches(1.2), Inches(0.8),
          header, 26, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if kicker:
        _text(s, Inches(0.6), Inches(1.25), prs.slide_width - Inches(1.2), Inches(0.45),
              kicker, 14, GREY, italic=True)
    return s


def _band(slide, prs, l, t, w, h, color):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _text(slide, l, t, w, h, text, size, color, bold=False, italic=False,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic; f.color.rgb = color
    return tb


def bullets(slide, l, t, w, h, items, size=15, color=SLATE):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(7)
        if isinstance(item, tuple):
            txt = item[0]
            lvl = int(item[1]) if len(item) > 1 and item[1] != "" else 0
            c = item[2] if len(item) > 2 and item[2] else color
            b = bool(item[3]) if len(item) > 3 else False
        else:
            txt, lvl, c, b = item, 0, color, False
        p.level = lvl
        r = p.add_run(); r.text = ("•  " if lvl == 0 else "–  ") + txt
        r.font.size = Pt(size); r.font.color.rgb = c; r.font.bold = b
    return tb


def kpi_strip(slide, prs, top, pairs):
    n = len(pairs)
    gap = Inches(0.2)
    total_w = prs.slide_width - Inches(1.2)
    cw = (total_w - gap * (n - 1)) / n
    left = Inches(0.6)
    for i, (label, value, good) in enumerate(pairs):
        x = left + i * (cw + gap)
        card = _band(slide, prs, x, top, cw, Inches(1.15), LIGHT)
        color = NAVY if good is None else (GREEN if good else RED)
        _text(slide, x, top + Inches(0.12), cw, Inches(0.5), value, 18, color, bold=True, align=PP_ALIGN.CENTER)
        _text(slide, x, top + Inches(0.66), cw, Inches(0.4), label, 10.5, GREY, align=PP_ALIGN.CENTER)


def cfo_prompt(slide, l, t, w, text):
    tb = slide.shapes.add_textbox(l, t, w, Inches(1.0))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(12); r.font.italic = True; r.font.color.rgb = RGBColor(0xB0, 0x60, 0x00)
    return tb


# ---------------- main ----------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--analysis", required=False)
    ap.add_argument("--business-case", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--decision-stores", type=int, default=10)
    args = ap.parse_args()

    D = args.decision_stores
    bc = extract_business_case(args.business_case, D)
    an = extract_analysis(args.analysis)

    print(f"Decision scenario: {bc['scenario_label']}")
    print(f"  capex={fmt_money(bc['capex'])}  NPV={fmt_money(bc['npv'])}  ROI={fmt_pct(bc['roi'])}  "
          f"IRR={fmt_pct(bc['irr'])}  payback={fmt_yrs(bc['payback'])}  PI={bc['pi']}")
    print(f"  passes policy: {bc['passes']}")
    if np.isnan(bc['npv']):
        print("  ! Could not read the decision-scenario metrics. Check the business case layout; "
              "slides will carry [CFO: ...] prompts to fill manually.")
    print(f"Base health: revenue={fmt_money(an['total_revenue'])}, blended margin={fmt_pct(an['blended_margin'])}, "
          f"stores={an['stores']}, regions={an['regions']}")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    signal = ("the 10-store case PASSES the board's investment thresholds" if bc["passes"] is True
              else "the 10-store case FAILS one or more board thresholds" if bc["passes"] is False
              else "policy pass/fail could not be auto-read — confirm against the business case")

    # 0 Title
    add_title_slide(prs, D)

    # 1 Executive Summary
    s = add_content_slide(prs, "Executive Summary", "The question, and the CFO's answer")
    cfo_prompt(s, Inches(0.6), Inches(1.75), Inches(12.1),
               f"[CFO: state your one-line recommendation here — proceed / proceed with conditions / "
               f"do not proceed. Data signal: {signal}.]")
    bullets(s, Inches(0.6), Inches(2.7), Inches(12.1), Inches(2.0), [
        ("[CFO: base health — is the current portfolio strong enough to expand from? cite worker 1.]",),
        (f"10-store economics: NPV {fmt_money(bc['npv'])}, ROI {fmt_pct(bc['roi'])}, payback {fmt_yrs(bc['payback'])}.",),
        ("[CFO: the single biggest risk, in one line.]",),
        ("[CFO: strategic fit, in one line — does this advance the strategy?]",),
    ])
    kpi_strip(s, prs, Inches(5.9), [
        ("Total capex", fmt_money(bc["capex"]), None),
        ("NPV @ hurdle", fmt_money(bc["npv"]), (bc["npv"] > 0) if not np.isnan(bc["npv"]) else None),
        ("ROI", fmt_pct(bc["roi"]), None),
        ("Payback", fmt_yrs(bc["payback"]), None),
        ("Meets policy", ("PASS" if bc["passes"] else "FAIL") if bc["passes"] is not None else "review",
         bc["passes"]),
    ])

    # 2 Strategic Assessment
    s = add_content_slide(prs, "Strategic Assessment", "Does opening 10 advance the strategy?")
    bullets(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(3.0), [
        ("[CFO: how 10-store expansion maps to the Executive Briefing strategy and leadership priorities.]",),
        ("[CFO: strengths it builds on — proven regions, unit economics, brand (cite worker 1).]",),
        ("[CFO: where it stretches the company beyond proven territory, if at all.]",),
    ])
    cfo_prompt(s, Inches(0.6), Inches(5.4), Inches(12.1),
               "[CFO: verdict on strategic fit — strong / qualified / weak, and why.]")

    # 3 Financial Assessment
    s = add_content_slide(prs, "Financial Assessment", "The 10-store investment case")
    _scenario_table(prs, s, bc, D)
    bullets(s, Inches(0.6), Inches(4.8), Inches(12.1), Inches(1.4), [
        (f"Base portfolio: revenue {fmt_money(an['total_revenue'])}, blended EBITDA margin {fmt_pct(an['blended_margin'])}, "
         f"{_intt(an['stores'])} stores across {_intt(an['regions'])} regions (worker 1).",),
        (f"Data signal: {signal}.",),
    ])
    cfo_prompt(s, Inches(0.6), Inches(6.1), Inches(12.1),
               "[CFO: is the financial case attractive, marginal, or weak — and is the base strong enough to expand from?]")

    # 4 Major Risks
    s = add_content_slide(prs, "Major Risks", "What could go wrong, and how we manage it")
    bullets(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.6), [
        ("[CFO: execution risk of opening many units at once] — [mitigation].",),
        ("[CFO: demand / ramp slower than forecast] — [mitigation].",),
        ("[CFO: cost & labor inflation pressure on margins] — [mitigation].",),
        ("[CFO: concentration / financing & liquidity risk] — [mitigation].",),
        ("[CFO: name the variable the case is MOST sensitive to, from worker 2's sensitivity analysis.]",),
    ])

    # 5 Major Opportunities
    s = add_content_slide(prs, "Major Opportunities", "The upside if we proceed")
    bullets(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.6), [
        ("[CFO: market capture / growth opportunity — cite worker 2 market assumptions.]",),
        ("[CFO: replicate the proven strong regions — cite worker 1.]",),
        ("[CFO: margin headroom / unit-economics upside.]",),
        ("[CFO: brand momentum and strategic positioning.]",),
    ])

    # 6 Recommendation
    s = add_content_slide(prs, "Recommendation", "The CFO's call")
    cfo_prompt(s, Inches(0.6), Inches(1.85), Inches(12.1),
               "[CFO: state the recommendation plainly — proceed / proceed with conditions / do not proceed this year.]")
    bullets(s, Inches(0.6), Inches(2.9), Inches(12.1), Inches(2.6), [
        ("[CFO: if conditional — list the gating conditions: phasing, financing secured, target regions, KPI gates.]",),
        ("[CFO: 2–3 sentences of reasoning — risk-adjusted return + strategic fit + capital discipline.]",),
    ])

    # 7 Supporting Rationale
    s = add_content_slide(prs, "Supporting Rationale", "The evidence chain behind the call")
    bullets(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.6), [
        (f"Base health (Finance Analyst): blended margin {fmt_pct(an['blended_margin'])}, "
         f"{_intt(an['stores'])} stores — [CFO: strong enough to expand from? one line].",),
        (f"Expansion economics (FP&A Strategy Partner): 10-store NPV {fmt_money(bc['npv'])}, "
         f"ROI {fmt_pct(bc['roi'])}, payback {fmt_yrs(bc['payback'])}, policy {('PASS' if bc['passes'] else 'FAIL') if bc['passes'] is not None else 'review'}.",),
        ("Board Investment Framework: [CFO: how the decision maps to the board's approval criteria].",),
        ("[CFO: why this recommendation beats the alternatives — the other scenarios or a delay.]",),
    ])

    # 8 Final Decision
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _band(s, prs, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _text(s, Inches(0.8), Inches(0.9), prs.slide_width - Inches(1.6), Inches(0.8),
          "Final Decision", 30, WHITE, bold=True)
    _text(s, Inches(0.8), Inches(2.1), prs.slide_width - Inches(1.6), Inches(1.0),
          f"Should Crave Kitchen Group open {D} new locations next year?", 22,
          RGBColor(0xC9, 0xD4, 0xEA), italic=True)
    box = _band(s, prs, Inches(0.8), Inches(3.2), prs.slide_width - Inches(1.6), Inches(1.5), WHITE)
    _text(s, Inches(1.0), Inches(3.45), prs.slide_width - Inches(2.0), Inches(1.0),
          "[CFO: YES  /  YES, WITH CONDITIONS  /  NO] — Crave Kitchen Group [should / should not] "
          f"open {D} new locations next year.", 22, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _text(s, Inches(0.8), Inches(5.1), prs.slide_width - Inches(1.6), Inches(1.2),
          "[CFO: one line — the single most important reason. If conditional, state the gating conditions "
          "and any approvals/next steps required.]", 15, RGBColor(0xD8, 0xE0, 0xF0), italic=True)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print(f"\nWrote {out}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
    print("Next: replace every [CFO: ...] prompt with director-ready prose, set the recommendation "
          "consistently on the Executive Summary and Final Decision slides, then verify before the board sees it.")


def _intt(v):
    return "n/a" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(int(v))


def _scenario_table(prs, slide, bc, D):
    """Comparison table of all scenarios with the decision scenario highlighted."""
    from pptx.enum.text import PP_ALIGN as A
    rows_data = bc["all_rows"] or [{
        "label": f"{D}-store", "stores": D, "capex": bc["capex"], "npv": bc["npv"],
        "roi": bc["roi"], "irr": bc["irr"], "payback": bc["payback"], "pi": bc["pi"], "meets": ""}]
    headers = ["Scenario", "Capex", "NPV", "ROI", "IRR", "Payback", "PI", "Policy"]
    nrows = len(rows_data) + 1
    tbl_shape = slide.shapes.add_table(nrows, len(headers), Inches(0.6), Inches(1.9),
                                       Inches(12.1), Inches(0.45) * nrows)
    table = tbl_shape.table
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]; p.alignment = A.CENTER
        p.runs[0].font.size = Pt(12); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = WHITE
    for i, r in enumerate(rows_data, 1):
        is_decision = (not np.isnan(num(r.get("stores", np.nan))) and int(num(r["stores"])) == D) or \
                      (str(D) in _norm(str(r.get("label", ""))))
        vals = [str(r.get("label", "")), fmt_money(r.get("capex")), fmt_money(r.get("npv")),
                fmt_pct(r.get("roi")), fmt_pct(r.get("irr")), fmt_yrs(r.get("payback")),
                ("n/a" if np.isnan(num(r.get("pi", np.nan))) else f"{num(r['pi']):.2f}"),
                (r.get("meets", "") or "")]
        for j, v in enumerate(vals):
            cell = table.cell(i, j)
            cell.text = v
            cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT if is_decision else WHITE
            p = cell.text_frame.paragraphs[0]; p.alignment = A.CENTER
            p.runs[0].font.size = Pt(11)
            p.runs[0].font.bold = is_decision
            p.runs[0].font.color.rgb = NAVY if is_decision else SLATE


if __name__ == "__main__":
    main()
