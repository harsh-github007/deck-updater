#!/usr/bin/env python3
"""Build the stock evaluation model (v2).

v1 had two scoring modules that were mathematically incapable of returning
anything but 100:

  Buffett valuation = 50 + (B32 - EPSgrowth) * 200, clamped to 100
      B32 is the *earnings-multiple fair value in dollars* (e.g. 129.06), where
      the module's own description says "upside vs margin of safety". Any
      fair value above ~0.6 pinned the score at 100. A company 50% overvalued
      scored as well as one 50% undervalued.

  Growth quality = AVERAGE(RevCAGR, EPSCAGR, FCFCAGR) * 1000, clamped to 100
      10% average growth reaches 100, so every grower scored full marks and
      the module could not rank anything. Worse, a CAGR that failed to compute
      (negative base year) returned "" and AVERAGE silently skipped it, so a
      missing number scored the same as a perfect one.

Both are fixed here, every component score is exposed in its own row so the
arithmetic is auditable, and the model now carries a data-quality panel that
can force the decision to REVIEW rather than emit a confident BUY/SELL on
inputs it knows are broken.

    python3 build_model.py -o stock_model_v2.xlsx
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------- appearance
INK = "1B2733"
MUTED = "5C6B7A"
ACCENT = "1F6F5C"
HEAD_FILL = PatternFill("solid", fgColor="1F6F5C")
BAND_FILL = PatternFill("solid", fgColor="EDF2F0")
INPUT_FILL = PatternFill("solid", fgColor="FFF6E5")   # user types here
CALC_FILL = PatternFill("solid", fgColor="F5F7F9")    # formula driven
TITLE_F = Font(name="Calibri", size=14, bold=True, color=INK)
SUB_F = Font(name="Calibri", size=10, color=MUTED, italic=True)
SEC_F = Font(name="Calibri", size=11, bold=True, color=ACCENT)
HEAD_F = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY_F = Font(name="Calibri", size=10, color=INK)
THIN = Side(style="thin", color="D8DEE3")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PCT = "0.0%"
PCT2 = "0.00%"
USD = "#,##0.00"
MM = "#,##0"
MULT = "0.00x"

IN = "'01_Input_Template'"
DV = "'02_Derived_Values'"

# ------------------------------------------------------------- sheet 1 layout
# company profile
R_TICKER, R_NAME, R_CCY, R_SECTOR = 6, 7, 8, 9
R_PRICE, R_SHARES, R_MOAT, R_NOTES = 10, 11, 12, 13
R_REPORTED_PE, R_PRICE_DATE = 14, 15
# historical block, values in columns B..F (FY-4 .. current FY)
R_REV, R_EPS, R_NI, R_FCF = 19, 20, 21, 22
R_DA, R_CAPEX, R_MAINT = 23, 24, 25
R_DEBT, R_CASH, R_EQUITY, R_DIV, R_OCF = 26, 27, 28, 29, 30
# assumptions
R_DISC, R_TERM, R_YEARS = 34, 35, 36
R_PE_METHOD, R_PE_MANUAL, R_PE_ANCHOR, R_PE_CAP, R_PE_FLOOR = 37, 38, 39, 40, 41
R_FCF_METHOD = 42
R_MOS, R_PEG_BUY, R_PEG_SELL = 43, 44, 45
R_G_REV, R_G_EPS, R_G_FCF = 46, 47, 48
R_W_QUAL, R_W_VAL, R_W_STR = 49, 50, 51
R_TOLERANCE = 52
# checklist
R_CHECK0 = 56  # five rows, 56..60

# ------------------------------------------------------------- sheet 2 layout
R_M0 = 5   # core metrics 5..18
(M_REVCAGR, M_EPSCAGR, M_FCFCAGR, M_PE, M_PEG, M_FCFMARGIN, M_FCFMARGIN_R,
 M_ROE, M_DE, M_NETCASH, M_OE, M_OEPS, M_NDOE, M_FCFBASE, M_EXITPE) = range(5, 20)
R_SC0 = 5  # score modules G5..G11
R_DET0 = 23      # score detail 23..31
R_DQ0 = 36       # data-quality checks 36..46
R_DQ_BLOCK, R_DQ_WARN, R_DQ_SCORE = 48, 49, 50
R_FC0 = 54       # forecast rows 54..57 (revenue, eps, fcf, fcf/share)
R_IV0 = 62       # intrinsic value 62..70


def sheet_frame(ws, title, subtitle):
    ws["A1"] = title
    ws["A1"].font = TITLE_F
    ws["A2"] = subtitle
    ws["A2"].font = SUB_F


def section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = SEC_F


def header(ws, row, labels, col0=1):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=label)
        c.font = HEAD_F
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 26


def put(ws, row, col, value, *, fill=None, fmt=None, bold=False, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", size=10, bold=bold, color=INK)
    c.border = BOX
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# =============================================================== input sheet
def build_inputs(ws):
    sheet_frame(
        ws,
        "Stock Evaluation Model: Peter Lynch and Warren Buffett Method (v2)",
        "Fill the amber cells only. Everything else is formula-driven. "
        "Every assumption is checked against history on the Derived Values sheet. "
        "Leaving a default in place is itself flagged.")

    section(ws, 4, "Table: tblCompanyProfile (company and market inputs)")
    header(ws, 5, ["Input Field", "User Input", "Unit", "Required?",
                   "Example / Guidance", "Source Type", "Source URL / Notes", "Model Cell Name"])
    profile = [
        (R_TICKER, "Ticker", "DEMO", "Text", "Yes", "Ticker symbol, e.g. AAPL", "Market data", "Exchange / broker", "inpTicker", None),
        (R_NAME, "Company Name", "Example Co", "Text", "Yes", "Legal company name", "Company filing", "10-K / annual report", "inpCompanyName", None),
        (R_CCY, "Currency", "USD", "Text", "Yes", "Reporting currency", "Company filing", "Annual report", "inpCurrency", None),
        (R_SECTOR, "Sector / Business Type", "Quality compounder", "Text", "Yes", "Sector and business model", "Company filing", "Business description", "inpSector", None),
        (R_PRICE, "Current Share Price", 100, "$/share", "Yes", "Latest market price", "Market data", "Quote source with date", "inpCurrentPrice", USD),
        (R_SHARES, "Shares Outstanding", 1000, "mm shares", "Yes", "Diluted or basic; stay consistent", "Company filing", "Balance sheet / 10-K", "inpSharesOut", MM),
        (R_MOAT, "Business Moat Notes", "Brand + switching cost", "Text", "No", "Moat explanation", "Analyst research", "User notes", "inpMoatNotes", None),
        (R_NOTES, "Analyst Notes", "Replace demo values before publishing", "Text", "No", "Caveats and adjustments made", "User judgment", "Research notes", "inpAnalystNotes", None),
        (R_REPORTED_PE, "Reported P/E (cross-check)", None, "x", "No",
         "Optional. If the data provider quotes a P/E, paste it and the model flags a mismatch "
         "against price divided by current-FY EPS, which usually means one figure is TTM and the other fiscal-year",
         "Market data", "Quote source", "inpReportedPE", MULT),
        (R_PRICE_DATE, "Price As Of", None, "Date", "No", "Date and time the price was taken", "Market data", "Quote source", "inpPriceDate", "yyyy-mm-dd"),
    ]
    for row, label, val, unit, req, guide, stype, url, name, fmt in profile:
        put(ws, row, 1, label)
        put(ws, row, 2, val, fill=INPUT_FILL, fmt=fmt)
        put(ws, row, 3, unit)
        put(ws, row, 4, req)
        put(ws, row, 5, guide, wrap=True)
        put(ws, row, 6, stype)
        put(ws, row, 7, url)
        put(ws, row, 8, name)

    section(ws, 17, "Table: tblHistoricalFinancials (five fiscal years, oldest on the left)")
    header(ws, 18, ["Metric", "FY-4", "FY-3", "FY-2", "FY-1", "Current FY", "Units", "Source / Note"])
    hist = [
        (R_REV, "Revenue", [8000, 9000, 10300, 11800, 13500], "$mm", "Income statement", MM),
        (R_EPS, "EPS", [4.1, 4.7, 5.35, 6.1, 6.85], "$/share", "Diluted EPS", USD),
        (R_NI, "Net Income", [4100, 4700, 5350, 6100, 6850], "$mm", "Income statement", MM),
        (R_FCF, "Free Cash Flow (as reported)", [3600, 4100, 4800, 5500, 6200], "$mm", "Operating cash flow less total capex", MM),
        (R_DA, "Depreciation & Amortization", [600, 640, 690, 735, 780], "$mm", "Cash flow statement", MM),
        (R_CAPEX, "Capital Expenditures", [900, 980, 1080, 1200, 1350], "$mm", "Cash flow statement", MM),
        (R_MAINT, "Maintenance Capex Estimate", [600, 640, 690, 735, 780], "$mm", "Usually near D&A for stable firms; edit for heavy investors", MM),
        (R_DEBT, "Total Debt", [2200, 2100, 2000, 1900, 1800], "$mm", "Balance sheet", MM),
        (R_CASH, "Cash & Equivalents", [900, 1100, 1300, 1500, 1750], "$mm", "Balance sheet", MM),
        (R_EQUITY, "Shareholders' Equity", [9800, 10600, 11600, 12800, 14200], "$mm", "Balance sheet", MM),
        (R_DIV, "Dividend / Share", [0.8, 0.9, 1.0, 1.1, 1.2], "$/share", "Dividend history", USD),
        (R_OCF, "Operating Cash Flow", [4200, 4740, 5490, 6235, 6980], "$mm",
         "Required for the owner-earnings cash base; falls back to net income + D&A if blank", MM),
    ]
    for row, label, vals, unit, note, fmt in hist:
        put(ws, row, 1, label)
        for i, v in enumerate(vals):
            put(ws, row, 2 + i, v, fill=INPUT_FILL, fmt=fmt)
        put(ws, row, 7, unit)
        put(ws, row, 8, note, wrap=True)

    section(ws, 32, "Table: tblModelAssumptions (valuation assumptions, methods and thresholds)")
    header(ws, 33, ["Assumption", "User Input", "Unit", "Purpose / Formula Link", "Model Cell Name"])
    assumptions = [
        (R_DISC, "Discount Rate", 0.10, "%", "DCF present value", "inpDiscountRate", PCT),
        (R_TERM, "Terminal Growth", 0.025, "%", "Gordon growth terminal value; must stay below the discount rate", "inpTerminalGrowth", PCT),
        (R_YEARS, "Projection Years", 5, "Years", "Explicit forecast horizon", "inpProjectionYears", "0"),
        (R_PE_METHOD, "Exit P/E Method", "GrowthImplied", "Manual | GrowthImplied",
         "GrowthImplied ties the exit multiple to the growth you forecast, so a compounder is not "
         "valued on a mature company's multiple", "inpExitPEMethod", None),
        (R_PE_MANUAL, "Target Exit P/E (manual)", 18, "x", "Used only when the method is Manual", "inpTargetPE", MULT),
        (R_PE_ANCHOR, "Exit P/E PEG Anchor", 1.5, "x", "GrowthImplied exit P/E = anchor x forecast EPS growth %", "inpPegAnchor", MULT),
        (R_PE_CAP, "Exit P/E Cap", 30, "x", "Upper bound on the implied multiple", "inpExitPECap", MULT),
        (R_PE_FLOOR, "Exit P/E Floor", 10, "x", "Lower bound on the implied multiple", "inpExitPEFloor", MULT),
        (R_FCF_METHOD, "FCF Base Method", "OwnerEarnings", "Reported | Average3 | OwnerEarnings",
         "The cash figure the DCF compounds. Reported uses the current year alone, which a single "
         "heavy-capex year can distort beyond recognition", "inpFCFBaseMethod", None),
        (R_MOS, "Margin of Safety Required", 0.25, "%", "Buffett/Graham entry cushion", "inpMarginSafety", PCT),
        (R_PEG_BUY, "Lynch PEG Buy Threshold", 1.0, "x", "GARP undervaluation signal", "inpPegBuy", MULT),
        (R_PEG_SELL, "Lynch PEG Sell Threshold", 2.0, "x", "Potentially expensive signal", "inpPegSell", MULT),
        (R_G_REV, "Forecast Revenue Growth", 0.09, "%", "Drives the revenue projection", "inpRevGrowth", PCT),
        (R_G_EPS, "Forecast EPS Growth", 0.10, "%", "Drives the EPS projection and the implied exit multiple", "inpEPSGrowth", PCT),
        (R_G_FCF, "Forecast FCF Growth", 0.09, "%", "Drives the owner cash flow projection", "inpFCFGrowth", PCT),
        (R_W_QUAL, "Quality Score Weight", 0.45, "%", "Overall score weighting", "inpQualityWeight", PCT),
        (R_W_VAL, "Valuation Score Weight", 0.35, "%", "Overall score weighting", "inpValuationWeight", PCT),
        (R_W_STR, "Financial Strength Weight", 0.20, "%", "Overall score weighting", "inpStrengthWeight", PCT),
        (R_TOLERANCE, "Assumption Tolerance vs History", 0.05, "%",
         "How far a forecast growth rate may sit from the matching historical CAGR before it is flagged, "
         "in percentage points, in either direction. Catches a default left untouched", "inpTolerance", PCT),
    ]
    for row, label, val, unit, purpose, name, fmt in assumptions:
        put(ws, row, 1, label)
        put(ws, row, 2, val, fill=INPUT_FILL, fmt=fmt)
        put(ws, row, 3, unit)
        put(ws, row, 4, purpose, wrap=True)
        put(ws, row, 5, name)

    section(ws, 54, "Table: tblQualitativeChecklist (judgment, scored 0 to 5)")
    header(ws, 55, ["Criteria", "User Score (0-5)", "Guidance", "Evidence / Notes", "Model Cell Name"])
    checks = [
        ("Durable competitive advantage / moat", 4, "Brand, network effect, switching cost, cost advantage", "inpMoatScore"),
        ("Management quality / capital allocation", 4, "ROIC discipline, buybacks, debt control", "inpMgmtScore"),
        ("Business simplicity / circle of competence", 4, "Can you explain how it makes money?", "inpSimplicityScore"),
        ("Earnings consistency", 4, "Low cyclicality, few one-off gains or losses", "inpConsistencyScore"),
        ("Reinvestment runway", 4, "Can capital be reinvested at attractive rates?", "inpRunwayScore"),
    ]
    for i, (label, val, guide, name) in enumerate(checks):
        row = R_CHECK0 + i
        put(ws, row, 1, label)
        put(ws, row, 2, val, fill=INPUT_FILL, fmt="0")
        put(ws, row, 3, guide, wrap=True)
        put(ws, row, 4, "Add source evidence")
        put(ws, row, 5, name)

    section(ws, 62, "Table: tblReferenceSources (method and data-source citations)")
    header(ws, 63, ["Source Name", "Source URL", "What to Use It For", "Type", "Date Accessed", "Notes"])
    sources = [
        ("Berkshire Hathaway shareholder letters", "https://www.berkshirehathaway.com/letters/letters.html",
         "Owner earnings, margin of safety, business quality", "Method reference", None, "Record the exact year used"),
        ("Company 10-K / Annual Report", None, "Revenue, EPS, cash flow, debt, share count", "Primary data", None, "Paste the filing URL"),
        ("Market data provider", None, "Current price, share count, quoted P/E", "Market data", None, "Paste the quote URL with date and time"),
        ("Peter Lynch / GARP notes", None, "PEG and growth-at-a-reasonable-price review", "Method reference", None, "Record the source used"),
    ]
    for i, row_vals in enumerate(sources):
        for j, v in enumerate(row_vals):
            put(ws, 64 + i, 1 + j, v, wrap=(j == 5))

    widths(ws, {"A": 34, "B": 17, "C": 13, "D": 12, "E": 46, "F": 16, "G": 26, "H": 30})
    ws.freeze_panes = "B6"


# ============================================================= derived sheet
def build_derived(ws):
    sheet_frame(
        ws,
        "Derived Values: metrics, scoring engine and valuation",
        "Every component score is shown on its own row so the arithmetic can be checked. "
        "A blank metric scores zero and raises a flag; it is never averaged away.")

    # --- core metrics -------------------------------------------------------
    section(ws, 3, "Table: tblCoreMetrics")
    header(ws, 4, ["Metric", "Formula Output", "Units", "Interpretation"])

    def cagr(row_ref):
        """Sign-safe CAGR. A negative or zero base year makes the growth rate
        undefined, so say so rather than returning an empty cell that AVERAGE
        will quietly skip."""
        return (f'=IF(OR(NOT(ISNUMBER({IN}!B{row_ref})),NOT(ISNUMBER({IN}!F{row_ref})),'
                f'{IN}!B{row_ref}<=0,{IN}!F{row_ref}<=0),"n/a",'
                f'({IN}!F{row_ref}/{IN}!B{row_ref})^(1/4)-1)')

    fcf_base = (
        f'=IF({IN}!B{R_FCF_METHOD}="Reported",{IN}!F{R_FCF},'
        f'IF({IN}!B{R_FCF_METHOD}="Average3",AVERAGE({IN}!D{R_FCF}:{IN}!F{R_FCF}),'
        f'IF(ISNUMBER({IN}!F{R_OCF}),{IN}!F{R_OCF}-{IN}!F{R_MAINT},'
        f'{IN}!F{R_NI}+{IN}!F{R_DA}-{IN}!F{R_MAINT})))')

    exit_pe = (
        f'=IF({IN}!B{R_PE_METHOD}="Manual",{IN}!B{R_PE_MANUAL},'
        f'MIN({IN}!B{R_PE_CAP},MAX({IN}!B{R_PE_FLOOR},'
        f'{IN}!B{R_PE_ANCHOR}*{IN}!B{R_G_EPS}*100)))')

    metrics = [
        (M_REVCAGR, "Revenue CAGR", cagr(R_REV), "%", "Four-year compound growth", PCT2),
        (M_EPSCAGR, "EPS CAGR", cagr(R_EPS), "%", "Feeds PEG and the growth score", PCT2),
        (M_FCFCAGR, "FCF CAGR", cagr(R_FCF), "%", "Undefined when the base year is negative; flagged rather than hidden", PCT2),
        (M_PE, "Current P/E", f'=IF(OR(NOT(ISNUMBER({IN}!F{R_EPS})),{IN}!F{R_EPS}<=0),"n/a",{IN}!B{R_PRICE}/{IN}!F{R_EPS})', "x", "Price per dollar of current-FY earnings", MULT),
        (M_PEG, "PEG Ratio",
         f'=IF(OR(NOT(ISNUMBER(B{M_PE})),NOT(ISNUMBER(B{M_EPSCAGR})),B{M_EPSCAGR}<=0),"n/a",'
         f'B{M_PE}/(B{M_EPSCAGR}*100))', "x",
         "P/E divided by EPS growth; undefined without positive growth", MULT),
        (M_FCFMARGIN, "FCF Margin (normalised base)", f'=IFERROR(B{M_FCFBASE}/{IN}!F{R_REV},"n/a")', "%", "Cash conversion on the base the DCF actually uses", PCT2),
        (M_FCFMARGIN_R, "FCF Margin (as reported)", f'=IFERROR({IN}!F{R_FCF}/{IN}!F{R_REV},"n/a")', "%", "Unadjusted, for comparison against the normalised figure", PCT2),
        (M_ROE, "ROE", f'=IFERROR({IN}!F{R_NI}/{IN}!F{R_EQUITY},"n/a")', "%", "Buffett-style business quality proxy", PCT2),
        (M_DE, "Debt / Equity", f'=IFERROR({IN}!F{R_DEBT}/{IN}!F{R_EQUITY},"n/a")', "x", "Balance-sheet leverage", MULT),
        (M_NETCASH, "Net Cash / (Debt)", f'=IFERROR({IN}!F{R_CASH}-{IN}!F{R_DEBT},"n/a")', "$mm", "Balance-sheet flexibility", MM),
        (M_OE, "Owner Earnings", f'=IFERROR({IN}!F{R_NI}+{IN}!F{R_DA}-{IN}!F{R_MAINT},"n/a")', "$mm", "Net income + D&A - maintenance capex", MM),
        (M_OEPS, "Owner Earnings / Share", f'=IFERROR(B{M_OE}/{IN}!B{R_SHARES},"n/a")', "$/share", "Owner earnings per share", USD),
        (M_NDOE, "Net Debt / Owner Earnings", f'=IF(B{M_NETCASH}>=0,0,IFERROR(-B{M_NETCASH}/B{M_OE},"n/a"))', "x", "Years of owner earnings to clear net debt", MULT),
        (M_FCFBASE, "Normalised FCF Base", fcf_base, "$mm", "The cash figure the DCF compounds; set by FCF Base Method", MM),
        (M_EXITPE, "Resolved Exit P/E", exit_pe, "x", "Set by Exit P/E Method; growth-implied unless overridden", MULT),
    ]
    for row, label, formula, unit, interp, fmt in metrics:
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt=fmt)
        put(ws, row, 3, unit)
        put(ws, row, 4, interp, wrap=True)

    # --- score modules ------------------------------------------------------
    ws.cell(row=3, column=6, value="Table: tblScoreModules").font = SEC_F
    header(ws, 4, ["Module", "Score", "Logic"], col0=6)

    D = lambda i: f"B{R_DET0 + i}"  # noqa: E731  component score cells
    modules = [
        ("Growth quality", f'=AVERAGE({D(0)},{D(1)},{D(2)})', "Revenue, EPS and FCF growth, each scored on its own curve"),
        ("Profitability quality", f'=AVERAGE({D(3)},{D(4)})', "ROE and normalised FCF margin"),
        ("Balance sheet", f'=AVERAGE({D(5)},{D(6)})', "Leverage and years of owner earnings to clear net debt"),
        ("Lynch valuation", f'={D(7)}', "PEG against the buy and sell thresholds"),
        ("Buffett valuation", f'={D(8)}', "Upside to fair value against the margin of safety required"),
        ("Qualitative", f'=IFERROR(AVERAGE({IN}!B{R_CHECK0}:{IN}!B{R_CHECK0 + 4})/5*100,0)', "Judgment checklist"),
        ("Overall score",
         f'=IFERROR(AVERAGE(G{R_SC0},G{R_SC0 + 1},G{R_SC0 + 5})*{IN}!B{R_W_QUAL}'
         f'+AVERAGE(G{R_SC0 + 3},G{R_SC0 + 4})*{IN}!B{R_W_VAL}'
         f'+G{R_SC0 + 2}*{IN}!B{R_W_STR},"n/a")', "Weighted composite"),
    ]
    for i, (label, formula, logic) in enumerate(modules):
        row = R_SC0 + i
        put(ws, row, 6, label, bold=(label == "Overall score"))
        put(ws, row, 7, formula, fill=CALC_FILL, fmt="0.0", bold=(label == "Overall score"))
        put(ws, row, 8, logic, wrap=True)

    # --- score detail -------------------------------------------------------
    section(ws, 21, "Table: tblScoreDetail (every component score, with the number behind it)")
    header(ws, 22, ["Component", "Score", "Basis"])

    def growth_score(cell):
        """0% -> 0, 20% and above -> 100. A non-numeric CAGR scores zero, which
        is the point: v1 averaged it away and scored the gap as perfection."""
        return f'=IF(NOT(ISNUMBER({cell})),0,MAX(0,MIN(100,{cell}/0.2*100)))'

    details = [
        ("Revenue growth", growth_score(f"B{M_REVCAGR}"), "0% scores 0, 20%+ scores 100; non-numeric scores 0"),
        ("EPS growth", growth_score(f"B{M_EPSCAGR}"), "Same curve as revenue"),
        ("FCF growth", growth_score(f"B{M_FCFCAGR}"), "Same curve; an undefined CAGR scores 0, not blank"),
        ("Return on equity", f'=IF(NOT(ISNUMBER(B{M_ROE})),0,MAX(0,MIN(100,B{M_ROE}/0.2*100)))', "20% ROE scores 100"),
        ("FCF margin", f'=IF(NOT(ISNUMBER(B{M_FCFMARGIN})),0,MAX(0,MIN(100,B{M_FCFMARGIN}/0.15*100)))', "15% normalised margin scores 100"),
        ("Leverage (debt / equity)", f'=IF(NOT(ISNUMBER(B{M_DE})),0,MAX(0,MIN(100,100-B{M_DE}*50)))', "D/E of 0 scores 100, 2.0x scores 0"),
        ("Net debt cover", f'=IF(NOT(ISNUMBER(B{M_NDOE})),0,MAX(0,MIN(100,100-B{M_NDOE}*25)))', "Net cash scores 100, 4x owner earnings scores 0"),
        ("Lynch PEG", f'=IF(NOT(ISNUMBER(B{M_PEG})),0,'
                      f'IF(B{M_PEG}<={IN}!B{R_PEG_BUY},100,'
                      f'IF(B{M_PEG}>={IN}!B{R_PEG_SELL},20,'
                      f'100-(B{M_PEG}-{IN}!B{R_PEG_BUY})/({IN}!B{R_PEG_SELL}-{IN}!B{R_PEG_BUY})*80)))',
         "100 at or below the buy threshold, 20 at or above the sell threshold"),
        ("Margin of safety", f'=IF(NOT(ISNUMBER(B{R_IV0 + 6})),0,'
                             f'MAX(0,MIN(100,50+(B{R_IV0 + 6}-{IN}!B{R_MOS})*100)))',
         "50 when upside equals the margin of safety required, 100 fifty points above it, 0 fifty below "
         "(v1 read the dollar fair value into this cell, which pinned it at 100 for every company)"),
    ]
    for i, (label, formula, basis) in enumerate(details):
        row = R_DET0 + i
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt="0.0")
        put(ws, row, 3, basis, wrap=True)

    # --- data quality -------------------------------------------------------
    section(ws, 34, "Table: tblDataQuality (what the model does not trust about its own inputs)")
    header(ws, 35, ["Check", "Status", "Severity", "Detail"])

    def flag(cond, blocking, ok_detail, bad_detail):
        # IFERROR around the whole condition: AND() does not short-circuit, so a
        # guarded division still evaluates and returns #VALUE! when the metric it
        # guards against is the text "n/a". A check that cannot run is not a failure.
        sev = "Blocking" if blocking else "Warning"
        safe = f'IFERROR({cond},FALSE)'
        status = f'=IF({safe},"{sev.upper()}","OK")'
        detail = f'=IF({safe},"{bad_detail}","{ok_detail}")'
        return status, sev, detail

    tol = f'{IN}!B{R_TOLERANCE}'
    checks = [
        ("Revenue CAGR computable", *flag(f'NOT(ISNUMBER(B{M_REVCAGR}))', False,
            "Four-year CAGR calculated", "Base year is zero or negative; growth rate undefined")),
        ("EPS CAGR computable", *flag(f'NOT(ISNUMBER(B{M_EPSCAGR}))', False,
            "Four-year CAGR calculated", "Base-year EPS is zero or negative; growth rate undefined")),
        ("FCF CAGR computable", *flag(f'NOT(ISNUMBER(B{M_FCFCAGR}))', False,
            "Four-year CAGR calculated", "Base-year FCF is zero or negative; growth rate undefined")),
        ("Terminal growth below discount rate", *flag(f'{IN}!B{R_TERM}>={IN}!B{R_DISC}', True,
            "Gordon growth is well defined", "Terminal growth meets or exceeds the discount rate; terminal value is meaningless")),
        ("Score weights sum to 100%", *flag(f'ABS({IN}!B{R_W_QUAL}+{IN}!B{R_W_VAL}+{IN}!B{R_W_STR}-1)>0.001', True,
            "Weights sum to 100%", "Score weights do not sum to 100%; the composite is not on a 0-100 scale")),
        ("Price and share count present", *flag(f'OR(NOT(ISNUMBER({IN}!B{R_PRICE})),{IN}!B{R_PRICE}<=0,NOT(ISNUMBER({IN}!B{R_SHARES})),{IN}!B{R_SHARES}<=0)', True,
            "Price and share count present", "Current price or share count is missing or zero")),
        ("Normalised FCF base positive", *flag(f'OR(NOT(ISNUMBER(B{M_FCFBASE})),B{M_FCFBASE}<=0)', True,
            "Cash base is positive", "The normalised cash base is zero or negative; the DCF cannot be relied on")),
        ("Forecast EPS growth consistent with history", *flag(
            f'AND(ISNUMBER(B{M_EPSCAGR}),ABS({IN}!B{R_G_EPS}-B{M_EPSCAGR})>{tol})', False,
            "Forecast sits within tolerance of the historical CAGR",
            "Forecast EPS growth is far from the historical CAGR; it also sets the implied exit multiple, so review it")),
        ("Forecast revenue growth consistent with history", *flag(
            f'AND(ISNUMBER(B{M_REVCAGR}),ABS({IN}!B{R_G_REV}-B{M_REVCAGR})>{tol})', False,
            "Forecast sits within tolerance of the historical CAGR",
            "Forecast revenue growth is far from the historical CAGR; review before publishing")),
        ("Quoted P/E agrees with price / EPS", *flag(
            f'AND(ISNUMBER({IN}!B{R_REPORTED_PE}),ISNUMBER(B{M_PE}),'
            f'ABS({IN}!B{R_REPORTED_PE}-B{M_PE})/B{M_PE}>0.1)', False,
            "Quoted and implied P/E agree, or no quote supplied",
            "Quoted P/E differs from price divided by current-FY EPS by more than 10%; usually trailing-twelve-month versus fiscal-year")),
        ("Reported cash flow is stable", *flag(
            f'OR(MIN({IN}!D{R_FCF}:{IN}!F{R_FCF})<=0,'
            f'AND(MIN({IN}!D{R_FCF}:{IN}!F{R_FCF})>0,MAX({IN}!D{R_FCF}:{IN}!F{R_FCF})/MIN({IN}!D{R_FCF}:{IN}!F{R_FCF})>3))', False,
            "Reported free cash flow is steady over three years",
            "Reported free cash flow is negative or swings more than threefold, use a normalised base, not the reported year")),
    ]
    for i, (label, status, sev, detail) in enumerate(checks):
        row = R_DQ0 + i
        put(ws, row, 1, label)
        put(ws, row, 2, status, fill=CALC_FILL)
        put(ws, row, 3, sev)
        put(ws, row, 4, detail, wrap=True)

    last = R_DQ0 + len(checks) - 1
    put(ws, R_DQ_BLOCK, 1, "Blocking flags", bold=True)
    put(ws, R_DQ_BLOCK, 2, f'=COUNTIF(B{R_DQ0}:B{last},"BLOCKING")', fill=CALC_FILL, fmt="0", bold=True)
    put(ws, R_DQ_BLOCK, 4, "A blocking flag forces the decision to REVIEW", wrap=True)
    put(ws, R_DQ_WARN, 1, "Warnings", bold=True)
    put(ws, R_DQ_WARN, 2, f'=COUNTIF(B{R_DQ0}:B{last},"WARNING")', fill=CALC_FILL, fmt="0", bold=True)
    put(ws, R_DQ_WARN, 4, "Warnings reduce the data-quality score but do not block a call", wrap=True)
    put(ws, R_DQ_SCORE, 1, "Data Quality Score", bold=True)
    put(ws, R_DQ_SCORE, 2, f'=MAX(0,100-40*B{R_DQ_BLOCK}-10*B{R_DQ_WARN})', fill=CALC_FILL, fmt="0.0", bold=True)
    put(ws, R_DQ_SCORE, 4, "A BUY also requires this to be 60 or better", wrap=True)

    # --- forecast -----------------------------------------------------------
    section(ws, 52, "Table: tblForecastProjection (five-year forecast)")
    header(ws, 53, ["Metric", "Current FY", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Source / Driver"])
    rows = [
        ("Revenue", f"={IN}!F{R_REV}", R_G_REV, "Forecast Revenue Growth", MM),
        ("EPS", f"={IN}!F{R_EPS}", R_G_EPS, "Forecast EPS Growth", USD),
        ("Free Cash Flow", f"=B{M_FCFBASE}", R_G_FCF, "Normalised cash base x Forecast FCF Growth", MM),
    ]
    for i, (label, base, grow_row, driver, fmt) in enumerate(rows):
        row = R_FC0 + i
        put(ws, row, 1, label)
        put(ws, row, 2, base, fill=CALC_FILL, fmt=fmt)
        for c in range(3, 8):
            prev = get_column_letter(c - 1)
            put(ws, row, c, f"={prev}{row}*(1+{IN}!B{grow_row})", fill=CALC_FILL, fmt=fmt)
        put(ws, row, 8, driver, wrap=True)

    row = R_FC0 + 3
    put(ws, row, 1, "FCF / Share")
    for c in range(2, 8):
        col = get_column_letter(c)
        put(ws, row, c, f"=IFERROR({col}{R_FC0 + 2}/{IN}!B{R_SHARES},\"n/a\")", fill=CALC_FILL, fmt=USD)
    put(ws, row, 8, "Shares outstanding", wrap=True)

    # --- intrinsic value ----------------------------------------------------
    section(ws, 60, "Table: tblIntrinsicValue (valuation and decision)")
    header(ws, 61, ["Metric", "Value", "Units", "Formula / Explanation"])
    fcfps = R_FC0 + 3
    disc = f"{IN}!B{R_DISC}"
    iv = [
        ("DCF PV of Explicit FCF / Share",
         f'=IFERROR(C{fcfps}/(1+{disc})^1+D{fcfps}/(1+{disc})^2+E{fcfps}/(1+{disc})^3'
         f'+F{fcfps}/(1+{disc})^4+G{fcfps}/(1+{disc})^5,"n/a")', "$/share",
         "Five years of projected owner cash flow per share, discounted", USD),
        ("Terminal Value / Share",
         f'=IF({IN}!B{R_TERM}>={disc},"n/a",IFERROR(G{fcfps}*(1+{IN}!B{R_TERM})/({disc}-{IN}!B{R_TERM}),"n/a"))',
         "$/share", "Gordon growth terminal value", USD),
        ("PV of Terminal Value / Share",
         f'=IFERROR(B{R_IV0 + 1}/(1+{disc})^{IN}!B{R_YEARS},"n/a")', "$/share", "Terminal value discounted back", USD),
        ("DCF Fair Value / Share",
         f'=IFERROR(B{R_IV0}+B{R_IV0 + 2},"n/a")', "$/share", "Buffett-style intrinsic value proxy", USD),
        ("Earnings Multiple Fair Value",
         f'=IFERROR(G{R_FC0 + 1}*B{M_EXITPE}/(1+{disc})^{IN}!B{R_YEARS},"n/a")', "$/share",
         "Year-5 EPS x resolved exit P/E, discounted back", USD),
        # The three published figures are withheld while a blocking flag stands.
        # A fair value the model already knows it cannot support is the one number
        # that must never reach a slide.
        ("Blended Fair Value",
         f'=IF(B{R_DQ_BLOCK}>0,"n/a",IFERROR(AVERAGE(B{R_IV0 + 3},B{R_IV0 + 4}),"n/a"))',
         "$/share", "Average of the two approaches; withheld while a blocking flag stands", USD),
        ("Upside / Downside",
         f'=IF(B{R_DQ_BLOCK}>0,"n/a",IFERROR(B{R_IV0 + 5}/{IN}!B{R_PRICE}-1,"n/a"))',
         "%", "Fair value against the current price", PCT),
        ("Buy Below Price",
         f'=IF(B{R_DQ_BLOCK}>0,"n/a",IFERROR(B{R_IV0 + 5}*(1-{IN}!B{R_MOS}),"n/a"))',
         "$/share", "Margin-of-safety entry price", USD),
        ("Decision Trigger",
         f'=IF(B{R_DQ_BLOCK}>0,"REVIEW",'
         f'IFERROR(IF(AND({IN}!B{R_PRICE}<=B{R_IV0 + 7},G{R_SC0 + 6}>=70,B{R_DQ_SCORE}>=60),"BUY",'
         f'IF(OR({IN}!B{R_PRICE}>B{R_IV0 + 5}*1.15,'
         f'AND(ISNUMBER(B{M_PEG}),B{M_PEG}>={IN}!B{R_PEG_SELL}),G{R_SC0 + 6}<45),"SELL","HOLD")),"REVIEW"))',
         "Decision", "REVIEW whenever a blocking flag is raised, the model will not issue a call on inputs it distrusts", None),
    ]
    for i, (label, formula, unit, expl, fmt) in enumerate(iv):
        row = R_IV0 + i
        bold = label in ("Blended Fair Value", "Decision Trigger")
        put(ws, row, 1, label, bold=bold)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt=fmt, bold=bold)
        put(ws, row, 3, unit)
        put(ws, row, 4, expl, wrap=True)

    widths(ws, {"A": 34, "B": 16, "C": 11, "D": 62, "E": 3, "F": 22, "G": 10, "H": 44})
    ws.freeze_panes = "A5"


# ============================================================== output sheet
def build_output(ws):
    sheet_frame(
        ws,
        "Output Decision, publishable summary",
        "Bound to the deck. Every figure here is a link, so the deck refreshes when the inputs change.")

    section(ws, 3, "Table: tblDecisionSummary")
    header(ws, 4, ["Output Label", "Result", "Linked From", "Interpretation"])
    summary = [
        ("Company", f"={IN}!B{R_NAME}", "Company name", None),
        ("Ticker", f"={IN}!B{R_TICKER}", "Ticker", None),
        ("Current Price", f"={IN}!B{R_PRICE}", "Latest price input", USD),
        ("Blended Fair Value", f"={DV}!B{R_IV0 + 5}", "Intrinsic value estimate", USD),
        ("Upside / Downside", f"={DV}!B{R_IV0 + 6}", "Return potential to fair value", PCT),
        ("Buy Below Price", f"={DV}!B{R_IV0 + 7}", "Margin-of-safety entry price", USD),
        ("Overall Score", f"={DV}!G{R_SC0 + 6}", "Weighted composite, 0-100", "0.0"),
        ("Decision", f"={DV}!B{R_IV0 + 8}", "BUY / HOLD / SELL / REVIEW", None),
        ("Lynch PEG", f"={DV}!B{M_PEG}", "GARP valuation check", MULT),
        ("Buffett Owner Earnings / Share", f"={DV}!B{M_OEPS}", "Owner earnings per share", USD),
        ("Data Quality Score", f"={DV}!B{R_DQ_SCORE}", "0-100; below 60 blocks a BUY", "0.0"),
    ]
    for i, (label, formula, interp, fmt) in enumerate(summary):
        row = 5 + i
        bold = label in ("Decision", "Blended Fair Value")
        put(ws, row, 1, label, bold=bold)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt=fmt, bold=bold)
        put(ws, row, 3, "02_Derived_Values" if formula.startswith(f"={DV}") else "01_Input_Template")
        put(ws, row, 4, interp, wrap=True)

    section(ws, 18, "Table: tblScoreDashboard")
    header(ws, 19, ["Metric", "Score", "Interpretation"])
    dash = [
        ("Growth quality", f"={DV}!G{R_SC0}", '"Strong","Neutral","Weak"'),
        ("Profitability quality", f"={DV}!G{R_SC0 + 1}", '"Strong","Neutral","Weak"'),
        ("Balance sheet", f"={DV}!G{R_SC0 + 2}", '"Strong","Neutral","Weak"'),
        ("Lynch valuation", f"={DV}!G{R_SC0 + 3}", '"Attractive","Neutral","Expensive"'),
        ("Buffett valuation", f"={DV}!G{R_SC0 + 4}", '"MOS attractive","MOS partial","MOS limited"'),
        ("Qualitative", f"={DV}!G{R_SC0 + 5}", '"Strong","Neutral","Weak"'),
    ]
    for i, (label, formula, words) in enumerate(dash):
        row = 20 + i
        a, b, c = [w.strip() for w in words.split(",")]
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt="0.0")
        put(ws, row, 3, f'=IF(B{row}>=75,{a},IF(B{row}>=50,{b},{c}))', fill=CALC_FILL)

    section(ws, 28, "Table: tblValuationBridge (how the fair value is built)")
    header(ws, 29, ["Step", "$ / share", "Note"])
    bridge = [
        ("Current price", f"={IN}!B{R_PRICE}", "What the market asks today"),
        ("DCF fair value", f"={DV}!B{R_IV0 + 3}", "Owner cash flow discounted, including terminal value"),
        ("Earnings multiple fair value", f"={DV}!B{R_IV0 + 4}", "Year-5 EPS on the resolved exit multiple, discounted"),
        ("Blended fair value", f"={DV}!B{R_IV0 + 5}", "The average of the two"),
        ("Buy below", f"={DV}!B{R_IV0 + 7}", "After the margin of safety required"),
    ]
    for i, (label, formula, note) in enumerate(bridge):
        row = 30 + i
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt=USD)
        put(ws, row, 3, note, wrap=True)

    section(ws, 37, "Table: tblAssumptionReview (every assumption against what history actually did)")
    header(ws, 38, ["Assumption", "Used", "History", "Status"])
    review = [
        ("Revenue growth", f"={IN}!B{R_G_REV}", f"={DV}!B{M_REVCAGR}", PCT2),
        ("EPS growth", f"={IN}!B{R_G_EPS}", f"={DV}!B{M_EPSCAGR}", PCT2),
        ("FCF growth", f"={IN}!B{R_G_FCF}", f"={DV}!B{M_FCFCAGR}", PCT2),
    ]
    for i, (label, used, hist, fmt) in enumerate(review):
        row = 39 + i
        put(ws, row, 1, label)
        put(ws, row, 2, used, fill=CALC_FILL, fmt=fmt)
        put(ws, row, 3, hist, fill=CALC_FILL, fmt=fmt)
        put(ws, row, 4, f'=IF(NOT(ISNUMBER(C{row})),"No history",'
                        f'IF(ABS(B{row}-C{row})>{IN}!B{R_TOLERANCE},"Review","In line"))', fill=CALC_FILL)
    row = 42
    put(ws, row, 1, "Exit P/E applied")
    put(ws, row, 2, f"={DV}!B{M_EXITPE}", fill=CALC_FILL, fmt=MULT)
    put(ws, row, 3, f"={DV}!B{M_PE}", fill=CALC_FILL, fmt=MULT)
    put(ws, row, 4, f'=IF(NOT(ISNUMBER(C{row})),"No history",IF(B{row}<C{row}*0.5,"Far below today","In range"))', fill=CALC_FILL)
    row = 43
    put(ws, row, 1, "Cash base used")
    put(ws, row, 2, f"={DV}!B{M_FCFBASE}", fill=CALC_FILL, fmt=MM)
    put(ws, row, 3, f"={IN}!F{R_FCF}", fill=CALC_FILL, fmt=MM)
    put(ws, row, 4, f'=IF(ABS(B{row}-C{row})/MAX(1,ABS(C{row}))>0.2,"Normalised","As reported")', fill=CALC_FILL)

    section(ws, 46, "Table: tblDataQualityFlags (what to say before anyone asks)")
    header(ws, 47, ["Check", "Status", "Detail"])
    for i in range(11):
        row, src = 48 + i, R_DQ0 + i
        put(ws, row, 1, f"={DV}!A{src}")
        put(ws, row, 2, f"={DV}!B{src}", fill=CALC_FILL)
        put(ws, row, 3, f"={DV}!D{src}", fill=CALC_FILL, wrap=True)

    section(ws, 61, "Table: tblFuturePrediction")
    header(ws, 62, ["Prediction Item", "Output", "Driver / Link", "Interpretation"])
    fcps = R_FC0 + 3
    preds = [
        ("Projected Year-5 Revenue", f"={DV}!G{R_FC0}", "Projection", "Future company scale", MM),
        ("Projected Year-5 EPS", f"={DV}!G{R_FC0 + 1}", "Projection", "Future earnings power", USD),
        ("Projected Year-5 FCF", f"={DV}!G{R_FC0 + 2}", "Projection", "Future cash generation", MM),
        ("Projected Year-5 FCF / Share", f"={DV}!G{fcps}", "Projection", "Future owner cash flow", USD),
    ]
    for i, (label, formula, driver, interp, fmt) in enumerate(preds):
        row = 63 + i
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fill=CALC_FILL, fmt=fmt)
        put(ws, row, 3, driver)
        put(ws, row, 4, interp, wrap=True)

    put(ws, 67, 1, "Prediction Narrative")
    put(ws, 67, 2,
        '=IF(B12="REVIEW","Not rated: the model has raised a blocking flag on its own inputs. '
        'Resolve the flags on the Derived Values sheet before quoting a fair value.",'
        'IF(B12="BUY","Quality and valuation align, and the price sits below the margin-of-safety entry point.",'
        'IF(B12="SELL","The price exceeds the blended fair value by more than the tolerance, or the quality '
        'and PEG tests fail.","Neither the valuation nor the quality case is decisive at this price.")))',
        fill=CALC_FILL, wrap=True)
    put(ws, 67, 3, "Rules-based synthesis")
    put(ws, 67, 4, "Plain-English summary", wrap=True)

    put(ws, 70, 1, "Notes and limitations", bold=True)
    put(ws, 71, 1,
        "Educational template, not financial advice. The decision is a rules-based output of the assumptions "
        "on the input sheet, not a view on the company. Check the data-quality flags above before showing "
        "this to anyone.", wrap=True)

    widths(ws, {"A": 34, "B": 18, "C": 22, "D": 56})
    ws.freeze_panes = "A5"


def recalculate(path: Path) -> None:
    """openpyxl writes formulas but no cached results, and Deck Updater reads
    cached results because a browser cannot evaluate Excel formulas. LibreOffice
    opens the workbook, evaluates everything and writes the values back."""
    soffice = "/mnt/skills/public/pptx/scripts/office/soffice.py"
    with tempfile.TemporaryDirectory() as tmp:
        r = subprocess.run([sys.executable, soffice, "--headless", "--convert-to", "xlsx",
                            "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=300)
        out = Path(tmp) / path.name
        if not out.exists():
            raise SystemExit("LibreOffice did not produce a workbook:\n" + (r.stderr or r.stdout))
        shutil.copy(out, path)


def main() -> int:
    dest = Path(sys.argv[sys.argv.index("-o") + 1]) if "-o" in sys.argv else Path("stock_model_v2.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_inputs(wb.create_sheet("01_Input_Template"))
    build_derived(wb.create_sheet("02_Derived_Values"))
    build_output(wb.create_sheet("03_Output_Decision"))
    wb.save(dest)
    recalculate(dest)
    print(f"wrote {dest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
