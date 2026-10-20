"""Builds the Nike financial statement analysis workbook.

Reads data/nike_financials.csv (figures as reported in Nike's Form 10-K filings, $ millions)
and writes nike_financial_analysis.xlsx: the three statements, common-size statements,
ratio analysis (profitability, DuPont, liquidity, efficiency, leverage, cash flow),
charts, a written summary and automatic checks. Every analysis cell is an Excel formula.

Run:  python build_analysis.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
OUT = ROOT / "nike_financial_analysis.xlsx"
YEARS = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]
COLS = ["B", "C", "D", "E", "F", "G"]          # statement columns, FY2021..FY2026
ACOLS = COLS[1:]                               # analysis columns, FY2022..FY2026 (same letters)
FIRST, LAST = ACOLS[0], ACOLS[-1]

# ---------- styles ----------
ARIAL = "Arial"
F_BASE = Font(name=ARIAL, size=10)
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="008000")
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_HEAD = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=ARIAL, size=14, bold=True, color="1F4E79")
F_SECTION = Font(name=ARIAL, size=10, bold=True, color="1F4E79")
F_NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor="1F4E79")
FILL_SECTION = PatternFill("solid", fgColor="DCE6F1")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_BASE = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="808080")
TOP = Border(top=THIN)
TOTAL = Border(top=THIN, bottom=Side(style="double", color="808080"))

M = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
X = '0.00"x";(0.00"x")'
DAYS = '0.0'
EPS = '$0.00'


def load():
    data = {"IS": {}, "BS": {}, "CF": {}}
    with open(ROOT / "data" / "nike_financials.csv", newline="") as f:
        for row in csv.DictReader(f):
            data[row["statement"]][row["item"]] = [float(row[y]) for y in YEARS]
    return data


def header(ws, title, subtitle, years, first_col=2):
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    ws["A2"] = subtitle
    ws["A2"].font = F_NOTE
    ws.cell(row=4, column=1, value="$ millions unless stated").font = F_HEAD
    ws.cell(row=4, column=1).fill = FILL_HEAD
    for i, y in enumerate(years):
        c = ws.cell(row=4, column=first_col + i, value=y)
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(horizontal="right")
    ws.freeze_panes = ws.cell(row=5, column=2)


def section(ws, r, text, width):
    for c in range(1, width + 1):
        ws.cell(row=r, column=c).fill = FILL_SECTION
    ws.cell(row=r, column=1, value=text).font = F_SECTION


# ---------- statements ----------
# Each line: (label, key, kind). kind: "in" = reported input, "sum:<a>:<b>" = sum of rows a..b keys,
# "f:<expr>" = formula using {key} placeholders, "sec" = section heading.
IS_LINES = [
    ("Revenues", "rev", "in:Revenues"),
    ("Cost of sales", "cogs", "in:Cost of sales"),
    ("Gross profit", "gp", "f:{rev}-{cogs}"),
    ("Demand creation expense", "dc", "in:Demand creation expense"),
    ("Operating overhead expense", "oh", "in:Operating overhead expense"),
    ("Total selling and administrative expense", "sga", "f:{dc}+{oh}"),
    ("Operating income (EBIT)", "ebit", "f:{gp}-{sga}"),
    ("Interest expense (income), net", "int", "in:Interest expense (income), net"),
    ("Other (income) expense, net", "oth", "in:Other (income) expense, net"),
    ("Income before income taxes", "pbt", "f:{ebit}-{int}-{oth}"),
    ("Income tax expense", "tax", "in:Income tax expense"),
    ("Net income", "ni", "f:{pbt}-{tax}"),
    ("", None, "blank"),
    ("Diluted EPS ($)", "eps", "in:Diluted EPS ($)"),
    ("Diluted weighted average shares (m)", "shares", "in:Diluted weighted average shares (m)"),
]
IS_TOTALS = {"gp": TOP, "sga": TOP, "ebit": TOP, "pbt": TOP, "ni": TOTAL}

BS_LINES = [
    ("Current assets", None, "sec"),
    ("Cash and equivalents", "cash", "in:Cash and equivalents"),
    ("Short-term investments", "sti", "in:Short-term investments"),
    ("Accounts receivable, net", "ar", "in:Accounts receivable, net"),
    ("Inventories", "inv", "in:Inventories"),
    ("Prepaid expenses and other current assets", "prep", "in:Prepaid expenses and other current assets"),
    ("Total current assets", "ca", "sum:cash:prep"),
    ("Property, plant and equipment, net", "ppe", "in:Property, plant and equipment, net"),
    ("Operating lease right-of-use assets, net", "rou", "in:Operating lease right-of-use assets, net"),
    ("Identifiable intangible assets, net", "intang", "in:Identifiable intangible assets, net"),
    ("Goodwill", "gw", "in:Goodwill"),
    ("Deferred income taxes and other assets", "oa", "in:Deferred income taxes and other assets"),
    ("Total assets", "ta", "f:{ca}+{ppe}+{rou}+{intang}+{gw}+{oa}"),
    ("", None, "blank"),
    ("Current liabilities", None, "sec"),
    ("Current portion of long-term debt", "cpltd", "in:Current portion of long-term debt"),
    ("Notes payable", "np", "in:Notes payable"),
    ("Accounts payable", "ap", "in:Accounts payable"),
    ("Current portion of operating lease liabilities", "cll", "in:Current portion of operating lease liabilities"),
    ("Accrued liabilities", "accr", "in:Accrued liabilities"),
    ("Income taxes payable", "itp", "in:Income taxes payable"),
    ("Total current liabilities", "cl", "sum:cpltd:itp"),
    ("Long-term debt", "ltd", "in:Long-term debt"),
    ("Operating lease liabilities", "ll", "in:Operating lease liabilities"),
    ("Deferred income taxes and other liabilities", "ol", "in:Deferred income taxes and other liabilities"),
    ("Total liabilities", "tl", "f:{cl}+{ltd}+{ll}+{ol}"),
    ("", None, "blank"),
    ("Shareholders' equity", None, "sec"),
    ("Common stock at stated value", "cs", "in:Common stock at stated value"),
    ("Capital in excess of stated value", "apic", "in:Capital in excess of stated value"),
    ("Accumulated other comprehensive income (loss)", "aoci", "in:Accumulated other comprehensive income (loss)"),
    ("Retained earnings (deficit)", "re", "in:Retained earnings (deficit)"),
    ("Total shareholders' equity", "eq", "sum:cs:re"),
    ("Total liabilities and shareholders' equity", "tle", "f:{tl}+{eq}"),
]
BS_TOTALS = {"ca": TOP, "ta": TOTAL, "cl": TOP, "tl": TOP, "eq": TOP, "tle": TOTAL}

CF_LINES = [
    ("Operating activities", None, "sec"),
    ("Net income", "ni", "link:ni"),
    ("Depreciation and amortization", "da", "in:Depreciation and amortization"),
    ("Deferred income taxes", "dt", "in:Deferred income taxes"),
    ("Stock-based compensation", "sbc", "in:Stock-based compensation"),
    ("Impairment and other", "imp", "in:Impairment and other"),
    ("Net foreign currency adjustments", "fx", "in:Net foreign currency adjustments"),
    ("(Increase) decrease in accounts receivable", "d_ar", "in:(Increase) decrease in accounts receivable"),
    ("(Increase) decrease in inventories", "d_inv", "in:(Increase) decrease in inventories"),
    ("(Increase) decrease in prepaid expenses and other assets", "d_prep", "in:(Increase) decrease in prepaid expenses and other assets"),
    ("Increase (decrease) in payables and other liabilities", "d_pay", "in:Increase (decrease) in payables and other liabilities"),
    ("Cash provided by operations", "cfo", "sum:ni:d_pay"),
    ("", None, "blank"),
    ("Investing activities", None, "sec"),
    ("Purchases of short-term investments", "p_sti", "in:Purchases of short-term investments"),
    ("Maturities of short-term investments", "m_sti", "in:Maturities of short-term investments"),
    ("Sales of short-term investments", "s_sti", "in:Sales of short-term investments"),
    ("Additions to property, plant and equipment", "capex", "in:Additions to property plant and equipment"),
    ("Other investing activities", "oinv", "in:Other investing activities"),
    ("Cash provided (used) by investing activities", "cfi", "sum:p_sti:oinv"),
    ("", None, "blank"),
    ("Financing activities", None, "sec"),
    ("Increase (decrease) in notes payable, net", "d_np", "in:Increase (decrease) in notes payable, net"),
    ("Repayment of borrowings", "repay", "in:Repayment of borrowings"),
    ("Proceeds from stock options and other issuances", "opt", "in:Proceeds from stock options and other issuances"),
    ("Repurchase of common stock", "buyback", "in:Repurchase of common stock"),
    ("Dividends paid", "div", "in:Dividends paid"),
    ("Other financing activities", "ofin", "in:Other financing activities"),
    ("Cash provided (used) by financing activities", "cff", "sum:d_np:ofin"),
    ("", None, "blank"),
    ("Effect of exchange rate changes", "fxc", "in:Effect of exchange rate changes"),
    ("Net increase (decrease) in cash", "net", "f:{cfo}+{cfi}+{cff}+{fxc}"),
    ("Cash at beginning of year", "beg", "begin"),
    ("Cash at end of year", "end", "f:{beg}+{net}"),
]
CF_TOTALS = {"cfo": TOP, "cfi": TOP, "cff": TOP, "net": TOP, "end": TOTAL}

SHEET = {"IS": "Income Statement", "BS": "Balance Sheet", "CF": "Cash Flow"}
ROWS = {"IS": {}, "BS": {}, "CF": {}}


def ref(stmt, key, col):
    return f"'{SHEET[stmt]}'!{col}{ROWS[stmt][key]}"


def write_statement(ws, stmt, lines, totals, data, title, subtitle):
    header(ws, title, subtitle, YEARS)
    r = 5
    for label, key, kind in lines:
        if kind == "sec":
            section(ws, r, label, len(COLS) + 1)
        elif kind != "blank":
            ws.cell(row=r, column=1, value=label).font = F_BOLD if key in totals else F_BASE
            ROWS[stmt][key] = r
        r += 1
    for label, key, kind in lines:
        if key is None:
            continue
        r = ROWS[stmt][key]
        fmt = EPS if key == "eps" else ('#,##0.0' if key == "shares" else M)
        for i, col in enumerate(COLS):
            c = ws[f"{col}{r}"]
            if kind.startswith("in:"):
                c.value, c.font = data[stmt][kind[3:]][i], F_INPUT
            elif kind.startswith("sum:"):
                a, b = kind[4:].split(":")
                c.value, c.font = f"=SUM({col}{ROWS[stmt][a]}:{col}{ROWS[stmt][b]})", F_BOLD
            elif kind.startswith("f:"):
                expr = kind[2:]
                for k, rr in ROWS[stmt].items():
                    expr = expr.replace("{" + k + "}", f"{col}{rr}")
                c.value, c.font = "=" + expr, F_BOLD if key in totals else F_BASE
            elif kind.startswith("link:"):
                c.value, c.font = "=" + ref("IS", kind[5:], col), F_LINK
            elif kind == "begin":
                # FY2021 opening cash is an input; later years link to prior year-end cash
                if i == 0:
                    c.value, c.font = 8348, F_INPUT
                else:
                    c.value, c.font = f"={COLS[i - 1]}{ROWS[stmt]['end']}", F_BASE
            c.number_format = fmt
            if key in totals:
                c.border = totals[key]
    ws.column_dimensions["A"].width = 52
    for col in COLS:
        ws.column_dimensions[col].width = 12


# ---------- analysis sheets ----------
def write_analysis(ws, title, subtitle, blocks, notes_col=True):
    """blocks: list of (section title, [(label, key, formula_fn(col, prev), fmt, note)])"""
    header(ws, title, subtitle, YEARS[1:], first_col=3)
    ws.cell(row=4, column=2, value="Formula").font = F_HEAD
    ws.cell(row=4, column=2).fill = FILL_HEAD
    if notes_col:
        ws.cell(row=4, column=8, value="What it tells you").font = F_HEAD
        ws.cell(row=4, column=8).fill = FILL_HEAD
    rows = {}
    r = 5
    for sec_title, items in blocks:
        section(ws, r, sec_title, 8 if notes_col else 7)
        r += 1
        for label, key, fn, fmt, formula_text, note in items:
            ws.cell(row=r, column=1, value=label).font = F_BASE
            ws.cell(row=r, column=2, value=formula_text).font = F_NOTE
            for i, col in enumerate(ACOLS):
                prev = COLS[COLS.index(col) - 1]
                c = ws[f"{col}{r}"]
                c.value = "=" + fn(col, prev, rows, r)
                c.number_format = fmt
                c.font = F_BASE
            if notes_col and note:
                n = ws.cell(row=r, column=8, value=note)
                n.font, n.alignment = F_NOTE, Alignment(wrap_text=True, vertical="top", indent=1)
                ws.row_dimensions[r].height = 26 if len(note) > 75 else 14
            rows[key] = r
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 42
    for col in ACOLS:
        ws.column_dimensions[col].width = 11
    ws.column_dimensions["H"].width = 60
    return rows


def I(k):
    return lambda col, prev, rows, r: ref("IS", k, col)


def avg(stmt, k, col, prev):
    return f"AVERAGE({ref(stmt, k, prev)},{ref(stmt, k, col)})"


def build():
    data = load()
    wb = Workbook()
    cover = wb.active
    cover.title = "Cover"
    summary = wb.create_sheet("Summary")
    charts = wb.create_sheet("Charts")
    ws_is = wb.create_sheet("Income Statement")
    ws_bs = wb.create_sheet("Balance Sheet")
    ws_cf = wb.create_sheet("Cash Flow")
    ws_cs = wb.create_sheet("Common Size")
    ws_r = wb.create_sheet("Ratios")
    ws_ck = wb.create_sheet("Checks")

    src = "Source: NIKE, Inc. Form 10-K filings for fiscal 2021 to 2026 (years end 31 May), as reported."
    write_statement(ws_is, "IS", IS_LINES, IS_TOTALS, data, "Consolidated Income Statement", src)
    write_statement(ws_bs, "BS", BS_LINES, BS_TOTALS, data, "Consolidated Balance Sheet", src + " Balances at 31 May.")
    write_statement(ws_cf, "CF", CF_LINES, CF_TOTALS, data, "Consolidated Cash Flow Statement", src)
    for ws in (ws_is, ws_bs, ws_cf):
        for row in ws.iter_rows(min_row=5, min_col=2, max_col=2):
            for c in row:
                c.fill = FILL_BASE
        ws.cell(row=3, column=2, value="Base year").font = F_NOTE

    # ----- Common size -----
    def pct_of(stmt, k, base_stmt, base):
        return lambda col, prev, rows, r: f"{ref(stmt, k, col)}/{ref(base_stmt, base, col)}"

    cs_is = [(lbl, "cs_" + k, pct_of("IS", k, "IS", "rev"), PCT, "Line / revenue", "")
             for lbl, k, kind in IS_LINES if k and k not in ("eps", "shares")]
    cs_bs = [(lbl, "csb_" + k, pct_of("BS", k, "BS", "ta"), PCT, "Line / total assets", "")
             for lbl, k, kind in BS_LINES if k]
    write_analysis(ws_cs, "Common-Size Statements",
                   "Each line as a % of revenue (income statement) or total assets (balance sheet), so years of different size compare directly.",
                   [("Income statement, % of revenue", cs_is), ("Balance sheet, % of total assets", cs_bs)], notes_col=False)

    # ----- Ratios -----
    def growth(stmt, k):
        return lambda col, prev, rows, r: f"{ref(stmt, k, col)}/{ref(stmt, k, prev)}-1"

    blocks = [
        ("Growth", [
            ("Revenue growth", "g_rev", growth("IS", "rev"), PCT, "Revenue / prior year - 1",
             "Sales peaked in FY2024, then fell almost 10% in FY2025 and were flat in FY2026."),
            ("Operating income growth", "g_ebit", growth("IS", "ebit"), PCT, "EBIT / prior year - 1",
             "Profit falls much faster than sales: costs did not shrink with revenue (operating leverage)."),
            ("Net income growth", "g_ni", growth("IS", "ni"), PCT, "Net income / prior year - 1", ""),
            ("Diluted EPS growth", "g_eps", growth("IS", "eps"), PCT, "EPS / prior year - 1",
             "EPS falls slightly less than net income because buybacks reduced the share count."),
        ]),
        ("Profitability", [
            ("Gross margin", "gm", lambda c, p, R, r: f"{ref('IS','gp',c)}/{ref('IS','rev',c)}", PCT, "Gross profit / revenue",
             "Down about 3 points from FY2022: heavier discounting to clear inventory and product mix."),
            ("Demand creation (marketing) % of revenue", "dc_pct", lambda c, p, R, r: f"{ref('IS','dc',c)}/{ref('IS','rev',c)}", PCT,
             "Demand creation / revenue", "Marketing spend rose from 8% to over 10% of sales while sales fell."),
            ("Operating overhead % of revenue", "oh_pct", lambda c, p, R, r: f"{ref('IS','oh',c)}/{ref('IS','rev',c)}", PCT,
             "Operating overhead / revenue", ""),
            ("Operating (EBIT) margin", "ebit_m", lambda c, p, R, r: f"{ref('IS','ebit',c)}/{ref('IS','rev',c)}", PCT, "EBIT / revenue",
             "The key profitability measure: from 14.3% to 8.2% in four years."),
            ("EBITDA margin", "ebitda_m", lambda c, p, R, r: f"({ref('IS','ebit',c)}+{ref('CF','da',c)})/{ref('IS','rev',c)}", PCT,
             "(EBIT + D&A) / revenue", ""),
            ("Net margin", "net_m", lambda c, p, R, r: f"{ref('IS','ni',c)}/{ref('IS','rev',c)}", PCT, "Net income / revenue", ""),
            ("Effective tax rate", "etr", lambda c, p, R, r: f"{ref('IS','tax',c)}/{ref('IS','pbt',c)}", PCT, "Tax / profit before tax",
             "Rising tax rate is another drag on net income in FY2026."),
            ("Return on equity (ROE)", "roe", lambda c, p, R, r: f"{ref('IS','ni',c)}/{avg('BS','eq',c,p)}", PCT,
             "Net income / average equity", "Still high for a consumer company, but roughly halved since FY2022."),
            ("Return on assets (ROA)", "roa", lambda c, p, R, r: f"{ref('IS','ni',c)}/{avg('BS','ta',c,p)}", PCT,
             "Net income / average total assets", ""),
        ]),
        ("DuPont analysis (ROE = net margin x asset turnover x equity multiplier)", [
            ("Net margin", "dp_nm", lambda c, p, R, r: f"{ref('IS','ni',c)}/{ref('IS','rev',c)}", PCT, "Net income / revenue",
             "This is the component that moved: ROE fell because margins fell."),
            ("Asset turnover", "dp_at", lambda c, p, R, r: f"{ref('IS','rev',c)}/{avg('BS','ta',c,p)}", X, "Revenue / average total assets",
             "Stable at about 1.2x to 1.4x."),
            ("Equity multiplier", "dp_em", lambda c, p, R, r: f"{avg('BS','ta',c,p)}/{avg('BS','eq',c,p)}", X,
             "Average assets / average equity", "Stable leverage, so not the driver."),
            ("ROE (DuPont)", "dp_roe", lambda c, p, R, r: f"{c}{r-3}*{c}{r-2}*{c}{r-1}", PCT, "Product of the three above",
             "Equals ROE above: proves the breakdown adds up."),
        ]),
        ("Liquidity", [
            ("Current ratio", "cr", lambda c, p, R, r: f"{ref('BS','ca',c)}/{ref('BS','cl',c)}", X, "Current assets / current liabilities",
             "Comfortable, though down from 2.6x to 2.0x. FY2026 includes $2bn of debt due within a year."),
            ("Quick ratio", "qr", lambda c, p, R, r: f"({ref('BS','cash',c)}+{ref('BS','sti',c)}+{ref('BS','ar',c)})/{ref('BS','cl',c)}", X,
             "(Cash + investments + receivables) / current liabilities", "Excludes inventory, which may need discounting to sell."),
            ("Cash ratio", "cashr", lambda c, p, R, r: f"({ref('BS','cash',c)}+{ref('BS','sti',c)})/{ref('BS','cl',c)}", X,
             "(Cash + short-term investments) / current liabilities", ""),
        ]),
        ("Efficiency (working capital)", [
            ("Days sales outstanding (DSO)", "dso", lambda c, p, R, r: f"{avg('BS','ar',c,p)}/{ref('IS','rev',c)}*365", DAYS,
             "Average receivables / revenue x 365", "Customers took about 42 days to pay in FY2026, up from 30: worth watching."),
            ("Days inventory outstanding (DIO)", "dio", lambda c, p, R, r: f"{avg('BS','inv',c,p)}/{ref('IS','cogs',c)}*365", DAYS,
             "Average inventory / cost of sales x 365", "Inventory was cut back after the FY2022 glut and is now steady."),
            ("Days payables outstanding (DPO)", "dpo", lambda c, p, R, r: f"{avg('BS','ap',c,p)}/{ref('IS','cogs',c)}*365", DAYS,
             "Average payables / cost of sales x 365", ""),
            ("Cash conversion cycle", "ccc", lambda c, p, R, r: f"{c}{r-3}+{c}{r-2}-{c}{r-1}", DAYS, "DSO + DIO - DPO",
             "Days of sales tied up in working capital. Slightly better than FY2022."),
            ("Inventory turnover", "invt", lambda c, p, R, r: f"{ref('IS','cogs',c)}/{avg('BS','inv',c,p)}", X, "Cost of sales / average inventory", ""),
            ("Asset turnover", "at", lambda c, p, R, r: f"{ref('IS','rev',c)}/{avg('BS','ta',c,p)}", X, "Revenue / average total assets", ""),
        ]),
        ("Leverage and solvency", [
            ("Total debt", "debt", lambda c, p, R, r: f"{ref('BS','cpltd',c)}+{ref('BS','np',c)}+{ref('BS','ltd',c)}", M,
             "Current + long-term debt + notes payable", "Excludes lease liabilities."),
            ("Cash and short-term investments", "liq", lambda c, p, R, r: f"{ref('BS','cash',c)}+{ref('BS','sti',c)}", M, "", ""),
            ("Net cash (debt)", "netcash", lambda c, p, R, r: f"{c}{r-1}-{c}{r-2}", M, "Cash and investments - total debt",
             "Still net cash, but the cushion shrank from $3.6bn to $1.1bn."),
            ("Debt to equity", "de", lambda c, p, R, r: f"{c}{r-3}/{ref('BS','eq',c)}", X, "Total debt / equity", ""),
            ("Debt to EBITDA", "debitda", lambda c, p, R, r: f"{c}{r-4}/({ref('IS','ebit',c)}+{ref('CF','da',c)})", X,
             "Total debt / EBITDA", "Rose from 1.3x to 1.7x because EBITDA fell, not because debt rose."),
            ("Interest cover", "icov", lambda c, p, R, r: f"{ref('IS','ebit',c)}/{ref('CF','int_paid',c)}", X, "EBIT / cash interest paid",
             "Uses interest paid because reported net interest is income (cash earns more than debt costs)."),
        ]),
        ("Cash flow and shareholder returns", [
            ("Cash from operations", "cfo", lambda c, p, R, r: ref("CF", "cfo", c), M, "From cash flow statement", ""),
            ("Capital expenditure", "capex", lambda c, p, R, r: f"-{ref('CF','capex',c)}", M, "Additions to PP&E", ""),
            ("Free cash flow (FCF)", "fcf", lambda c, p, R, r: f"{c}{r-2}-{c}{r-1}", M, "Cash from operations - capex",
             "Fell by half from FY2022 to FY2026."),
            ("FCF margin", "fcf_m", lambda c, p, R, r: f"{c}{r-1}/{ref('IS','rev',c)}", PCT, "FCF / revenue", ""),
            ("FCF conversion", "fcf_conv", lambda c, p, R, r: f"{c}{r-2}/{ref('IS','ni',c)}", PCT, "FCF / net income",
             "Above 100% means profits are backed by cash. FY2026 is weaker because receivables grew."),
            ("Share buybacks", "bb", lambda c, p, R, r: f"-{ref('CF','buyback',c)}", M, "", ""),
            ("Dividends", "dv", lambda c, p, R, r: f"-{ref('CF','div',c)}", M, "", "Dividends kept rising every year."),
            ("Total cash returned to shareholders", "ret", lambda c, p, R, r: f"{c}{r-2}+{c}{r-1}", M, "Buybacks + dividends", ""),
            ("Cash returned as % of FCF", "ret_pct", lambda c, p, R, r: f"{c}{r-1}/{c}{r-6}", PCT, "Returns / FCF",
             "Above 100% in most years: Nike returned more than it generated, funded from its cash pile."),
        ]),
    ]
    # cash interest paid is a supplementary disclosure: add it to the cash flow sheet as an input row
    int_paid = [293, 290, 347, 381, 389, 323]
    r = ws_cf.max_row + 2
    section(ws_cf, r, "Supplementary disclosure", 7)
    ws_cf.cell(row=r + 1, column=1, value="Cash paid for interest, net of capitalized interest").font = F_BASE
    for i, col in enumerate(COLS):
        c = ws_cf[f"{col}{r + 1}"]
        c.value, c.font, c.number_format = int_paid[i], F_INPUT, M
    ROWS["CF"]["int_paid"] = r + 1

    R = write_analysis(ws_r, "Ratio Analysis", "Fiscal years 2022 to 2026. Balance sheet ratios use the average of opening and closing balances.", blocks)

    # ----- Checks -----
    ck_rows = []
    header(ws_ck, "Checks", "Every check should show OK. Reported totals are typed in from the 10-K and compared with the totals the model calculates.", YEARS)
    reported = [
        ("Gross profit agrees to 10-K", "IS", "gp", "Gross profit"),
        ("Income before taxes agrees to 10-K", "IS", "pbt", "Income before income taxes"),
        ("Net income agrees to 10-K", "IS", "ni", "Net income"),
        ("Total current assets agree to 10-K", "BS", "ca", "Total current assets"),
        ("Total assets agree to 10-K", "BS", "ta", "Total assets"),
        ("Total current liabilities agree to 10-K", "BS", "cl", "Total current liabilities"),
        ("Total equity agrees to 10-K", "BS", "eq", "Total shareholders' equity"),
        ("Cash from operations agrees to 10-K", "CF", "cfo", "Cash provided by operations"),
        ("Cash from investing agrees to 10-K", "CF", "cfi", "Cash provided (used) by investing activities"),
        ("Cash from financing agrees to 10-K", "CF", "cff", "Cash provided (used) by financing activities"),
    ]
    r = 5
    for label, stmt, key, src_item in reported:
        ws_ck.cell(row=r, column=1, value=label + " (difference)").font = F_BASE
        for i, col in enumerate(COLS):
            c = ws_ck[f"{col}{r}"]
            c.value = f"=ROUND({ref(stmt, key, col)}-{data[stmt][src_item][i]},0)"
            c.number_format, c.font = M, F_BASE
        ck_rows.append(r)
        r += 1
    logic = [
        ("Balance sheet balances (assets - liabilities - equity)", lambda col: f"ROUND({ref('BS','ta',col)}-{ref('BS','tle',col)},0)"),
        ("Cash flow ending cash = balance sheet cash", lambda col: f"ROUND({ref('CF','end',col)}-{ref('BS','cash',col)},0)"),
        ("Diluted EPS ties to net income / shares (cents)", lambda col: f"ROUND(({ref('IS','ni',col)}/{ref('IS','shares',col)}-{ref('IS','eps',col)})*100,0)"),
    ]
    for label, fn in logic:
        ws_ck.cell(row=r, column=1, value=label).font = F_BASE
        for col in COLS:
            c = ws_ck[f"{col}{r}"]
            c.value, c.number_format, c.font = "=" + fn(col), M, F_BASE
        ck_rows.append(r)
        r += 1
    r += 1
    ws_ck.cell(row=r, column=1, value="Overall status").font = F_BOLD
    ws_ck[f"B{r}"] = f'=IF(SUMPRODUCT(ABS(B{ck_rows[0]}:G{ck_rows[-1]}))=0,"OK","ERROR")'
    ws_ck[f"B{r}"].font = F_BOLD
    ws_ck[f"B{r}"].fill = FILL_KEY
    ck_status = f"Checks!B{r}"
    ws_ck.column_dimensions["A"].width = 52
    for col in COLS:
        ws_ck.column_dimensions[col].width = 12

    # ----- Summary -----
    header(summary, "Nike: Five-Year Financial Statement Analysis", "Key figures, fiscal 2022 to 2026 ($ millions unless stated). All figures link to the statements and ratio sheets.", YEARS[1:], first_col=3)
    summary.cell(row=4, column=2, value="").fill = FILL_HEAD
    summary.cell(row=4, column=8, value="FY22 to FY26").font = F_HEAD
    summary.cell(row=4, column=8).fill = FILL_HEAD
    srows = [
        ("Revenue", lambda c: ref("IS", "rev", c), M, "cagr"),
        ("Revenue growth", lambda c: f"Ratios!{c}{R['g_rev']}", PCT, None),
        ("Gross margin", lambda c: f"Ratios!{c}{R['gm']}", PCT, "pts"),
        ("Operating (EBIT) margin", lambda c: f"Ratios!{c}{R['ebit_m']}", PCT, "pts"),
        ("Net income", lambda c: ref("IS", "ni", c), M, "chg"),
        ("Diluted EPS ($)", lambda c: ref("IS", "eps", c), EPS, "chg"),
        ("Return on equity", lambda c: f"Ratios!{c}{R['roe']}", PCT, "pts"),
        ("Current ratio", lambda c: f"Ratios!{c}{R['cr']}", X, None),
        ("Cash conversion cycle (days)", lambda c: f"Ratios!{c}{R['ccc']}", DAYS, None),
        ("Net cash (debt)", lambda c: f"Ratios!{c}{R['netcash']}", M, "chg"),
        ("Free cash flow", lambda c: f"Ratios!{c}{R['fcf']}", M, "chg"),
        ("Cash returned to shareholders", lambda c: f"Ratios!{c}{R['ret']}", M, None),
    ]
    SR = {}
    r = 5
    for label, fn, fmt, delta in srows:
        summary.cell(row=r, column=1, value=label).font = F_BASE
        for col in ACOLS:
            c = summary[f"{col}{r}"]
            c.value, c.number_format, c.font = "=" + fn(col), fmt, F_LINK
        h = summary[f"H{r}"]
        if delta == "cagr":
            h.value, h.number_format = f"=({LAST}{r}/{FIRST}{r})^(1/4)-1", PCT
            summary[f"I{r}"] = "a year (CAGR)"
        elif delta == "pts":
            h.value, h.number_format = f"=({LAST}{r}-{FIRST}{r})*100", '+0.0" pts";-0.0" pts"'
        elif delta == "chg":
            h.value, h.number_format = f"={LAST}{r}/{FIRST}{r}-1", PCT
        h.font = F_BOLD
        summary[f"I{r}"].font = F_NOTE
        SR[label] = r
        r += 1
    summary.cell(row=r + 1, column=1, value="Checks status").font = F_BOLD
    summary.cell(row=r + 1, column=3, value=f"={ck_status}").font = F_LINK
    r += 3
    section(summary, r, "Key findings", 9)
    findings = [
        "1. Growth has stalled. Revenue rose from $46.7bn (FY2022) to a peak of $51.4bn (FY2024), then fell 9.8% to $46.3bn in FY2025 and was flat at $46.4bn in FY2026.",
        "2. Profitability has roughly halved. EBIT margin fell from 14.3% to 8.2% and net income from $6.0bn to $3.1bn. Diluted EPS fell from $3.75 to $2.10.",
        "3. The margin squeeze comes from both lines. Gross margin is down about 3 points (46.0% to 42.9%), and demand creation (marketing) rose from 8.2% to 10.2% of revenue as Nike spent to win back growth.",
        "4. ROE fell from 43% to 22%. DuPont analysis shows asset turnover and leverage were stable, so the whole decline comes from lower margins.",
        "5. The balance sheet is still strong but the cushion is thinner. Nike remains net cash ($1.1bn), but down from $3.6bn. The current ratio fell from 2.6x to 2.0x, partly because $2bn of debt falls due within a year.",
        "6. Cash returns outran cash generation. Over FY2022 to FY2026, Nike returned $27.6bn through buybacks and dividends against $21.4bn of free cash flow. Buybacks were cut to $0.1bn in FY2026, while the dividend kept rising.",
        "7. Working capital: inventory days are steady at about 103 after the FY2022 glut was cleared, but receivable days rose from 30 to 42, which held back FY2026 cash flow (FCF conversion 70%).",
        "Overall: a financially solid company in a turnaround. The priority to watch is whether gross margin and sales growth recover enough to justify the higher marketing spend.",
    ]
    for t in findings:
        r += 1
        summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        c = summary.cell(row=r, column=1, value=t)
        c.font = F_BOLD if t.startswith("Overall") else F_BASE
        c.alignment = Alignment(wrap_text=True, vertical="top")
        summary.row_dimensions[r].height = 30
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 2
    for col in ACOLS:
        summary.column_dimensions[col].width = 12
    summary.column_dimensions["H"].width = 14
    summary.column_dimensions["I"].width = 14

    # ----- Charts -----
    charts["A1"] = "Charts"
    charts["A1"].font = F_TITLE
    charts["A2"] = "All charts read from the Ratios and Summary sheets. Fiscal years 2022 to 2026."
    charts["A2"].font = F_NOTE
    # chart data block (links), so each chart has a clean range with headers
    charts["A40"] = "Chart data (linked)"
    charts["A40"].font = F_SECTION
    cd = [("Year", None), ("Revenue", f"Summary!{{c}}{SR['Revenue']}"), ("Net income", f"Summary!{{c}}{SR['Net income']}"),
          ("Gross margin", f"Ratios!{{c}}{R['gm']}"), ("EBIT margin", f"Ratios!{{c}}{R['ebit_m']}"), ("Net margin", f"Ratios!{{c}}{R['net_m']}"),
          ("Free cash flow", f"Ratios!{{c}}{R['fcf']}"), ("Cash returned", f"Ratios!{{c}}{R['ret']}"),
          ("DSO", f"Ratios!{{c}}{R['dso']}"), ("DIO", f"Ratios!{{c}}{R['dio']}"), ("DPO", f"Ratios!{{c}}{R['dpo']}"),
          ("ROE", f"Ratios!{{c}}{R['roe']}"), ("Demand creation %", f"Ratios!{{c}}{R['dc_pct']}")]
    CD = {}
    for j, (name, f) in enumerate(cd):
        rr = 41 + j
        charts.cell(row=rr, column=1, value=name).font = F_BASE
        for i, col in enumerate(ACOLS):
            c = charts.cell(row=rr, column=2 + i)
            if f is None:
                c.value, c.font = YEARS[1 + i], F_BOLD
            else:
                c.value, c.font = "=" + f.format(c=col), F_LINK
                c.number_format = PCT if ("margin" in name or name in ("ROE", "Demand creation %")) else (DAYS if name in ("DSO", "DIO", "DPO") else M)
        CD[name] = rr

    BLUE, ORANGE, GRAY = "2A78D6", "EB6834", "898781"
    cats = Reference(charts, min_col=2, max_col=6, min_row=CD["Year"])

    def add_series(ch, names, colors, line=False):
        for name, color in zip(names, colors):
            ref_ = Reference(charts, min_col=1, max_col=6, min_row=CD[name])
            ch.add_data(ref_, from_rows=True, titles_from_data=True)
            s = ch.series[-1]
            if line:
                s.graphicalProperties.line.solidFill = color
                s.graphicalProperties.line.width = 28575
                s.smooth = False
            else:
                s.graphicalProperties.solidFill = color
                s.graphicalProperties.line.solidFill = color
        ch.set_categories(cats)

    def style(ch, title, ytitle, fmt):
        ch.title = title
        ch.y_axis.title = ytitle
        ch.y_axis.numFmt = fmt
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
        ch.legend.position = "b"
        ch.height, ch.width = 7.5, 15
        ch.x_axis.delete = False
        ch.y_axis.delete = False

    c1 = BarChart(); add_series(c1, ["Revenue", "Net income"], [BLUE, ORANGE]); style(c1, "Revenue and net income ($m)", "$m", "#,##0")
    c2 = LineChart(); add_series(c2, ["Gross margin", "EBIT margin", "Net margin"], [BLUE, ORANGE, GRAY], line=True); style(c2, "Margins", "% of revenue", "0%")
    c3 = BarChart(); add_series(c3, ["Free cash flow", "Cash returned"], [BLUE, ORANGE]); style(c3, "Free cash flow vs cash returned to shareholders ($m)", "$m", "#,##0")
    c4 = LineChart(); add_series(c4, ["DSO", "DIO", "DPO"], [BLUE, ORANGE, GRAY], line=True); style(c4, "Working capital days", "Days", "0")
    c5 = LineChart(); add_series(c5, ["ROE"], [BLUE], line=True); style(c5, "Return on equity", "ROE", "0%"); c5.legend = None
    c6 = LineChart(); add_series(c6, ["Demand creation %"], [ORANGE], line=True); style(c6, "Demand creation (marketing) as % of revenue", "% of revenue", "0.0%"); c6.legend = None
    for ch, anchor in [(c1, "A4"), (c2, "J4"), (c3, "A20"), (c4, "J20")]:
        charts.add_chart(ch, anchor)
    charts.add_chart(c5, "A56")
    charts.add_chart(c6, "J56")
    charts.column_dimensions["A"].width = 18

    # ----- Cover -----
    cover["A1"] = "Nike, Inc.: Financial Statement Analysis"
    cover["A1"].font = Font(name=ARIAL, size=18, bold=True, color="1F4E79")
    cover["A2"] = "Five fiscal years, FY2022 to FY2026 (years ending 31 May). Figures in $ millions as reported in Nike's Form 10-K filings."
    cover["A2"].font = F_NOTE
    lines = [
        ("Purpose", "Assess how Nike's growth, profitability, efficiency, liquidity, leverage and cash generation have changed over five years."),
        ("Data", "Income statement, balance sheet and cash flow statement typed in from the 10-K (FY2021 is used only as the base year for growth and averages)."),
        ("Method", "Common-size statements, trend (growth) analysis, ratio analysis, DuPont breakdown of ROE and a cash flow review."),
        ("", ""),
        ("Sheets", ""),
        ("Summary", "Key figures and written findings"),
        ("Charts", "Six charts of the main trends"),
        ("Income Statement / Balance Sheet / Cash Flow", "Source statements (blue inputs, black formulas)"),
        ("Common Size", "Each line as a % of revenue or total assets"),
        ("Ratios", "41 ratios and measures, each with its formula and what it shows"),
        ("Checks", "Model totals reconciled to the 10-K"),
        ("", ""),
        ("Colour key", ""),
        ("Blue text", "Input typed in from the 10-K"),
        ("Black text", "Formula"),
        ("Green text", "Link to another sheet"),
        ("", ""),
        ("Checks status", f"={ck_status}"),
        ("Sources", "NIKE, Inc. Form 10-K, fiscal 2021, 2022, 2023, 2025 and 2026, from SEC EDGAR (sec.gov)."),
    ]
    for i, (a, b) in enumerate(lines):
        r = 4 + i
        cover.cell(row=r, column=1, value=a).font = F_BOLD
        c = cover.cell(row=r, column=2, value=b)
        c.font = F_LINK if str(b).startswith("=") else F_BASE
        if a == "Blue text":
            c.font = F_INPUT
        if a == "Green text":
            c.font = F_LINK
        if a == "Checks status":
            c.fill = FILL_KEY
            c.font = F_BOLD
    cover.column_dimensions["A"].width = 44
    cover.column_dimensions["B"].width = 110

    # print setup
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_view.showGridLines = False
    wb.save(OUT)
    print(f"Wrote {OUT.name}")


if __name__ == "__main__":
    build()
