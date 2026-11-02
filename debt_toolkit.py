"""
Debt Headroom & Compliance Toolkit (v4 add-on to the loan repayment template)

Adds 7 sheets to the workbook built by build_template.py:
  Blind_Spots            - prioritised checklist of commonly missed items
  Covenant_Monitor       - DSCR / gearing / TNW tests vs facility covenants,
                           MFRS 101 reclassification flag
  Feed_Stress            - corn / SBM / USD-MYR shock -> DSCR heat-map and
                           break-even feed prices before covenant breach
  Debt_Service_Calendar  - 24-month principal + interest/profit cash calendar
  MFRS7_Risk             - undiscounted contractual cash-flow maturity table
                           (MFRS 7.39) + interest-rate sensitivity
  MFRS123_Capitalisation - profit on availability-period tranches that may
                           need capitalising into CWIP (MFRS 123)
  Change_Log             - what changed in v4 and why

All yellow cells are ILLUSTRATIVE PLACEHOLDERS - replace with actuals /
facility-letter terms before relying on outputs.
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import datetime as dt

HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
AMBER_FILL = PatternFill("solid", fgColor="FFE699")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="top", wrap_text=True)
NOTE_FONT = Font(italic=True, size=9, color="595959")

NUM = '#,##0.00;[Red](#,##0.00);"-"'
NUM0 = '#,##0;[Red](#,##0);"-"'
PCT = '0.0%;[Red]-0.0%'
X = '0.00"x";[Red]-0.00"x"'
DATE = 'dd-mmm-yyyy'
MONTH = 'mmm-yyyy'


def _title(ws, text, span="A1:H1"):
    ws["A1"] = text
    ws["A1"].font = HDR_FONT
    ws["A1"].fill = HDR_FILL
    ws.merge_cells(span)


def _section(ws, row, text, span_to="H"):
    c = ws.cell(row=row, column=1, value=text)
    c.font = SUB_FONT
    c.fill = SUB_FILL
    ws.merge_cells(f"A{row}:{span_to}{row}")


def _hdr(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER


def _line(ws, row, label, value, fmt=NUM, kind="calc", note=None):
    """label in A, value in B, note in C. kind: input|calc|total"""
    ws.cell(row=row, column=1, value=label).border = BORDER
    c = ws.cell(row=row, column=2, value=value)
    c.number_format = fmt
    c.border = BORDER
    c.fill = {"input": INPUT_FILL, "calc": CALC_FILL, "total": TOTAL_FILL}[kind]
    if note:
        n = ws.cell(row=row, column=3, value=note)
        n.font = NOTE_FONT
    return f"$B${row}"


def _rag(ws, rng, formula_red, formula_amber, formula_green):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula_red], fill=RED_FILL))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula_amber], fill=AMBER_FILL))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula_green], fill=GREEN_FILL))


# ---------------------------------------------------------------------------
# Covenant_Monitor
# ---------------------------------------------------------------------------
def _covenant_monitor(wb):
    ws = wb.create_sheet("Covenant_Monitor")
    _title(ws, "COVENANT MONITOR — DSCR / GEARING / TNW HEADROOM (MFRS 101 classification risk)")
    ws["A2"] = ("Covenants are normally tested on the BORROWER / GUARANTOR entity's audited FS. "
                "Run per entity (copy this sheet) — Group figures are shown for a first view.")
    ws["A2"].font = NOTE_FONT
    ws.merge_cells("A2:H2")
    R = {}

    _section(ws, 3, "A. FINANCIAL INPUTS — LTM / audited (YELLOW = placeholder, replace)")
    R["rev"] = _line(ws, 4, "Revenue (RM)", 180_000_000, NUM0, "input")
    R["ebitda"] = _line(ws, 5, "EBITDA (RM)", 9_000_000, NUM0, "input",
                        "Exclude MFRS 141 biological-asset FV gains if the facility letter does")
    R["tax"] = _line(ws, 6, "Tax paid (RM)", 900_000, NUM0, "input")
    R["cash"] = _line(ws, 7, "Cash & bank balances (RM)", 4_000_000, NUM0, "input")
    R["equity"] = _line(ws, 8, "Total equity (RM)", 35_000_000, NUM0, "input")
    R["intang"] = _line(ws, 9, "Intangibles / goodwill / amounts due from directors (RM)", 0, NUM0, "input",
                        "Banks usually deduct these for Tangible Net Worth")
    R["trade"] = _line(ws, 10, "Trade lines & others: OD / BA / RC / TR / HP / leases (RM)", 8_000_000, NUM0, "input",
                       "NOT in the amortisation schedules — omitting understates gearing")
    R["test"] = _line(ws, 11, "Covenant test date", dt.datetime(2025, 12, 31), DATE, "input")

    _section(ws, 13, "B. DEBT SERVICE NEXT 12 MONTHS & BORROWINGS (auto from schedules)")
    def ds(sheet, col):
        return (f'=SUMIFS({sheet}!${col}$9:${col}$400,{sheet}!$B$9:$B$400,">"&{R["test"]},'
                f'{sheet}!$B$9:$B$400,"<="&EDATE({R["test"]},12))')
    R["ds_mbb"] = _line(ws, 14, "Maybank CMTF-i — instalments (P + profit)", ds("Amortisation_MBB", "F"))
    R["ds_alb"] = _line(ws, 15, "Alliance Bank FL — principal + interest", ds("Amortisation_ALB", "G"))
    R["ds_mbsb"] = _line(ws, 16, "MBSB Tawarruq — EPM instalments", ds("Amortisation_MBSB", "F"))
    R["ds_trade"] = _line(ws, 17, "Trade-line interest / HP / lease payments (estimate)", 600_000, NUM0, "input")
    R["ds"] = _line(ws, 18, "TOTAL DEBT SERVICE (RM)", "=SUM(B14:B17)", NUM, "total")

    def bal(sheet, col):
        return f'=SUMIFS({sheet}!${col}$9:${col}$400,{sheet}!$B$9:$B$400,">"&{R["test"]})'
    _line(ws, 19, "Maybank principal outstanding", bal("Amortisation_MBB", "G"))
    _line(ws, 20, "Alliance Bank principal outstanding", bal("Amortisation_ALB", "F"))
    _line(ws, 21, "MBSB principal outstanding", bal("Amortisation_MBSB", "E"))
    _line(ws, 22, "Trade lines & others", f"={R['trade']}")
    R["debt"] = _line(ws, 23, "TOTAL BORROWINGS (RM)", "=SUM(B19:B22)", NUM, "total")

    _section(ws, 25, "C. RATIOS (definitions vary — align to each facility letter)")
    R["dscr"] = _line(ws, 26, "DSCR = (EBITDA − tax) ÷ debt service",
                      f"=IFERROR(({R['ebitda']}-{R['tax']})/{R['ds']},0)", X)
    R["gear"] = _line(ws, 27, "Gearing = total borrowings ÷ equity",
                      f"=IFERROR({R['debt']}/{R['equity']},0)", X)
    R["ngear"] = _line(ws, 28, "Net gearing = (borrowings − cash) ÷ equity",
                       f"=IFERROR(({R['debt']}-{R['cash']})/{R['equity']},0)", X)
    R["nde"] = _line(ws, 29, "Net debt ÷ EBITDA",
                     f"=IFERROR(({R['debt']}-{R['cash']})/{R['ebitda']},0)", X)
    R["tnw"] = _line(ws, 30, "Tangible net worth (RM)", f"={R['equity']}-{R['intang']}", NUM0)

    _section(ws, 32, "D. COVENANT REGISTER — thresholds are PLACEHOLDERS: key in from each facility letter")
    _hdr(ws, 33, ["Bank / Facility", "Covenant", "Test", "Threshold", "Actual",
                  "Headroom %", "Status", "MFRS 101 consequence if breached at reporting date"])
    cov = [
        ("Maybank CMTF-i", "DSCR", "Min", 1.25, f"={R['dscr']}", X),
        ("Maybank CMTF-i", "Gearing", "Max", 2.00, f"={R['gear']}", X),
        ("Alliance Bank FL", "DSCR", "Min", 1.25, f"={R['dscr']}", X),
        ("Alliance Bank FL", "Net gearing", "Max", 2.50, f"={R['ngear']}", X),
        ("MBSB Tawarruq", "FSCR / DSCR", "Min", 1.50, f"={R['dscr']}", X),
        ("MBSB Tawarruq", "Tangible net worth (RM)", "Min", 20_000_000, f"={R['tnw']}", NUM0),
    ]
    for i, (bank, name, test, thr, act, fmt) in enumerate(cov, start=34):
        vals = [bank, name, test, thr, act]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.fill = INPUT_FILL if c <= 4 else CALC_FILL
        ws.cell(row=i, column=4).number_format = fmt
        ws.cell(row=i, column=5).number_format = fmt
        h = ws.cell(row=i, column=6,
                    value=f'=IF(D{i}="","",IF(C{i}="Min",(E{i}-D{i})/D{i},(D{i}-E{i})/D{i}))')
        h.number_format = PCT; h.border = BORDER; h.fill = CALC_FILL
        s = ws.cell(row=i, column=7,
                    value=f'=IF(D{i}="","INPUT",IF(F{i}<0,"BREACH",IF(F{i}<0.15,"WATCH","OK")))')
        s.border = BORDER; s.alignment = CENTER; s.font = Font(bold=True)
        m = ws.cell(row=i, column=8,
                    value=(f'=IF(G{i}="BREACH","Whole facility → CURRENT unless waiver obtained ON/BEFORE '
                           f'reporting date (MFRS 101.74). Also check cross-default on other banks.",'
                           f'IF(G{i}="WATCH","Headroom <15% — pre-agree waiver / relaxation now; '
                           f'disclose covenant + risk (MFRS 101.76ZA)","Non-current OK; disclose covenant terms (MFRS 101.76ZA)"))')
                    )
        m.border = BORDER; m.alignment = WRAP
        dv = DataValidation(type="list", formula1='"Min,Max"', allow_blank=False)
        ws.add_data_validation(dv); dv.add(f"C{i}")
    ws.conditional_formatting.add("G34:G39", CellIsRule(operator="equal", formula=['"BREACH"'], fill=RED_FILL))
    ws.conditional_formatting.add("G34:G39", CellIsRule(operator="equal", formula=['"WATCH"'], fill=AMBER_FILL))
    ws.conditional_formatting.add("G34:G39", CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN_FILL))

    ws["A41"] = "Overall status"
    ws["A41"].font = SUB_FONT
    ws["B41"] = '=IF(COUNTIF(G34:G39,"BREACH")>0,"BREACH",IF(COUNTIF(G34:G39,"WATCH")>0,"WATCH","OK"))'
    ws["B41"].font = Font(bold=True); ws["B41"].alignment = CENTER
    ws.conditional_formatting.add("B41", CellIsRule(operator="equal", formula=['"BREACH"'], fill=RED_FILL))
    ws.conditional_formatting.add("B41", CellIsRule(operator="equal", formula=['"WATCH"'], fill=AMBER_FILL))
    ws.conditional_formatting.add("B41", CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN_FILL))
    ws["A42"] = ("Cross-default: a breach on ONE bank usually triggers default on the others — "
                 "read the cross-default clause in every facility letter.")
    ws["A42"].font = NOTE_FONT
    ws.merge_cells("A42:H42")

    for col, w in zip("ABCDEFGH", [58, 20, 8, 16, 16, 12, 10, 70]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"
    return R


# ---------------------------------------------------------------------------
# Feed_Stress
# ---------------------------------------------------------------------------
def _feed_stress(wb, CR):
    ws = wb.create_sheet("Feed_Stress")
    _title(ws, "FEED COST STRESS TEST — how far can corn / SBM / USD-MYR move before a covenant breach?",
           "A1:I1")
    R = {}
    _section(ws, 3, "A. FEED & COMMODITY INPUTS (placeholders — replace with mill data & latest CFR quotes)", "I")
    R["mt"] = _line(ws, 4, "Annual compound feed produced / purchased (MT)", 60_000, NUM0, "input")
    R["corn_pct"] = _line(ws, 5, "Corn inclusion (% of feed)", 0.55, PCT, "input")
    R["sbm_pct"] = _line(ws, 6, "Soybean meal inclusion (% of feed)", 0.28, PCT, "input")
    R["corn"] = _line(ws, 7, "Corn price — USD/MT CFR Port Klang", 220, NUM, "input")
    R["sbm"] = _line(ws, 8, "SBM price — USD/MT CFR Port Klang", 400, NUM, "input")
    R["fx"] = _line(ws, 9, "USD/MYR", 4.20, '0.0000', "input")
    R["other"] = _line(ws, 10, "Other ingredients & premix — RM per MT of feed", 350, NUM, "input",
                       "Wheat pollard, palm oil, DCP, vitamins, amino acids, etc.")
    R["mill"] = _line(ws, 11, "Milling, packing & delivery — RM per MT", 180, NUM, "input")
    R["fcr"] = _line(ws, 12, "Broiler FCR (kg feed / kg live weight)", 1.60, '0.000', "input")
    R["pass"] = _line(ws, 13, "Selling-price pass-through of feed cost increases (%)", 0.30, PCT, "input",
                      "Since ceiling prices were floated, pass-through depends on market — be conservative")
    R["dscr_min"] = _line(ws, 14, "DSCR covenant to protect (min)", "=MAX(Covenant_Monitor!D34,Covenant_Monitor!D36,Covenant_Monitor!D38)", X,
                          "calc", "Tightest DSCR in Covenant_Monitor")

    _section(ws, 16, "B. BASE CASE", "I")
    R["imp"] = _line(ws, 17, "Imported commodity cost — RM per MT feed",
                     f"=({R['corn_pct']}*{R['corn']}+{R['sbm_pct']}*{R['sbm']})*{R['fx']}")
    R["cpm"] = _line(ws, 18, "Total feed cost — RM per MT", f"={R['imp']}+{R['other']}+{R['mill']}")
    R["share"] = _line(ws, 19, "Imported share of feed cost", f"=IFERROR({R['imp']}/{R['cpm']},0)", PCT)
    R["annual"] = _line(ws, 20, "Annual feed cost (RM)", f"={R['cpm']}*{R['mt']}", NUM0)
    R["perkg"] = _line(ws, 21, "Feed cost per kg live weight (RM)", f"={R['cpm']}/1000*{R['fcr']}", '0.000')
    R["cfads"] = _line(ws, 22, "Cash available for debt service = EBITDA − tax (RM)",
                       f"=Covenant_Monitor!{CR['ebitda'].replace('$','')}-Covenant_Monitor!{CR['tax'].replace('$','')}", NUM0)
    R["ds"] = _line(ws, 23, "Debt service next 12m (RM)", f"=Covenant_Monitor!{CR['ds'].replace('$','')}", NUM0)
    R["dscr"] = _line(ws, 24, "Base DSCR", f"=IFERROR({R['cfads']}/{R['ds']},0)", X, "total")
    R["cushion"] = _line(ws, 25, "Cash cushion above covenant (RM)", f"={R['cfads']}-{R['dscr_min']}*{R['ds']}", NUM0)

    _section(ws, 27, "C. BREAK-EVEN — the single number to watch each month", "I")
    absorb = f"(1-{R['pass']})"
    _line(ws, 28, "Max rise in corn & SBM USD prices (FX flat) before breach",
          f"=IFERROR({R['cushion']}/({R['mt']}*{R['imp']}*{absorb}),0)", PCT, "total")
    _line(ws, 29, "Break-even CORN price USD/MT (SBM & FX flat)",
          f"=IFERROR({R['corn']}+{R['cushion']}/({R['mt']}*{R['corn_pct']}*{R['fx']}*{absorb}),0)", NUM, "total")
    _line(ws, 30, "Break-even SBM price USD/MT (corn & FX flat)",
          f"=IFERROR({R['sbm']}+{R['cushion']}/({R['mt']}*{R['sbm_pct']}*{R['fx']}*{absorb}),0)", NUM, "total")
    _line(ws, 31, "Break-even USD/MYR (commodities flat)",
          f"=IFERROR({R['fx']}+{R['cushion']}/({R['mt']}*({R['corn_pct']}*{R['corn']}+{R['sbm_pct']}*{R['sbm']})*{absorb}),0)",
          '0.0000', "total")
    ws["A32"] = ("Negative cushion = already below covenant at base case. "
                 "Track these break-evens against weekly CBOT / CFR quotes and BNM reference rate.")
    ws["A32"].font = NOTE_FONT
    ws.merge_cells("A32:I32")

    _section(ws, 34, "D. DSCR HEAT-MAP — rows: corn & SBM USD price change | columns: USD/MYR change", "I")
    fx_moves = [-0.05, -0.025, 0, 0.025, 0.05, 0.075, 0.10]
    com_moves = [-0.10, -0.05, 0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30]
    ws.cell(row=35, column=2, value="Commodity Δ \\ FX Δ").font = HDR_FONT
    ws.cell(row=35, column=2).fill = HDR_FILL
    for j, f in enumerate(fx_moves, start=3):
        c = ws.cell(row=35, column=j, value=f)
        c.number_format = '+0.0%;-0.0%;0.0%'; c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER
    for i, m in enumerate(com_moves, start=36):
        c = ws.cell(row=i, column=2, value=m)
        c.number_format = '+0%;-0%;0%'; c.font = SUB_FONT; c.fill = SUB_FILL; c.alignment = CENTER
        for j in range(3, 3 + len(fx_moves)):
            col = get_column_letter(j)
            f = (f"=IFERROR(({R['cfads']}-{R['mt']}*{R['imp']}*((1+$B{i})*(1+{col}$35)-1)*{absorb})"
                 f"/{R['ds']},0)")
            cell = ws.cell(row=i, column=j, value=f)
            cell.number_format = X; cell.border = BORDER; cell.alignment = CENTER
    last = 35 + len(com_moves)
    rng = f"C36:{get_column_letter(2 + len(fx_moves))}{last}"
    _rag(ws, rng, f"C36<{R['dscr_min']}", f"C36<{R['dscr_min']}*1.15", f"C36>={R['dscr_min']}*1.15")
    ws.cell(row=last + 1, column=1,
            value="Red = below covenant · Amber = within 15% of covenant · Green = comfortable").font = NOTE_FONT

    for col, w in zip("ABCDEFGHI", [58, 18, 11, 11, 11, 11, 11, 11, 11]):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["C"].width = 14
    ws.freeze_panes = "A3"


# ---------------------------------------------------------------------------
# Debt_Service_Calendar
# ---------------------------------------------------------------------------
def _calendar(wb):
    ws = wb.create_sheet("Debt_Service_Calendar")
    _title(ws, "DEBT SERVICE CALENDAR — next 24 months (treasury cash planning)", "A1:K1")
    ws["A3"] = "First month:"
    ws["B3"] = dt.datetime(2026, 1, 1)
    ws["B3"].number_format = MONTH; ws["B3"].fill = INPUT_FILL
    ws["D3"] = ("Maybank column covers only the tranche selected on Amortisation_MBB!B3. "
                "Add rows for other tranches / trade lines once their schedules exist.")
    ws["D3"].font = NOTE_FONT
    ws.merge_cells("D3:K3")
    _hdr(ws, 5, ["Month", "MBB principal", "MBB profit", "ALB principal", "ALB interest",
                 "MBSB principal", "MBSB profit", "Total principal", "Total interest/profit",
                 "TOTAL CASH OUT", "Cumulative"])
    srcs = [("Amortisation_MBB", "G"), ("Amortisation_MBB", "E"),
            ("Amortisation_ALB", "F"), ("Amortisation_ALB", "E"),
            ("Amortisation_MBSB", "E"), ("Amortisation_MBSB", "D")]
    for n in range(24):
        r = 6 + n
        a = ws.cell(row=r, column=1, value=("=$B$3" if n == 0 else f"=EDATE(A{r-1},1)"))
        a.number_format = MONTH; a.border = BORDER
        for k, (sh, col) in enumerate(srcs, start=2):
            c = ws.cell(row=r, column=k,
                        value=(f'=SUMIFS({sh}!${col}$9:${col}$400,{sh}!$B$9:$B$400,">="&$A{r},'
                               f'{sh}!$B$9:$B$400,"<="&EOMONTH($A{r},0))'))
            c.number_format = NUM; c.border = BORDER
        for k, f in [(8, f"=B{r}+D{r}+F{r}"), (9, f"=C{r}+E{r}+G{r}"), (10, f"=H{r}+I{r}"),
                     (11, f"=J{r}" if n == 0 else f"=K{r-1}+J{r}")]:
            c = ws.cell(row=r, column=k, value=f)
            c.number_format = NUM; c.border = BORDER; c.fill = CALC_FILL
    tr = 30
    ws.cell(row=tr, column=1, value="TOTAL").font = SUB_FONT
    for k in range(2, 11):
        col = get_column_letter(k)
        c = ws.cell(row=tr, column=k, value=f"=SUM({col}6:{col}29)")
        c.number_format = NUM; c.fill = TOTAL_FILL; c.font = SUB_FONT; c.border = BORDER
    ws["A32"] = "Peak month"
    ws["B32"] = "=INDEX(A6:A29,MATCH(MAX(J6:J29),J6:J29,0))"
    ws["B32"].number_format = MONTH
    ws["C32"] = "=MAX(J6:J29)"
    ws["C32"].number_format = NUM
    ws["A33"] = "Average monthly cash out"
    ws["C33"] = "=AVERAGE(J6:J29)"
    ws["C33"].number_format = NUM
    ws["A34"] = ("Tip: overlay against the broiler sales cycle — festive peaks (Hari Raya, CNY, Deepavali) "
                 "lift receipts, while flock placements 6–7 weeks earlier lift feed & DOC outflows.")
    ws["A34"].font = NOTE_FONT
    ws.merge_cells("A34:K34")
    for i, w in enumerate([12, 15, 14, 15, 14, 15, 14, 16, 18, 16, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B6"


# ---------------------------------------------------------------------------
# MFRS7_Risk
# ---------------------------------------------------------------------------
def _mfrs7(wb):
    ws = wb.create_sheet("MFRS7_Risk")
    _title(ws, "MFRS 7 — LIQUIDITY (UNDISCOUNTED CASH FLOWS) & INTEREST-RATE SENSITIVITY", "A1:G1")
    ws["A2"] = ("MFRS 7.39 / B11: the liquidity table must show CONTRACTUAL UNDISCOUNTED cash flows "
                "(principal + future interest/profit) — the FS_Disclosure maturity profile is principal only.")
    ws["A2"].font = NOTE_FONT
    ws.merge_cells("A2:G2")

    _section(ws, 4, "A. LIQUIDITY RISK — contractual undiscounted cash flows by bucket", "G")
    _hdr(ws, 5, ["Bucket", "Maybank Islamic", "Alliance Bank", "MBSB", "TOTAL", "Check"])
    srcs = [("Amortisation_MBB", "F", "J"), ("Amortisation_ALB", "G", "J"), ("Amortisation_MBSB", "F", "I")]
    buckets = [("Within 1 year", "Current"), ("1 to 2 years", "1-2 yrs"),
               ("2 to 5 years", "2-5 yrs"), ("Over 5 years", "Over 5 yrs")]
    for i, (lbl, b) in enumerate(buckets, start=6):
        ws.cell(row=i, column=1, value=lbl).border = BORDER
        for k, (sh, amt, bk) in enumerate(srcs, start=2):
            c = ws.cell(row=i, column=k, value=f'=SUMIFS({sh}!${amt}$9:${amt}$400,{sh}!${bk}$9:${bk}$400,"{b}")')
            c.number_format = NUM; c.border = BORDER
        c = ws.cell(row=i, column=5, value=f"=SUM(B{i}:D{i})")
        c.number_format = NUM; c.border = BORDER; c.fill = CALC_FILL
    ws.cell(row=10, column=1, value="Total contractual cash flows").font = SUB_FONT
    for k in range(2, 6):
        col = get_column_letter(k)
        c = ws.cell(row=10, column=k, value=f"=SUM({col}6:{col}9)")
        c.number_format = NUM; c.fill = TOTAL_FILL; c.font = SUB_FONT
    ws.cell(row=11, column=1, value="Carrying amount (principal at FYE)")
    for k, sh in enumerate(["FS_Disclosure_MBB", "FS_Disclosure_ALB", "FS_Disclosure_MBSB"], start=2):
        ws.cell(row=11, column=k, value=f"={sh}!B7").number_format = NUM
    ws.cell(row=11, column=5, value="=SUM(B11:D11)").number_format = NUM
    ws.cell(row=12, column=1, value="Future interest / profit included above")
    for k in range(2, 6):
        col = get_column_letter(k)
        ws.cell(row=12, column=k, value=f"={col}10-{col}11").number_format = NUM
    ws.cell(row=12, column=6, value='=IF(MIN(B12:D12)<0,"CHECK: cash flows < carrying amount","OK")')

    _section(ws, 14, "B. INTEREST / PROFIT RATE SENSITIVITY — impact on profit before & after tax", "G")
    ws["A15"] = "Corporate tax rate"
    ws["B15"] = 0.24
    ws["B15"].number_format = PCT; ws["B15"].fill = INPUT_FILL
    _hdr(ws, 17, ["Facility", "Carrying amount", "Floating? (Y/N)", "+25 bp", "+50 bp", "+100 bp", "Basis"])
    fac = [("Maybank CMTF-i", "=FS_Disclosure_MBB!B7", "Y", "BFR-i / COF-linked — confirm"),
           ("Alliance Bank FL", "=FS_Disclosure_ALB!B7", "Y", "BLR / BR-linked — confirm"),
           ("MBSB Tawarruq", "=FS_Disclosure_MBSB!B7", "N", "Fixed selling price; if variable w/ ceiling, Ibra moves — confirm"),
           ("Trade lines (OD/BA/RC)", "=Covenant_Monitor!B10", "Y", "Mostly floating")]
    for i, (n, amt, fl, basis) in enumerate(fac, start=18):
        ws.cell(row=i, column=1, value=n).border = BORDER
        c = ws.cell(row=i, column=2, value=amt); c.number_format = NUM; c.border = BORDER
        c = ws.cell(row=i, column=3, value=fl); c.fill = INPUT_FILL; c.alignment = CENTER; c.border = BORDER
        for k, bp in zip((4, 5, 6), (25, 50, 100)):
            c = ws.cell(row=i, column=k, value=f'=IF(UPPER($C{i})="Y",-$B{i}*{bp}/10000,0)')
            c.number_format = NUM; c.border = BORDER
        ws.cell(row=i, column=7, value=basis).font = NOTE_FONT
    dv = DataValidation(type="list", formula1='"Y,N"'); ws.add_data_validation(dv); dv.add("C18:C21")
    ws.cell(row=22, column=1, value="Impact on profit before tax").font = SUB_FONT
    ws.cell(row=23, column=1, value="Impact on profit after tax / equity").font = SUB_FONT
    for k in (4, 5, 6):
        col = get_column_letter(k)
        c = ws.cell(row=22, column=k, value=f"=SUM({col}18:{col}21)")
        c.number_format = NUM; c.fill = TOTAL_FILL
        c = ws.cell(row=23, column=k, value=f"={col}22*(1-$B$15)")
        c.number_format = NUM; c.fill = TOTAL_FILL
    ws["A24"] = "A rate cut gives the mirror-image (favourable) impact. Annualised on year-end balances."
    ws["A24"].font = NOTE_FONT
    for col, w in zip("ABCDEFG", [38, 18, 16, 16, 16, 16, 55]):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# MFRS123_Capitalisation
# ---------------------------------------------------------------------------
def _mfrs123(wb):
    ws = wb.create_sheet("MFRS123_Capitalisation")
    _title(ws, "MFRS 123 — BORROWING COSTS ON AVAILABILITY-PERIOD TRANCHES (CMTF-i 2–5)", "A1:J1")
    ws["A2"] = ("CMTF-i 2–5 carry a 36-month availability (drawdown, profit-only) period — typical of farm "
                "construction financing. Profit on specific borrowings for a QUALIFYING ASSET (farm houses, "
                "hatchery, plant) must be capitalised to CWIP until the asset is ready for use, not expensed. "
                "Does NOT apply to biological assets at fair value (breeder flocks — MFRS 141).")
    ws["A2"].font = NOTE_FONT; ws["A2"].alignment = WRAP
    ws.merge_cells("A2:J2")
    ws.row_dimensions[2].height = 45

    ws["A4"] = "Entity FYE"
    ws["B4"] = '=VLOOKUP("MBB-CMTFi-2",Facility_Master!A:R,18,FALSE)'
    ws["B4"].number_format = DATE; ws["B4"].fill = CALC_FILL
    ws["A5"] = "Financial year start"
    ws["B5"] = "=EDATE(B4,-12)+1"
    ws["B5"].number_format = DATE; ws["B5"].fill = CALC_FILL

    _section(ws, 7, "A. QUALIFYING ASSET PER TRANCHE (inputs)", "J")
    _hdr(ws, 8, ["Facility_ID", "Qualifying asset", "Construction start", "Ready-for-use date", "Profit rate"])
    for i, fid in enumerate(["MBB-CMTFi-2", "MBB-CMTFi-3", "MBB-CMTFi-4", "MBB-CMTFi-5"], start=9):
        ws.cell(row=i, column=1, value=fid).border = BORDER
        for k, v in [(2, "[e.g. Breeder farm house phase — confirm]"), (3, None), (4, None)]:
            c = ws.cell(row=i, column=k, value=v)
            c.fill = INPUT_FILL; c.border = BORDER
            if k in (3, 4):
                c.number_format = DATE
        c = ws.cell(row=i, column=5, value=f"=VLOOKUP(A{i},Facility_Master!A:M,13,FALSE)")
        c.number_format = '0.00%'; c.fill = CALC_FILL; c.border = BORDER
    ws["A13"] = "Leave ready-for-use date blank if still under construction at FYE."
    ws["A13"].font = NOTE_FONT

    _section(ws, 15, "B. CAPITALISABLE PROFIT IN CURRENT FY (per drawdown, from Drawdown_Log)", "J")
    _hdr(ws, 16, ["Facility_ID", "Drawdown date", "Amount (RM)", "Ready-for-use", "Rate",
                  "Window start", "Window end", "Days", "Capitalise (RM)", "Note"])
    for n in range(12):
        r = 17 + n
        src = 2 + n
        ws.cell(row=r, column=1, value=f'=IF(Drawdown_Log!A{src}="","",Drawdown_Log!A{src})')
        ws.cell(row=r, column=2, value=f'=IF(A{r}="","",Drawdown_Log!B{src})').number_format = DATE
        ws.cell(row=r, column=3, value=f'=IF(A{r}="","",Drawdown_Log!C{src})').number_format = NUM
        ws.cell(row=r, column=4,
                value=(f'=IF(A{r}="","",IF(ISNA(MATCH(A{r},$A$9:$A$12,0)),"n/a",'
                       f'IF(INDEX($D$9:$D$12,MATCH(A{r},$A$9:$A$12,0))="","Under construction",'
                       f'INDEX($D$9:$D$12,MATCH(A{r},$A$9:$A$12,0)))))')).number_format = DATE
        ws.cell(row=r, column=5,
                value=f'=IF(OR(A{r}="",D{r}="n/a"),"",INDEX($E$9:$E$12,MATCH(A{r},$A$9:$A$12,0)))'
                ).number_format = '0.00%'
        ws.cell(row=r, column=6, value=f'=IF(E{r}="","",MAX(B{r},$B$5))').number_format = DATE
        ws.cell(row=r, column=7,
                value=f'=IF(E{r}="","",IF(D{r}="Under construction",$B$4,MIN(D{r},$B$4)))'
                ).number_format = DATE
        ws.cell(row=r, column=8, value=f'=IF(E{r}="",0,MAX(0,G{r}-F{r}))')
        c = ws.cell(row=r, column=9, value=f'=IF(E{r}="",0,ROUND(C{r}*E{r}*H{r}/365,2))')
        c.number_format = NUM; c.fill = CALC_FILL
        ws.cell(row=r, column=10,
                value=(f'=IF(D{r}="Under construction","Ready date blank — assumed under construction to FYE",'
                       f'IF(D{r}="n/a","Not an availability-period tranche",""))')).font = NOTE_FONT
        for k in range(1, 11):
            ws.cell(row=r, column=k).border = BORDER
    ws["H29"] = "TOTAL"
    ws["H29"].font = SUB_FONT
    ws["I29"] = "=SUM(I17:I28)"
    ws["I29"].number_format = NUM; ws["I29"].fill = TOTAL_FILL; ws["I29"].font = SUB_FONT

    _section(ws, 31, "C. SUGGESTED JOURNAL (if capitalisation criteria met)", "J")
    ws["A32"] = "Dr  Property, plant & equipment — Capital work-in-progress"
    ws["I32"] = "=I29"
    ws["A33"] = "    Cr  Finance cost — profit on Islamic financing (P&L)"
    ws["I33"] = "=-I29"
    for c in ("I32", "I33"):
        ws[c].number_format = NUM
    ws["A34"] = ("Assumptions: specific borrowing, fully applied to the asset, no temporary investment income, "
                 "balance = drawn amount (profit-only during availability). Suspend during extended "
                 "inactive periods (MFRS 123.20); cease when ready for use (MFRS 123.22). "
                 "Tax: capitalised profit is generally NOT deductible under s.33(1) ITA but may qualify for "
                 "capital allowances as part of qualifying expenditure — confirm with tax agent.")
    ws["A34"].font = NOTE_FONT; ws["A34"].alignment = WRAP
    ws.merge_cells("A34:J34")
    ws.row_dimensions[34].height = 55
    for i, w in enumerate([16, 36, 16, 18, 9, 13, 13, 7, 16, 48], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# Blind_Spots
# ---------------------------------------------------------------------------
BLIND_SPOTS = [
    ("High", "Template integrity",
     "v1–v3: (a) dates stored as TEXT, so current / non-current split and maturity profile computed as ZERO; "
     "(b) whole-column SUMIFS picked up header cells, so MBSB principal at FYE was overstated by the RM1.23m Ibra memo.",
     "Both fixed in v4. Re-check any FS figures or bank submissions produced from v3.",
     "All FS_Disclosure sheets"),
    ("High", "Covenants (MFRS 101.69–76ZA)",
     "A covenant breach at reporting date makes the WHOLE facility current unless a waiver is obtained "
     "on/before that date. 2024 amendments also require disclosure of covenants for non-current loans.",
     "Key in each facility letter's covenants; obtain waivers BEFORE FYE; disclose covenant terms.",
     "Covenant_Monitor"),
    ("High", "Facility_Master — Maybank CMTF-i 1",
     "Schedule opens at the RM21m LIMIT, not the RM1.08m drawn per Drawdown_Log; RM144,660 EPI does not "
     "amortise RM21m in 180 months → large balloon in final period. Outstanding may be overstated.",
     "Confirm whether RM21m is the umbrella limit for tranches 1–5; set opening principal = drawn amount.",
     "Facility_Master / Amortisation_MBB"),
    ("High", "Group consolidation — FYE mismatch",
     "Maybank borrower FYE 31-Mar-2026 vs Alliance / MBSB 31-Dec-2025, yet FS_Aggregate adds them. "
     "MFRS 10.B92–B93: max 3-month gap with adjustments for significant transactions.",
     "Re-run each schedule at the GROUP reporting date before aggregating.",
     "FS_Aggregate"),
    ("High", "Feed-cost exposure vs debt",
     "Corn & SBM are USD-priced imports; a combined commodity + ringgit move can erode DSCR within a quarter.",
     "Track break-even corn / SBM / USD-MYR monthly; consider forward FX cover & hedge accounting (MFRS 9).",
     "Feed_Stress"),
    ("Medium", "Borrowing costs (MFRS 123)",
     "Profit on CMTF-i 2–5 during the 36-month availability period may relate to farm construction and "
     "should be capitalised to CWIP, not expensed.",
     "Identify qualifying assets & ready-for-use dates; book capitalisation journal.",
     "MFRS123_Capitalisation"),
    ("Medium", "MFRS 7 liquidity table",
     "Maturity analysis must show undiscounted contractual cash flows (principal + interest), "
     "not principal only.", "Use the MFRS7_Risk table for the note.", "MFRS7_Risk"),
    ("Medium", "MFRS 7 rate sensitivity",
     "Floating-rate facilities (BFR/BLR-linked) need a sensitivity disclosure; OPR moves hit P&L directly.",
     "Confirm floating/fixed per facility; disclose ±bp impact.", "MFRS7_Risk"),
    ("Medium", "Trade lines & leases",
     "OD / BA / RC / trust receipts for feed imports, HP and MFRS 16 leases are outside the schedules "
     "but count in bank gearing definitions.", "Add them to Covenant_Monitor row 10 and the calendar.",
     "Covenant_Monitor"),
    ("Medium", "Holding co. on-lending (tax)",
     "MBSB is borrowed at Holdings level. If funds are on-lent interest-free or invested in non-income-"
     "producing assets, s.33(2) ITA restricts the profit deduction; intercompany loans fall under the "
     "Transfer Pricing Rules (arm's-length rate).",
     "Document use of funds; charge arm's-length intercompany profit; confirm with tax agent.", "—"),
    ("Medium", "Corporate guarantees",
     "Guarantees by Holdings for subsidiaries' facilities are financial guarantee contracts (MFRS 9) "
     "in the company-level FS — measure & assess ECL.", "List all guarantees; assess fair value / ECL.", "—"),
    ("Medium", "Contract-farmer payments — e-Invoicing",
     "Payments to individual contract farmers who don't issue e-invoices may require SELF-BILLED "
     "e-invoices under LHDN's e-Invoice guidelines.",
     "Confirm farmer status & thresholds with tax agent; build self-billing into the settlement run.", "—"),
    ("Low", "Islamic financing presentation",
     "Disclose Ibra, Ta'widh/Gharamah (late charges), and present Islamic vs conventional borrowings "
     "separately where material.", "Add to borrowings note.", "FS_Disclosure_MBSB"),
    ("Low", "Biological assets in covenant EBITDA",
     "MFRS 141 fair-value gains on breeders / broilers inflate EBITDA & equity; many banks exclude them.",
     "Check definition in each facility letter.", "Covenant_Monitor"),
]


def _blind_spots(wb):
    ws = wb.create_sheet("Blind_Spots", index=1)
    _title(ws, "BLIND-SPOT CHECKLIST — items commonly missed on group borrowings (review before each FYE)", "A1:H1")
    _hdr(ws, 3, ["#", "Priority", "Area", "What could be missed", "Action", "Toolkit sheet", "Owner", "Status"])
    for i, (pri, area, what, action, sheet) in enumerate(BLIND_SPOTS, start=1):
        r = 3 + i
        vals = [i, pri, area, what, action, sheet, None, "Open"]
        for k, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=k, value=v)
            c.border = BORDER; c.alignment = WRAP
            if k in (7, 8):
                c.fill = INPUT_FILL
        ws.row_dimensions[r].height = 62
    last = 3 + len(BLIND_SPOTS)
    dv = DataValidation(type="list", formula1='"Open,In progress,Done,N/A"')
    ws.add_data_validation(dv); dv.add(f"H4:H{last}")
    ws.conditional_formatting.add(f"B4:B{last}", CellIsRule(operator="equal", formula=['"High"'], fill=RED_FILL))
    ws.conditional_formatting.add(f"B4:B{last}", CellIsRule(operator="equal", formula=['"Medium"'], fill=AMBER_FILL))
    ws.conditional_formatting.add(f"H4:H{last}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN_FILL))
    ws.cell(row=last + 2, column=3, value="Open items").font = SUB_FONT
    ws.cell(row=last + 2, column=4, value=f'=COUNTIF(H4:H{last},"Open")+COUNTIF(H4:H{last},"In progress")')
    ws.cell(row=last + 3, column=4,
            value="Tax / regulatory points are flags for discussion with your tax agent, not conclusions.").font = NOTE_FONT
    for col, w in zip("ABCDEFGH", [4, 9, 24, 60, 48, 22, 12, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"


def _change_log(wb):
    ws = wb.create_sheet("Change_Log")
    _title(ws, "CHANGE LOG", "A1:C1")
    rows = [
        ("v1", "Maybank Islamic CMTF-i template"),
        ("v2", "Added Alliance Bank Fixed Loan"),
        ("v3", "Added MBSB Tawarruq / Murabahah"),
        ("v4", "FIX: text dates → real dates (bucket / current-portion formulas were returning zero)"),
        ("v4", "FIX: SUMIFS ranges bounded to schedule rows (MBSB principal was overstated by header-cell Ibra)"),
        ("v4", "NEW: Blind_Spots, Covenant_Monitor, Feed_Stress, Debt_Service_Calendar, MFRS7_Risk, "
               "MFRS123_Capitalisation"),
    ]
    _hdr(ws, 3, ["Version", "Change"])
    for i, (v, t) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=v).border = BORDER
        ws.cell(row=i, column=2, value=t).border = BORDER
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 110


def add_toolkit_sheets(wb):
    cr = _covenant_monitor(wb)
    _feed_stress(wb, cr)
    _calendar(wb)
    _mfrs7(wb)
    _mfrs123(wb)
    _blind_spots(wb)
    _change_log(wb)
    # README pointer
    ws = wb["README"]
    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="v4 DEBT HEADROOM & COMPLIANCE TOOLKIT").font = SUB_FONT
    ws.cell(row=r, column=1).fill = SUB_FILL
    for k, t in enumerate([
        "  Blind_Spots            — prioritised checklist of commonly missed items (start here)",
        "  Covenant_Monitor       — DSCR / gearing / TNW headroom, MFRS 101 reclassification flag",
        "  Feed_Stress            — corn / SBM / USD-MYR shocks → DSCR heat-map & break-even prices",
        "  Debt_Service_Calendar  — 24-month principal + interest/profit cash calendar",
        "  MFRS7_Risk             — undiscounted liquidity table + rate sensitivity",
        "  MFRS123_Capitalisation — profit to capitalise on availability-period tranches",
        "  YELLOW inputs in these sheets are ILLUSTRATIVE placeholders — replace before use.",
    ], start=1):
        ws.cell(row=r + k, column=1, value=t)
