"""
Work & Personal Tracker v1 — builder
Consolidates open items from all Claude Code sessions (May–Nov 2026),
a recurring Malaysian compliance / close schedule, blind spots, and the
SharePoint reorganisation plan (pre-access).
"""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

OUT = "/home/user/Claude-code-learning/Work_Personal_Tracker_v1.xlsx"
F = "Arial"
HDR = Font(name=F, bold=True, color="FFFFFF", size=10)
HFILL = PatternFill("solid", fgColor="1F4E78")
TITLE = Font(name=F, bold=True, size=14, color="1F4E78")
BOLD = Font(name=F, bold=True, size=10)
NORM = Font(name=F, size=10)
BLUE = Font(name=F, size=10, color="0000FF")
INPUT = PatternFill("solid", fgColor="FFF2CC")
CALC = PatternFill("solid", fgColor="E2EFDA")
SUB = PatternFill("solid", fgColor="D9E1F2")
TH = Side(border_style="thin", color="BFBFBF")
BRD = Border(left=TH, right=TH, top=TH, bottom=TH)
WRAP = Alignment(wrap_text=True, vertical="top")
DFMT = "dd-mmm-yyyy"


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c)
        x.font, x.fill, x.border = HDR, HFILL, BRD
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def cell(ws, r, c, v, font=NORM, fill=None, fmt=None):
    x = ws.cell(row=r, column=c, value=v)
    x.font, x.border, x.alignment = font, BRD, WRAP
    if fill:
        x.fill = fill
    if fmt:
        x.number_format = fmt
    return x


wb = Workbook()

# ============================================================ README
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 110
lines = [
    ("Work & Personal Tracker — v1 (built 04-Nov-2026)", TITLE),
    ("", NORM),
    ("PURPOSE", BOLD),
    ("One place to track every open item from our Claude sessions (May–Nov 2026), recurring Malaysian", NORM),
    ("compliance / month-end deadlines, blind spots, and the SharePoint reorganisation plan.", NORM),
    ("", NORM),
    ("SHEETS", BOLD),
    ("  1. Dashboard       — live counts by area / status / RAG, overdue and next-14-day items", NORM),
    ("  2. Tracker         — master action list (one row per item). Update Status + Due Date weekly.", NORM),
    ("  3. Schedule        — recurring calendar; 'Next Due' recalculates from TODAY() automatically", NORM),
    ("  4. Blind_Spots     — things likely to be overlooked, with suggested action and linked Tracker ID", NORM),
    ("  5. SharePoint_Plan — phased approach + draft target structure; nothing executes until you approve", NORM),
    ("  6. Sources         — which past session each item came from", NORM),
    ("", NORM),
    ("COLOUR LEGEND", BOLD),
    ("  YELLOW fill = you edit (status, due date, owner, notes)", NORM),
    ("  GREEN fill  = formula — do not overwrite", NORM),
    ("  Blue text   = date / figure I assumed — verify before relying on it", NORM),
    ("", NORM),
    ("HOW TO RUN IT (suggested rhythm)", BOLD),
    ("  Monday 15 min : filter Tracker on RAG = Overdue / Due ≤14d; update Status; add new items at the bottom.", NORM),
    ("  Month-end     : walk the Schedule sheet top to bottom (close WD1–WD7, statutory 15th payments).", NORM),
    ("  Quarter-end   : refresh Blind_Spots — close what's resolved, add what's new.", NORM),
    ("", NORM),
    ("IMPORTANT CAVEATS", BOLD),
    ("  • Statutory dates reflect general Malaysian rules (LHDN, KWSP, PERKESO, HRD Corp, SSM, Bursa). Confirm", NORM),
    ("    each entity's FYE and any extension/grace period with your tax agent / company secretary.", NORM),
    ("  • Items marked 'To verify' were inferred, not confirmed — e.g. the property's VP date, entity FYEs.", NORM),
    ("  • Tax points are prompts for your tax agent, not conclusions.", NORM),
]
for i, (t, f) in enumerate(lines, 1):
    ws.cell(row=i, column=1, value=t).font = f

# ============================================================ TRACKER
tr = wb.create_sheet("Tracker")
tr["A1"] = "Master Tracker"; tr["A1"].font = TITLE
tr["A2"] = "Today:"; tr["A2"].font = BOLD
tr["B2"] = "=TODAY()"; tr["B2"].number_format = DFMT; tr["B2"].fill = CALC; tr["B2"].font = NORM
cols = ["ID", "Area", "Stream", "Item", "Next action", "Priority", "Owner", "Status",
        "Due date", "Days to due", "RAG", "Source", "Notes"]
header(tr, 4, cols, [7, 9, 18, 40, 48, 8, 10, 12, 12, 9, 11, 16, 36])
tr.freeze_panes = "E5"

T = "Today"  # owner shorthand
items = [
    # ---------- WORK: Loan / treasury ----------
    ("W01", "Work", "Loans & covenants", "Recheck any figures taken from Loan Template v3",
     "v3 had text-dates (current/non-current = 0) and MBSB overstated RM3.47m vs RM2.24m. Recheck any draft FS / bank submission using v3.",
     "P1", "Me", "Not started", date(2026, 11, 10), "Budget session (2-Nov)", "v4 on branch claude/busy-curie-jg6597 has the fix"),
    ("W02", "Work", "Loans & covenants", "Fix Maybank CMTF-i 1 opening balance",
     "Confirm whether RM21m is the combined limit for CMTF-i 1–5; ~RM1.08m drawn. Current schedule implies RM8.26m balloon in 2039.",
     "P1", "Me", "Not started", date(2026, 11, 13), "Budget session (2-Nov)", ""),
    ("W03", "Work", "Loans & covenants", "Load actual covenant terms into Covenant_Monitor (v4)",
     "From each facility letter: DSCR, gearing, TNW thresholds, test dates, cross-default clauses. Replace placeholders.",
     "P1", "Me", "Not started", date(2026, 11, 20), "Budget session (2-Nov)", "Feed_Stress showed breach at +5.3% corn/SBM on placeholder inputs"),
    ("W04", "Work", "Loans & covenants", "Complete CMTF-i 2–5 tranche data",
     "Facility limits, first instalment dates, last statement balance per tranche.", "P2", "Me", "Not started",
     date(2026, 11, 20), "Loan template session (May)", ""),
    ("W05", "Work", "Loans & covenants", "Confirm MBSB borrower entity, FYE and first repayment date",
     "Placeholder = Farmiera Holdings, FYE 31-Dec-2025, first repayment Dec-24 (sample schedule showed 01-Apr-2024).",
     "P2", "Me", "Not started", date(2026, 11, 20), "Loan template session (May)", ""),
    ("W06", "Work", "Loans & covenants", "Add trade lines to borrowings model",
     "OD, BA, RC, trust receipts (feed imports), hire purchase, MFRS 16 leases — they count in bank gearing definitions.",
     "P2", "Me", "Not started", date(2026, 11, 27), "Budget session (2-Nov)", ""),
    ("W07", "Work", "Loans & covenants", "Resolve FYE mismatch in FS_Aggregate",
     "Breeder Farm FYE 31-Mar vs Agro/Holdings 31-Dec — align periods before consolidating.",
     "P2", "Me", "Not started", date(2026, 11, 27), "Budget session (2-Nov)", ""),
    ("W08", "Work", "Accounting / MFRS", "MFRS 123 borrowing-cost capitalisation (CMTF-i 2–5)",
     "Confirm construction assets + ready-for-use dates; agree capitalisation with auditor before FYE 31-Dec-2026.",
     "P1", "Me", "Not started", date(2026, 12, 4), "Budget session (2-Nov)", "Illustrative ≈RM600k"),
    ("W09", "Work", "Accounting / MFRS", "Add COA mapping sheet to loan template",
     "Share loan-related GL codes (Term Loan–Maybank, Murabahah Payable–MBSB, etc.).", "P3", "Me", "Not started",
     date(2026, 12, 11), "Loan template session (May)", ""),
    # ---------- WORK: Tax ----------
    ("W10", "Work", "Tax", "Tax agent review: s.33(2), transfer pricing, corporate guarantees",
     "MBSB loan sits at Holdings — interest restriction on on-lending, TP documentation, guarantee fees.",
     "P2", "Tax agent", "Not started", date(2026, 11, 30), "Budget session (2-Nov)", ""),
    ("W11", "Work", "Tax", "Self-billed e-invoices for contract-farmer payments",
     "Confirm each entity's e-Invoice phase/relaxation status and set up self-billed e-invoice flow for farmer settlements.",
     "P1", "Me", "Not started", date(2026, 11, 30), "Budget session (2-Nov)", "To verify current LHDN phase dates"),
    ("W12", "Work", "Tax", "CP204A revision — FYE 31-Dec-2026 entities (11th month)",
     "November is the last revision window for YA2026 estimate; avoid >30% under-estimate penalty.",
     "P1", "Me", "Not started", date(2026, 11, 30), "Blind spot", "To verify with tax agent"),
    ("W13", "Work", "Tax", "CP204 YA2027 estimate — FYE 31-Dec entities",
     "Submit ≥30 days before basis period starts (1-Jan-2027).", "P1", "Me", "Not started",
     date(2026, 12, 1), "Blind spot", "To verify"),
    ("W14", "Work", "Statutory", "Confirm Breeder Farm (FYE 31-Mar-2026) Form C + SSM lodgement done",
     "Form C due 7 months after FYE (31-Oct-2026); FS circulated ≤6 months, lodged with SSM ≤30 days after.",
     "P1", "Co Sec / Tax agent", "Not started", date(2026, 11, 6), "Blind spot", "If not filed — check grace period immediately"),
    # ---------- WORK: Automation projects ----------
    ("W15", "Work", "Automation", "Utility bill importer — go-live tuning",
     "Set real SQL Account GL + tax codes; run --inspect on real TNB/Air Selangor/TM/Maxis bills; tune regex.",
     "P2", "Me", "Not started", date(2026, 12, 11), "Utility importer session (Jun)", "Branch claude/awesome-hawking-eCHT8"),
    ("W16", "Work", "Automation", "Poultry farm tracker — pilot on one farm",
     "Decide hosting; add CSV export, mobile daily-entry form, weight-drift alerts.", "P3", "Me", "Not started",
     date(2027, 1, 15), "Poultry tracker session (May)", "Branch claude/poultry-farm-tracker-Jwqv2"),
    ("W17", "Work", "Automation", "Build contract-farmer settlement statement generator",
     "FCR, mortality, grading → payment; include e-invoice fields (links to W11).", "P2", "Me", "Not started",
     date(2026, 12, 18), "Budget session (2-Nov)", "Recommended next build"),
    ("W18", "Work", "Automation", "Consolidate GitHub branches",
     "5 unmerged branches, no 'main'. Merge finished work (v4 loan template) into one default branch; archive the rest.",
     "P3", "Me", "Not started", date(2026, 11, 20), "Blind spot", ""),
    # ---------- WORK: Market / peers ----------
    ("W19", "Work", "Market & peers", "Set up monthly market update",
     "Corn (CBOT), soybean meal, USD/MYR, local DOC/broiler prices, policy changes. Feed into Feed_Stress sheet.",
     "P2", "Me", "Not started", date(2026, 11, 30), "Your stated Pillar 1 — not yet built", ""),
    ("W20", "Work", "Market & peers", "Peer benchmarking — Sep-2026 quarter results",
     "Leong Hup (FYE Dec, Q3), QL & Lay Hong (FYE Mar, Q2 FY27). Bursa deadline ≈30-Nov-2026.",
     "P2", "Me", "Not started", date(2026, 12, 5), "Your stated Pillar 1 — not yet built", ""),
    ("W21", "Work", "Market & peers", "Review Budget 2027 measures affecting poultry/agri",
     "Subsidies, tax incentives, SST/e-invoice changes, foreign-worker levy.", "P3", "Me", "Not started",
     date(2026, 11, 20), "Blind spot", "To verify"),
    # ---------- WORK: FYE 2026 ----------
    ("W22", "Work", "Year-end", "Plan FYE 31-Dec-2026 stock counts",
     "Live birds, feed, vaccines/medication, eggs. Agree count dates and auditor attendance.", "P1", "Me", "Not started",
     date(2026, 12, 1), "Blind spot", ""),
    ("W23", "Work", "SharePoint", "SharePoint reorganisation — grant access & run inventory",
     "See SharePoint_Plan sheet. Phase 0: grant access / export folder inventory.", "P2", "Me", "Waiting",
     date(2026, 11, 18), "This session", "No SharePoint connector attached yet"),
    # ---------- PERSONAL ----------
    ("P01", "Personal", "Admin", "Use remaining Claude credit before expiry",
     "Credit noted as expiring 5-Nov-2026 — use it for W03 / W17 today.", "P1", "Me", "Not started",
     date(2026, 11, 5), "Budget session (2-Nov)", "Tomorrow"),
    ("P02", "Personal", "Property B-43-3A", "Confirm EXSIM handover package items",
     "Mattresses, TV, fridge/washer, ceiling lights — included or not? Gates final fit-out budget.", "P1", "Me",
     "Not started", date(2026, 11, 13), "Interior design session (Nov)", "Tenant-ready RM21,140; cap RM50,000"),
    ("P03", "Personal", "Property B-43-3A", "Decide scheme + build BOQ",
     "Teal/blond-ash scheme vs earlier clay scheme; BOQ with SKUs, store links, running total vs RM50k.", "P2", "Me",
     "Not started", date(2026, 11, 20), "Interior design session (Nov)", "Reference-look hero items ≈RM2,600"),
    ("P04", "Personal", "Property B-43-3A", "Defect inspection & DLP log",
     "Record defects with photos and submit before Defect Liability Period expires (typically 24 months from VP).",
     "P1", "Me", "Not started", date(2026, 11, 30), "Blind spot", "Enter VP date in Notes"),
    ("P05", "Personal", "Property B-43-3A", "Landlord setup",
     "Tenancy agreement (stamp duty within 30 days of signing), landlord/fire insurance, deposit terms, tenant screening.",
     "P2", "Me", "Not started", date(2026, 12, 15), "Blind spot", ""),
    ("P06", "Personal", "Investments", "Webull — decide on official OpenAPI scope",
     "Market data vs account read vs paper trading; apply for app key if proceeding.", "P3", "Me", "Not started",
     date(2026, 12, 31), "Webull session (May)", ""),
    ("P07", "Personal", "Investments", "US-situs asset exposure check",
     "Non-resident aliens face US estate tax above USD60k of US assets; check W-8BEN validity (3-year renewal).",
     "P2", "Me", "Not started", date(2026, 12, 15), "Blind spot", "Discuss with adviser"),
    ("P08", "Personal", "Tax", "Prepare YA2026 personal tax file",
     "Rental income/expenses, foreign-sourced income, reliefs (EPF, insurance, lifestyle, medical).", "P2", "Me",
     "Not started", date(2027, 3, 31), "Blind spot", "Form BE due 30-Apr-2027 (e-filing usually +15 days)"),
    ("P09", "Personal", "Admin", "Close World Cup tracker PR #1",
     "Tournament finished Jul-2026 — merge or close PR #1 and archive branch.", "P3", "Me", "Not started",
     date(2026, 11, 20), "World Cup session (May)", ""),
]
r0 = 5
for i, it in enumerate(items):
    r = r0 + i
    (iid, area, stream, item, nxt, pri, own, st, due, src, notes) = it
    vals = [iid, area, stream, item, nxt, pri, own, st, due, None, None, src, notes]
    for c, v in enumerate(vals, 1):
        f = INPUT if c in (6, 7, 8, 9, 13) else None
        cell(tr, r, c, v, fill=f, fmt=DFMT if c == 9 else None)
    cell(tr, r, 10, f'=IF(OR(I{r}="",H{r}="Done"),"",I{r}-TODAY())', fill=CALC, fmt="0")
    cell(tr, r, 11, f'=IF(H{r}="Done","Done",IF(I{r}="","No date",IF(I{r}<TODAY(),"Overdue",'
                    f'IF(I{r}-TODAY()<=14,"Due ≤14d","On track"))))', fill=CALC)
last = r0 + len(items) - 1
# blank capacity rows with formulas
for r in range(last + 1, 151):
    for c in range(1, 14):
        f = INPUT if c in (1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13) else None
        cell(tr, r, c, None, fill=f, fmt=DFMT if c == 9 else None)
    cell(tr, r, 10, f'=IF(OR(I{r}="",H{r}="Done"),"",I{r}-TODAY())', fill=CALC, fmt="0")
    cell(tr, r, 11, f'=IF(D{r}="","",IF(H{r}="Done","Done",IF(I{r}="","No date",IF(I{r}<TODAY(),"Overdue",'
                    f'IF(I{r}-TODAY()<=14,"Due ≤14d","On track")))))', fill=CALC)

for rng, lst in [("B5:B150", '"Work,Personal"'), ("F5:F150", '"P1,P2,P3"'),
                 ("H5:H150", '"Not started,In progress,Waiting,Done,Dropped"')]:
    dv = DataValidation(type="list", formula1=lst, allow_blank=True)
    dv.add(rng); tr.add_data_validation(dv)
tr.auto_filter.ref = "A4:M150"
red = PatternFill("solid", fgColor="F8CBAD"); amb = PatternFill("solid", fgColor="FFE699")
grn = PatternFill("solid", fgColor="C6EFCE"); gry = PatternFill("solid", fgColor="D9D9D9")
tr.conditional_formatting.add("K5:K150", FormulaRule(formula=['K5="Overdue"'], fill=red))
tr.conditional_formatting.add("K5:K150", FormulaRule(formula=['K5="Due ≤14d"'], fill=amb))
tr.conditional_formatting.add("K5:K150", FormulaRule(formula=['K5="On track"'], fill=grn))
tr.conditional_formatting.add("A5:M150", FormulaRule(formula=['$H5="Done"'], fill=gry, font=Font(name=F, color="808080", strike=True)))

# ============================================================ SCHEDULE
sc = wb.create_sheet("Schedule")
sc["A1"] = "Recurring Schedule"; sc["A1"].font = TITLE
sc["A2"] = "Next Due recalculates from TODAY(). Month-end working days use the Holidays list (col J) — add MY public holidays."
sc["A2"].font = NORM
header(sc, 4, ["#", "Area", "Category", "Task", "Frequency", "Rule", "Next due", "Days to go", "Notes"],
       [5, 9, 16, 44, 11, 30, 13, 9, 46])
sc.column_dimensions["J"].width = 13
x = sc.cell(row=4, column=10, value="Holidays"); x.font, x.fill, x.border = HDR, HFILL, BRD
for r in range(5, 35):
    cell(sc, r, 10, None, fill=INPUT, fmt=DFMT)
HOL = "$J$5:$J$34"


def wd(n):  # nth working day of current month, else next month
    a = f"WORKDAY(EOMONTH(TODAY(),-1),{n},{HOL})"
    b = f"WORKDAY(EOMONTH(TODAY(),0),{n},{HOL})"
    return f"=IF({a}>=TODAY(),{a},{b})"


def monthly(d):
    return (f"=IF(DAY(TODAY())<={d},DATE(YEAR(TODAY()),MONTH(TODAY()),{d}),"
            f"DATE(YEAR(TODAY()),MONTH(TODAY())+1,{d}))")


def weekly(wdn):  # WEEKDAY: Sun=1..Sat=7
    return f"=TODAY()+MOD({wdn}-WEEKDAY(TODAY()),7)"


def annual(m, d):
    return f"=DATE(YEAR(TODAY())+(TODAY()>DATE(YEAR(TODAY()),{m},{d})),{m},{d})"


EOM = "=EOMONTH(TODAY(),0)"
QBURSA = "=EOMONTH(TODAY(),MOD(2-MONTH(TODAY()),3))"

sched = [
    ("Work", "Rhythm", "Weekly tracker review (filter Overdue / Due ≤14d)", "Weekly", "Every Monday", weekly(2), ""),
    ("Work", "Market", "Feed & FX check: CBOT corn, SBM, USD/MYR → Feed_Stress", "Weekly", "Every Friday", weekly(6),
     "Compare to break-even prices in loan template v4"),
    ("Work", "Close", "WD1 — cut-off: AP/AR, bank downloads, farm daily records locked", "Monthly", "1st working day", wd(1), ""),
    ("Work", "Close", "WD2 — bank recs, intercompany recs, accruals (utilities, feed, payroll)", "Monthly", "2nd working day", wd(2), "Utility importer output feeds accruals"),
    ("Work", "Close", "WD3 — inventory: feed, live birds (biological assets), eggs, medication", "Monthly", "3rd working day", wd(3), "MFRS 141 fair value / cost"),
    ("Work", "Close", "WD4 — farmer settlements, loan interest/profit accruals, depreciation", "Monthly", "4th working day", wd(4), "Loan template Amortisation sheets"),
    ("Work", "Close", "WD5 — draft P&L/BS, variance analysis, COGS per bird / per kg", "Monthly", "5th working day", wd(5), ""),
    ("Work", "Close", "WD7 — management pack + covenant monitor update", "Monthly", "7th working day", wd(7), "DSCR / gearing / TNW headroom"),
    ("Work", "Statutory", "EPF (KWSP), SOCSO (PERKESO), EIS, PCB/MTD", "Monthly", "15th of following month", monthly(15), "Late payment charges apply"),
    ("Work", "Statutory", "HRD Corp levy", "Monthly", "15th of following month", monthly(15), ""),
    ("Work", "Tax", "CP204 monthly instalment (each entity)", "Monthly", "15th of each month", monthly(15), ""),
    ("Work", "Tax", "Withholding tax (s.107A / s.109) on non-resident payments", "Per payment", "Within 1 month of paying/crediting", None, "Enter date manually when relevant"),
    ("Work", "Tax", "SST-02 return (if registered)", "Bi-monthly", "End of month after taxable period", None, "Enter per your taxable period"),
    ("Work", "Treasury", "Loan repayments — Maybank / Alliance / MBSB", "Monthly", "Per facility (see Debt_Service_Calendar)", None, "Enter actual debit dates"),
    ("Work", "Market", "Monthly market & policy update memo", "Monthly", "Last day of month", EOM, "Feed prices, DOC & broiler prices, subsidies, regulation"),
    ("Work", "Peers", "Bursa quarterly results — Leong Hup, QL, Lay Hong", "Quarterly", "≈2 months after quarter-end", QBURSA, "Benchmark margins, FCR commentary, capex, gearing"),
    ("Work", "Tax", "CP204A revision window (FYE Dec entities: Jun/Sep/Nov)", "Annual", "11th month = November", annual(11, 30), "Breeder Farm (FYE Mar): Sep/Dec/Feb"),
    ("Work", "Tax", "CP204 next-YA estimate (FYE Dec entities)", "Annual", "≥30 days before basis period", annual(12, 1), "To verify"),
    ("Work", "Statutory", "Form C — FYE 31-Dec entities", "Annual", "7 months after FYE", annual(7, 31), ""),
    ("Work", "Statutory", "Form C — Breeder Farm (FYE 31-Mar)", "Annual", "7 months after FYE", annual(10, 31), ""),
    ("Work", "Statutory", "Form E + EA forms to employees", "Annual", "Form E by 31-Mar; EA by end-Feb", annual(2, 28), ""),
    ("Work", "Year-end", "FYE stock count (Dec entities)", "Annual", "31-Dec", annual(12, 31), "Auditor attendance"),
    ("Personal", "Tax", "Personal income tax — Form BE", "Annual", "30-Apr (e-filing usually +15 days)", annual(4, 30), "Use Form B instead if you have business income (30-Jun)"),
    ("Personal", "Property", "Quit rent (cukai tanah) — B-43-3A", "Annual", "Typically by 31-May", annual(5, 31), "To verify with land office"),
    ("Personal", "Property", "Assessment tax (cukai taksiran) — 1st half", "Half-yearly", "Typically by end-Feb", annual(2, 28), "To verify with local council"),
    ("Personal", "Property", "Assessment tax (cukai taksiran) — 2nd half", "Half-yearly", "Typically by end-Aug", annual(8, 31), "To verify with local council"),
    ("Personal", "Property", "Maintenance fee + sinking fund", "Monthly", "Per JMB/MC billing", None, "Enter billing date"),
    ("Personal", "Property", "Landlord / fire insurance renewal", "Annual", "Policy anniversary", None, "Enter date once bought"),
    ("Personal", "Investments", "Portfolio review + W-8BEN expiry check", "Quarterly", "Quarter-end", "=EOMONTH(TODAY(),MOD(3-MONTH(TODAY()),3))", ""),
]
for i, (area, cat, task, freq, rule, f, note) in enumerate(sched, 1):
    r = 4 + i
    cell(sc, r, 1, i); cell(sc, r, 2, area); cell(sc, r, 3, cat); cell(sc, r, 4, task)
    cell(sc, r, 5, freq); cell(sc, r, 6, rule)
    if f:
        cell(sc, r, 7, f, fill=CALC, fmt=DFMT)
    else:
        cell(sc, r, 7, None, fill=INPUT, fmt=DFMT)
    cell(sc, r, 8, f'=IF(G{r}="","",G{r}-TODAY())', fill=CALC, fmt="0")
    cell(sc, r, 9, note, font=BLUE if "verify" in note.lower() else NORM)
sc_last = 4 + len(sched)
sc.freeze_panes = "E5"
sc.auto_filter.ref = f"A4:I{sc_last}"
sc.conditional_formatting.add(f"H5:H{sc_last}", FormulaRule(formula=['AND(H5<>"",H5<=7)'], fill=amb))

# ============================================================ BLIND SPOTS
bs = wb.create_sheet("Blind_Spots")
bs["A1"] = "Things you may overlook"; bs["A1"].font = TITLE
header(bs, 3, ["#", "Area", "Blind spot", "Why it matters", "Suggested action", "Tracker ID", "Basis"],
       [5, 9, 34, 50, 46, 10, 20])
spots = [
    ("Work", "v3 loan figures may already be in use", "Text-dates zeroed the current/non-current split; MBSB overstated by RM1.23m. Errors flow into FS notes and bank submissions.", "Trace every place v3 numbers were used; switch to v4.", "W01", "Found in session"),
    ("Work", "Covenant breach → whole loan current (MFRS 101)", "Unless the waiver is dated on/before reporting date. Cross-default can pull other banks in.", "Test covenants monthly, not just at year-end; get waivers before 31-Dec.", "W03", "Found in session"),
    ("Work", "Feed price sensitivity is tight", "On placeholder inputs, +5.3% corn/SBM breaches DSCR. Real headroom unknown until inputs are actual.", "Load actuals, then watch weekly against break-even prices.", "W03", "Illustrative"),
    ("Work", "Borrowing costs during construction", "CMTF-i 2–5 profit during build may need capitalising (MFRS 123) — affects P&L and tax.", "Agree treatment with auditor before FYE.", "W08", "Found in session"),
    ("Work", "Trade facilities missing from borrowings model", "OD/BA/TR/HP/leases count in bank gearing; omitting them overstates headroom.", "Add to Facility_Master.", "W06", "Found in session"),
    ("Work", "Self-billed e-invoices for contract farmers", "Payments to individual farmers typically require self-billed e-invoices; settlement statements must carry the fields.", "Confirm phase; design into settlement generator.", "W11", "To verify"),
    ("Work", "CP204A — November is the last revision window", "For FYE-Dec entities; under-estimate >30% attracts penalty on the difference.", "Revise with Oct management accounts.", "W12", "Statutory rule"),
    ("Work", "Breeder Farm FYE-Mar filings", "Form C and SSM FS lodgement for FYE 31-Mar-2026 fell due in Oct-2026.", "Confirm filed with tax agent / co sec this week.", "W14", "To verify"),
    ("Work", "Your Pillar 1 (market & peers) has no system yet", "No session has built feed-price or Bursa peer tracking; cost decisions lack an external benchmark.", "Weekly feed/FX check + quarterly peer pack.", "W19", "Gap"),
    ("Work", "Work scattered across 5 GitHub branches", "No 'main' branch; latest loan template v4 is not on the default branch — easy to open the wrong version.", "Consolidate, then keep one default branch.", "W18", "Found in repo"),
    ("Work", "Credentials in automation tools", "Utility importer uses a Gmail app password in .env; Webull would need API keys.", "Store in a password manager; rotate; use a dedicated mailbox.", "W15", "Good practice"),
    ("Work", "Biological assets at year-end", "Live birds / breeders under MFRS 141 — count and valuation basis need auditor agreement.", "Plan the count with the auditor now.", "W22", "Good practice"),
    ("Personal", "Defect Liability Period on B-43-3A", "Developer fixes defects free only within the DLP (typically 24 months from VP). Fit-out can hide defects.", "Inspect and log defects BEFORE renovation starts.", "P04", "To verify VP date"),
    ("Personal", "Tenancy stamp duty & insurance", "Unstamped tenancy agreements are hard to enforce; standard fire insurance often excludes landlord contents/liability.", "Stamp within 30 days; buy landlord cover.", "P05", "Good practice"),
    ("Personal", "Rental income is taxable", "Net rental income goes in your YA return; keep receipts for allowable expenses (maintenance, quit rent, assessment, insurance, interest).", "Start a rental ledger from day 1.", "P08", "Good practice"),
    ("Personal", "US estate tax on US stocks", "Non-resident aliens: US estate tax above USD60k of US-situs assets; Malaysia has no US estate tax treaty.", "Review holding size / structure with an adviser.", "P07", "To verify"),
    ("Personal", "Claude credit expiry 5-Nov-2026", "Unused balance lapses tomorrow.", "Spend on covenant inputs or farmer settlement generator.", "P01", "Found in session"),
]
for i, s in enumerate(spots, 1):
    r = 3 + i
    cell(bs, r, 1, i)
    for c, v in enumerate(s, 2):
        cell(bs, r, c, v, font=BLUE if c == 7 and "verify" in v.lower() else NORM)
bs.freeze_panes = "C4"
bs.auto_filter.ref = f"A3:G{3 + len(spots)}"

# ============================================================ SHAREPOINT PLAN
sp = wb.create_sheet("SharePoint_Plan")
sp["A1"] = "SharePoint Reorganisation — Plan (nothing executes until you approve Phase 3)"; sp["A1"].font = TITLE
header(sp, 3, ["Phase", "Step", "What happens", "Output", "Your decision needed", "Status"], [8, 28, 58, 30, 34, 12])
phases = [
    ("0", "Access", "Either (a) connect a Microsoft 365 / SharePoint connector to Claude with read-only scope, or (b) you run a PnP PowerShell / SharePoint 'Export to CSV' inventory and upload the file.", "Inventory CSV: site, library, folder path, file name, type, size, created/modified, owner, sharing", "Which route? Read-only first.", "Waiting"),
    ("1", "Inventory & analysis", "Map every site/library/folder; flag duplicates, deep nesting (>4 levels), stale files (>2y untouched), personal files, broken permissions, external sharing links.", "Findings sheet + heat map by department", "None", "Not started"),
    ("2", "Target design", "Propose structure by function → entity → year; naming convention; metadata (Entity, FY, Period, DocType); permission groups; retention.", "Target tree + mapping old path → new path", "Function-first or entity-first?", "Not started"),
    ("3", "Approval", "You review the mapping file row by row; agree owners per library; agree freeze window.", "Signed-off mapping", "Approve / amend", "Not started"),
    ("4", "Migration", "Move in batches (one department at a time) using SharePoint 'Move to' so version history is kept; fix links; re-apply permissions.", "Migration log", "Go / no-go per batch", "Not started"),
    ("5", "Governance", "Folder-creation rules, quarterly permission review, retention labels, owner per library.", "1-page SOP", "Approve SOP", "Not started"),
]
for i, p in enumerate(phases):
    r = 4 + i
    for c, v in enumerate(p, 1):
        cell(sp, r, c, v, fill=INPUT if c == 6 else None)

r = 12
sp.cell(row=r, column=1, value="Draft target structure (to be validated against the actual inventory)").font = BOLD
header(sp, r + 1, ["Level", "Site / Library", "Folder pattern", "Example", "Access", ""])
tree = [
    ("Site", "Finance", "/{Entity}/{FY}/{Period}", "/Farmiera Agro/FY2026/2026-11 Nov", "Finance team"),
    ("Library", "Finance – Monthly Close", "/{Entity}/{FY}/{MM}/01 Recs, 02 Accruals, 03 Mgmt Pack", "/Breeder Farm/FY2027/03/01 Recs", "Finance team"),
    ("Library", "Finance – Treasury & Loans", "/{Bank}/{Facility}/Letters, Statements, Covenants", "/Maybank/CMTF-i/Statements", "HoF + Treasury"),
    ("Library", "Finance – Tax & Statutory", "/{Entity}/{YA}/CP204, Form C, SST, e-Invoice", "/Holdings/YA2026/Form C", "HoF + tax agent (guest)"),
    ("Library", "Finance – Audit", "/{Entity}/{FY}/PBC, Queries, Signed FS", "/Agro/FY2026/PBC", "HoF + auditor (guest, time-limited)"),
    ("Library", "Farm Operations", "/{Farm}/{House}/{Batch}", "/Farm A/House 3/B2026-45", "Farm managers"),
    ("Library", "Contract Farmers", "/{Farmer}/{Cycle}/Settlement, e-Invoice", "/Farmer 012/Cycle 2026-09", "Finance + Ops (restricted)"),
    ("Library", "Procurement – Feed & DOC", "/{Supplier}/{Year}/Contracts, Price lists", "/Supplier X/2026", "Procurement + HoF"),
    ("Site", "HR (separate, restricted)", "/Payroll, Staff files, Policies", "/Payroll/2026/11", "HR only"),
    ("Site", "Management", "/Board, Budgets, Peer benchmarking", "/Board/2026-Q4", "Directors + HoF"),
]
for i, t in enumerate(tree):
    for c, v in enumerate(t, 1):
        cell(sp, r + 2 + i, c, v)
rr = r + 3 + len(tree)
sp.cell(row=rr, column=1, value="Naming convention: YYYY-MM-DD_Entity_DocType_Description_vN  (e.g. 2026-11-30_AGRO_BankRec_Maybank_v1.xlsx)").font = NORM
sp.cell(row=rr + 1, column=1, value="Rule of thumb: max 4 folder levels; metadata (Entity, FY, DocType) instead of deeper folders; break permission inheritance only at library level.").font = NORM

# ============================================================ SOURCES
so = wb.create_sheet("Sources")
so["A1"] = "Sessions reviewed"; so["A1"].font = TITLE
header(so, 3, ["Date", "Session", "Area", "Outcome", "Open items → Tracker"], [12, 36, 9, 60, 20])
srcs = [
    (date(2026, 5, 4), "Malaysia commercial banking loan template", "Work", "v1–v3 built (Maybank CMTF-i, Alliance FL, MBSB Tawarruq)", "W04, W05, W09"),
    (date(2026, 5, 5), "Poultry farm management tracking system", "Work", "Flask app: batches, FCR, mortality, harvest projection", "W16"),
    (date(2026, 5, 5), "Build World Cup 2026 tracking app", "Personal", "index.html tracker; PR #1 still open", "P09"),
    (date(2026, 5, 5), "Integrate Webull API connection", "Personal", "Confirmed official Webull OpenAPI exists; nothing built", "P06, P07"),
    (date(2026, 5, 5), "Implement a straightforward TODO", "—", "No substantive outcome", "—"),
    (date(2026, 6, 8), "Automated utility bills import system", "Work", "Email→PDF→SQL Account import pipeline; needs tuning", "W15"),
    (date(2026, 11, 1), "Interior design for tower B-43-3A", "Personal", "Design pack v5, 12 interior views; tenant-ready RM21,140", "P02, P03, P04, P05"),
    (date(2026, 11, 2), "Budget allocation brainstorm", "Work", "Loan template v4: covenants, feed stress, MFRS 7/123; v3 errors fixed", "W01–W03, W06–W08, W10, W11, W17"),
]
for i, s in enumerate(srcs):
    for c, v in enumerate(s, 1):
        cell(so, 4 + i, c, v, fmt=DFMT if c == 1 else None)

# ============================================================ DASHBOARD
db = wb.create_sheet("Dashboard", 1)
db["A1"] = "Dashboard"; db["A1"].font = TITLE
db["A2"] = "As at"; db["A2"].font = BOLD
db["B2"] = "=TODAY()"; db["B2"].number_format = DFMT; db["B2"].fill = CALC
for c, w in zip("ABCDEFG", [22, 12, 12, 12, 12, 12, 12]):
    db.column_dimensions[c].width = w
A, H, K, Fp, I = "Tracker!$B$5:$B$150", "Tracker!$H$5:$H$150", "Tracker!$K$5:$K$150", "Tracker!$F$5:$F$150", "Tracker!$I$5:$I$150"
header(db, 4, ["By RAG", "Work", "Personal", "Total"])
for i, k in enumerate(["Overdue", "Due ≤14d", "On track", "No date", "Done"]):
    r = 5 + i
    cell(db, r, 1, k, font=BOLD)
    cell(db, r, 2, f'=COUNTIFS({A},"Work",{K},A{r})', fill=CALC)
    cell(db, r, 3, f'=COUNTIFS({A},"Personal",{K},A{r})', fill=CALC)
    cell(db, r, 4, f"=B{r}+C{r}", fill=CALC)
cell(db, 10, 1, "Total", font=BOLD)
for c in "BCD":
    cell(db, 10, " BCD".index(c), f"=SUM({c}5:{c}9)", font=BOLD, fill=CALC)

header(db, 12, ["Open by priority", "Work", "Personal", "Total"])
for i, p in enumerate(["P1", "P2", "P3"]):
    r = 13 + i
    cell(db, r, 1, p, font=BOLD)
    cell(db, r, 2, f'=COUNTIFS({A},"Work",{Fp},A{r},{H},"<>Done",{H},"<>Dropped")', fill=CALC)
    cell(db, r, 3, f'=COUNTIFS({A},"Personal",{Fp},A{r},{H},"<>Done",{H},"<>Dropped")', fill=CALC)
    cell(db, r, 4, f"=B{r}+C{r}", fill=CALC)

header(db, 17, ["By status", "Work", "Personal", "Total"])
for i, s in enumerate(["Not started", "In progress", "Waiting", "Done", "Dropped"]):
    r = 18 + i
    cell(db, r, 1, s, font=BOLD)
    cell(db, r, 2, f'=COUNTIFS({A},"Work",{H},A{r})', fill=CALC)
    cell(db, r, 3, f'=COUNTIFS({A},"Personal",{H},A{r})', fill=CALC)
    cell(db, r, 4, f"=B{r}+C{r}", fill=CALC)

cell(db, 24, 1, "Schedule: due in next 7 days", font=BOLD)
cell(db, 24, 2, f'=COUNTIFS(Schedule!$H$5:$H${sc_last},">=0",Schedule!$H$5:$H${sc_last},"<=7")', fill=CALC)
cell(db, 25, 1, "Next P1 due date (open)", font=BOLD)
cell(db, 25, 2, f'=IFERROR(_xlfn.MINIFS({I},{Fp},"P1",{H},"<>Done"),"")', fill=CALC, fmt=DFMT)
db.conditional_formatting.add("B5:D5", FormulaRule(formula=["B5>0"], fill=red))
db.conditional_formatting.add("B6:D6", FormulaRule(formula=["B6>0"], fill=amb))

wb.save(OUT)
print("saved", OUT)
