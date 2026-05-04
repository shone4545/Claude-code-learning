"""
Malaysia Commercial Banking Loan Repayment Template Builder
Bank 1: Maybank Islamic - CMTF-i (Commodity Murabahah Term Financing-i)
Borrower: FARMIERA BREEDER FARM SDN. BHD. (Account 462393920113)

Generates a multi-sheet xlsx workbook designed for MFRS-compliant
financial statement disclosure (current vs non-current principal,
profit expense recognition, and movement reconciliation).
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = "/home/user/Claude-code-learning/Loan_Repayment_Template_v1_MBB.xlsx"

# ---- Styles ----
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
NOTE_FONT = Font(italic=True, color="595959", size=9)
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'
PCT_FMT = '0.0000%'
DATE_FMT = 'dd-mmm-yyyy'

wb = Workbook()

# =============================================================
# SHEET 1: README
# =============================================================
ws = wb.active
ws.title = "README"
readme = [
    ("Malaysia Commercial Banking — Loan Repayment Template", True),
    ("Version 1.0 (Maybank Islamic CMTF-i baseline)", False),
    ("", False),
    ("PURPOSE", True),
    ("Capture full repayment schedule per facility / tranche so that the entity can:", False),
    ("  1. Recognise profit (or interest) expense on accrual basis (MFRS 9 / 109).", False),
    ("  2. Split principal outstanding into Current vs Non-current at FYE (MFRS 101).", False),
    ("  3. Disclose maturity profile: <1yr, 1-2yr, 2-5yr, >5yr (MFRS 7).", False),
    ("  4. Reconcile borrowings movement for cash flow note (MFRS 107).", False),
    ("", False),
    ("COLOUR LEGEND", True),
    ("  YELLOW cells = manual input (from facility letter / bank statement)", False),
    ("  GREEN  cells = auto-calculated formula — do not overwrite", False),
    ("  BLUE   header = section / sheet header", False),
    ("", False),
    ("SHEETS", True),
    ("  1. Facility_Master   — one row per facility / tranche (key terms)", False),
    ("  2. Drawdown_Log      — log of each drawdown with date and amount", False),
    ("  3. Amortisation_MBB  — forward-projected schedule per instalment", False),
    ("  4. FS_Disclosure     — auto-summary for financial statements", False),
    ("  5. Ref_TxnCodes      — Maybank entry-code dictionary (for statement reading)", False),
    ("", False),
    ("WORKFLOW", True),
    ("  Step 1: Fill Facility_Master with each tranche's terms.", False),
    ("  Step 2: Log every drawdown on Drawdown_Log (especially during availability period).", False),
    ("  Step 3: Amortisation_MBB auto-projects period-by-period balances.", False),
    ("  Step 4: Confirm closing balance reconciles to bank statement at last statement date.", False),
    ("  Step 5: Read FS_Disclosure for the figures to lift into the FS notes.", False),
    ("", False),
    ("KEY ISLAMIC vs CONVENTIONAL NOTE", True),
    ("  Maybank CMTF-i is Islamic financing — Commodity Murabahah Term Financing-i.", False),
    ("  Use 'PROFIT' (Bank's earnings) instead of 'INTEREST' in disclosure wording.", False),
    ("  Economic substance for FS classification (current/non-current) is identical.", False),
    ("  The bank statement uses 'PROGRESSIVE PROFIT' for monthly profit accrual entries.", False),
    ("", False),
    ("AVAILABILITY PERIOD (CMTF-i 2-5)", True),
    ("  36-month availability: drawdowns progressive, profit accrues on outstanding,", False),
    ("  no principal instalment yet (profit-only servicing). Repayment of principal", False),
    ("  begins after availability period and runs for the remaining tenure.", False),
    ("", False),
    ("EXTENSIBILITY", True),
    ("  Add new sheets 'Amortisation_ALB' (Alliance Bank) and 'Amortisation_MBSB'", False),
    ("  by copying the Amortisation_MBB sheet structure and adjusting for each", False),
    ("  bank's quirks (e.g. BLR/BR-spread floating rate, balloon payment, daily rest).", False),
]
for i, (text, is_header) in enumerate(readme, start=1):
    cell = ws.cell(row=i, column=1, value=text)
    if is_header:
        cell.font = HDR_FONT if i == 1 else SUB_FONT
        if i == 1:
            cell.fill = HDR_FILL
        else:
            cell.fill = SUB_FILL
ws.column_dimensions['A'].width = 100

# =============================================================
# SHEET 2: Facility_Master
# =============================================================
ws = wb.create_sheet("Facility_Master")
fm_cols = [
    ("Facility_ID", 14, "MBB-CMTFi-1"),
    ("Bank", 18, "Maybank Islamic Bhd"),
    ("Account_No", 18, "462393920113"),
    ("Borrower", 32, "FARMIERA BREEDER FARM SDN. BHD."),
    ("Facility_Type", 22, "CMTF-i (Islamic TL)"),
    ("Tranche", 12, "CMTF-i 1"),
    ("Facility_Limit_RM", 16, 21000000),
    ("Drawdown_Date_First", 16, "2024-07-01"),
    ("Tenure_Months", 12, 180),
    ("Availability_Months", 14, 0),
    ("Repayment_Months", 14, 180),
    ("Maturity_Date", 14, "=EDATE(H2,I2)"),
    ("Profit_Rate_pa", 14, 0.0567),
    ("Rest_Basis", 12, "Monthly"),
    ("Instalment_RM", 14, 144660.00),
    ("Instalment_Type", 14, "EPI (incl. profit)"),
    ("First_Instalment_Date", 16, "2024-08-01"),
    ("Reporting_FYE", 14, "2026-03-31"),
    ("Last_Stmt_Date", 14, "2024-12-31"),
    ("Last_Stmt_Balance_RM", 18, 11749844.16),
    ("Currency", 10, "MYR"),
    ("Notes", 40, "180 equal monthly instalments inclusive of profit. Negative balance on stmt = bank's books convention."),
]
# Headers
for col_idx, (h, w, _) in enumerate(fm_cols, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.alignment = CENTER
    c.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = w

# Sample row 1: CMTF-i 1
for col_idx, (_, _, v) in enumerate(fm_cols, start=1):
    c = ws.cell(row=2, column=col_idx, value=v)
    c.fill = INPUT_FILL
    c.border = BORDER
    if col_idx in (8, 12, 17, 18, 19):
        c.number_format = DATE_FMT
    elif col_idx in (7, 15, 20):
        c.number_format = NUM_FMT
    elif col_idx == 13:
        c.number_format = PCT_FMT
ws.cell(row=2, column=12).fill = CALC_FILL  # maturity is formula

# Sample rows for CMTF-i 2..5 (placeholder skeleton — user fills actuals)
sub_rows = [
    ("MBB-CMTFi-2", "CMTF-i 2", 216, 36, 180),
    ("MBB-CMTFi-3", "CMTF-i 3", 216, 36, 180),
    ("MBB-CMTFi-4", "CMTF-i 4", 216, 36, 180),
    ("MBB-CMTFi-5", "CMTF-i 5", 216, 36, 180),
]
for i, (fid, tr, ten, av, rep) in enumerate(sub_rows, start=3):
    ws.cell(row=i, column=1, value=fid).fill = INPUT_FILL
    ws.cell(row=i, column=2, value="Maybank Islamic Bhd").fill = INPUT_FILL
    ws.cell(row=i, column=3, value="462393920113").fill = INPUT_FILL
    ws.cell(row=i, column=4, value="FARMIERA BREEDER FARM SDN. BHD.").fill = INPUT_FILL
    ws.cell(row=i, column=5, value="CMTF-i (Islamic TL)").fill = INPUT_FILL
    ws.cell(row=i, column=6, value=tr).fill = INPUT_FILL
    ws.cell(row=i, column=7, value=None).fill = INPUT_FILL
    ws.cell(row=i, column=8, value=None).fill = INPUT_FILL
    ws.cell(row=i, column=9, value=ten).fill = INPUT_FILL
    ws.cell(row=i, column=10, value=av).fill = INPUT_FILL
    ws.cell(row=i, column=11, value=rep).fill = INPUT_FILL
    ws.cell(row=i, column=12, value=f"=IF(H{i}=\"\",\"\",EDATE(H{i},I{i}))").fill = CALC_FILL
    ws.cell(row=i, column=12).number_format = DATE_FMT
    ws.cell(row=i, column=13, value=0.0567).fill = INPUT_FILL
    ws.cell(row=i, column=13).number_format = PCT_FMT
    ws.cell(row=i, column=14, value="Monthly").fill = INPUT_FILL
    ws.cell(row=i, column=15, value=None).fill = INPUT_FILL
    ws.cell(row=i, column=15).number_format = NUM_FMT
    ws.cell(row=i, column=16, value="EPI (incl. profit)").fill = INPUT_FILL
    ws.cell(row=i, column=17, value=None).fill = INPUT_FILL
    ws.cell(row=i, column=17).number_format = DATE_FMT
    ws.cell(row=i, column=18, value="2026-03-31").fill = INPUT_FILL
    ws.cell(row=i, column=18).number_format = DATE_FMT
    ws.cell(row=i, column=19, value="2024-12-31").fill = INPUT_FILL
    ws.cell(row=i, column=19).number_format = DATE_FMT
    ws.cell(row=i, column=20, value=None).fill = INPUT_FILL
    ws.cell(row=i, column=20).number_format = NUM_FMT
    ws.cell(row=i, column=21, value="MYR").fill = INPUT_FILL
    ws.cell(row=i, column=22, value="216 mths inc. 36 mths availability period.").fill = INPUT_FILL
    for c in range(1, 23):
        ws.cell(row=i, column=c).border = BORDER

ws.freeze_panes = "B2"

# =============================================================
# SHEET 3: Drawdown_Log
# =============================================================
ws = wb.create_sheet("Drawdown_Log")
dl_cols = [
    ("Facility_ID", 16),
    ("Drawdown_Date", 14),
    ("Drawdown_Amount_RM", 18),
    ("Stmt_Ref", 14),
    ("Description", 40),
    ("Cumulative_Outstanding_RM", 22),
    ("Notes", 30),
]
for col_idx, (h, w) in enumerate(dl_cols, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.alignment = CENTER
    c.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = w

# Pre-populate observed drawdowns from the screenshot (CMTF-i 1)
# Drawdown debits seen on screenshot (DR amounts mean drawdowns): 1,079,864.66 + others on same statement
# Per screenshot first drawdown 0410 = 1,079,864.66 (likely CMTF-i 1)
mbb_dd = [
    ("MBB-CMTFi-1", "2024-07-01", 1079864.66, "0410/N", "DRAWDOWN DEBIT", "First tranche drawdown"),
    ("MBB-CMTFi-2", "2024-10-01", 1345971.14, "0402/N", "DRAWDOWN DEBIT", "Per stmt entry 0402"),
    ("MBB-CMTFi-2", "2024-10-01", 2283980.12, "1003/N", "DRAWDOWN DEBIT", "Per stmt entry 1003"),
    ("MBB-CMTFi-3", "2024-11-01", 1238287.64, "0605/N", "DRAWDOWN DEBIT", "Per stmt entry 0605"),
    ("MBB-CMTFi-3", "2024-11-01", 2284631.16, "3006/N", "DRAWDOWN DEBIT", "Per stmt entry 3006"),
    ("MBB-CMTFi-4", "2024-11-01", 1395804.56, "2807/N", "DRAWDOWN DEBIT", "Per stmt entry 2807"),
    ("MBB-CMTFi-4", "2024-12-01", 917380.84,  "0309/N", "DRAWDOWN DEBIT", "Per stmt entry 0309"),
    ("MBB-CMTFi-5", "2024-12-01", 1158197.81, "0810/N", "DRAWDOWN DEBIT", "Per stmt entry 0810"),
]
for i, row in enumerate(mbb_dd, start=2):
    for col_idx, v in enumerate(row, start=1):
        c = ws.cell(row=i, column=col_idx, value=v)
        c.fill = INPUT_FILL
        c.border = BORDER
        if col_idx == 2:
            c.number_format = DATE_FMT
        elif col_idx == 3:
            c.number_format = NUM_FMT
    # Cumulative running total
    cum = ws.cell(row=i, column=6, value=f"=SUMIFS($C$2:C{i},$A$2:A{i},A{i})")
    cum.fill = CALC_FILL
    cum.number_format = NUM_FMT
    cum.border = BORDER

ws.freeze_panes = "B2"

# =============================================================
# SHEET 4: Amortisation_MBB
# =============================================================
ws = wb.create_sheet("Amortisation_MBB")

# Header / Lookup section (rows 1-6)
ws["A1"] = "MAYBANK ISLAMIC — CMTF-i AMORTISATION SCHEDULE (Forward Projection)"
ws["A1"].font = HDR_FONT
ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:K1")

ws["A3"] = "Facility_ID:"
ws["B3"] = "MBB-CMTFi-1"
ws["B3"].fill = INPUT_FILL

ws["D3"] = "Reporting FYE:"
ws["E3"] = "=VLOOKUP(B3,Facility_Master!A:R,18,FALSE)"
ws["E3"].number_format = DATE_FMT
ws["E3"].fill = CALC_FILL

ws["G3"] = "Profit Rate p.a.:"
ws["H3"] = "=VLOOKUP(B3,Facility_Master!A:M,13,FALSE)"
ws["H3"].number_format = PCT_FMT
ws["H3"].fill = CALC_FILL

ws["A4"] = "Opening Principal:"
ws["B4"] = "=VLOOKUP(B3,Facility_Master!A:G,7,FALSE)"
ws["B4"].number_format = NUM_FMT
ws["B4"].fill = CALC_FILL

ws["D4"] = "Tenure (mths):"
ws["E4"] = "=VLOOKUP(B3,Facility_Master!A:I,9,FALSE)"
ws["E4"].fill = CALC_FILL

ws["G4"] = "Instalment RM:"
ws["H4"] = "=VLOOKUP(B3,Facility_Master!A:O,15,FALSE)"
ws["H4"].number_format = NUM_FMT
ws["H4"].fill = CALC_FILL

ws["A5"] = "First Instalment Date:"
ws["B5"] = "=VLOOKUP(B3,Facility_Master!A:Q,17,FALSE)"
ws["B5"].number_format = DATE_FMT
ws["B5"].fill = CALC_FILL

ws["D5"] = "Availability (mths):"
ws["E5"] = "=VLOOKUP(B3,Facility_Master!A:J,10,FALSE)"
ws["E5"].fill = CALC_FILL

ws["G5"] = "Maturity Date:"
ws["H5"] = "=VLOOKUP(B3,Facility_Master!A:L,12,FALSE)"
ws["H5"].number_format = DATE_FMT
ws["H5"].fill = CALC_FILL

# Schedule headers
sch_hdr_row = 8
sch_cols = [
    "Period", "Due_Date", "Phase",
    "Opening_Bal_RM", "Profit_RM", "Instalment_RM",
    "Principal_RM", "Closing_Bal_RM",
    "Months_to_FYE", "Bucket", "Cumulative_Profit_RM"
]
for col_idx, h in enumerate(sch_cols, start=1):
    c = ws.cell(row=sch_hdr_row, column=col_idx, value=h)
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.alignment = CENTER
    c.border = BORDER

widths = [8, 14, 16, 16, 14, 14, 14, 16, 12, 14, 18]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Build 180 rows of formulas
start_row = 9
n_periods = 180
for p in range(1, n_periods + 1):
    r = start_row + p - 1
    # A: period
    ws.cell(row=r, column=1, value=p).border = BORDER
    # B: due date = EDATE(first_instalment, p-1)
    ws.cell(row=r, column=2, value=f"=EDATE($B$5,A{r}-1)").number_format = DATE_FMT
    # C: phase
    ws.cell(row=r, column=3,
            value=f"=IF(A{r}<=$E$5,\"Availability\",\"Repayment\")")
    # D: opening balance — period 1 = facility limit; subsequent = prior closing
    if p == 1:
        ws.cell(row=r, column=4, value=f"=$B$4")
    else:
        ws.cell(row=r, column=4, value=f"=H{r-1}")
    ws.cell(row=r, column=4).number_format = NUM_FMT
    # E: profit = opening * rate / 12
    ws.cell(row=r, column=5, value=f"=ROUND(D{r}*$H$3/12,2)").number_format = NUM_FMT
    # F: instalment — during availability profit-only; else fixed instalment; final period plug
    ws.cell(row=r, column=6,
            value=f"=IF(C{r}=\"Availability\",E{r},IF(A{r}=$E$4,D{r}+E{r},$H$4))").number_format = NUM_FMT
    # G: principal = instalment - profit (floor 0)
    ws.cell(row=r, column=7, value=f"=MAX(F{r}-E{r},0)").number_format = NUM_FMT
    # H: closing = opening - principal
    ws.cell(row=r, column=8, value=f"=MAX(D{r}-G{r},0)").number_format = NUM_FMT
    # I: months to FYE = DATEDIF(due_date, FYE, "m")  (negative if past FYE)
    ws.cell(row=r, column=9,
            value=f"=IF(B{r}>$E$3,-(DATEDIF($E$3,B{r},\"m\")+1),DATEDIF(B{r},$E$3,\"m\"))")
    # J: bucket
    ws.cell(row=r, column=10,
            value=(
                f'=IF(B{r}<=$E$3,"Past/Current",'
                f'IF(I{r}<=12,"Current",'
                f'IF(I{r}<=24,"1-2 yrs",'
                f'IF(I{r}<=60,"2-5 yrs","Over 5 yrs"))))'
            ))
    # K: cumulative profit
    if p == 1:
        ws.cell(row=r, column=11, value=f"=E{r}")
    else:
        ws.cell(row=r, column=11, value=f"=K{r-1}+E{r}")
    ws.cell(row=r, column=11).number_format = NUM_FMT
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = BORDER

ws.freeze_panes = "A9"

# =============================================================
# SHEET 5: FS_Disclosure
# =============================================================
ws = wb.create_sheet("FS_Disclosure")
ws["A1"] = "FINANCIAL STATEMENT DISCLOSURE — BORROWINGS"
ws["A1"].font = HDR_FONT
ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:E1")

ws["A3"] = "Facility:"
ws["B3"] = "=Amortisation_MBB!B3"
ws["A4"] = "Reporting Date (FYE):"
ws["B4"] = "=Amortisation_MBB!E3"
ws["B4"].number_format = DATE_FMT

# Section: Outstanding split
ws["A6"] = "OUTSTANDING PRINCIPAL AT FYE — MFRS 101 Classification"
ws["A6"].font = SUB_FONT
ws["A6"].fill = SUB_FILL
ws.merge_cells("A6:E6")

ws["A7"] = "Closing balance at FYE (RM)"
ws["B7"] = '=INDEX(Amortisation_MBB!H:H,MATCH(B4,Amortisation_MBB!B:B,0))'
ws["B7"].number_format = NUM_FMT

ws["A8"] = "Current portion (next 12 mths principal) (RM)"
ws["B8"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"Current")'
ws["B8"].number_format = NUM_FMT

ws["A9"] = "Non-current portion (RM)"
ws["B9"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"1-2 yrs")+SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"2-5 yrs")+SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"Over 5 yrs")'
ws["B9"].number_format = NUM_FMT

ws["A10"] = "Cross-check: Current + Non-current"
ws["B10"] = "=B8+B9"
ws["B10"].number_format = NUM_FMT
ws["A11"] = "Variance vs Closing balance"
ws["B11"] = "=B7-B10"
ws["B11"].number_format = NUM_FMT

# Section: Maturity profile
ws["A13"] = "MATURITY PROFILE — MFRS 7 Liquidity Risk Disclosure"
ws["A13"].font = SUB_FONT
ws["A13"].fill = SUB_FILL
ws.merge_cells("A13:E13")

ws["A14"] = "Within 1 year (RM)"
ws["B14"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"Current")'
ws["B14"].number_format = NUM_FMT

ws["A15"] = "1 to 2 years (RM)"
ws["B15"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"1-2 yrs")'
ws["B15"].number_format = NUM_FMT

ws["A16"] = "2 to 5 years (RM)"
ws["B16"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"2-5 yrs")'
ws["B16"].number_format = NUM_FMT

ws["A17"] = "Over 5 years (RM)"
ws["B17"] = '=SUMIFS(Amortisation_MBB!G:G,Amortisation_MBB!J:J,"Over 5 yrs")'
ws["B17"].number_format = NUM_FMT

ws["A18"] = "Total"
ws["A18"].font = SUB_FONT
ws["B18"] = "=SUM(B14:B17)"
ws["B18"].number_format = NUM_FMT
ws["B18"].font = SUB_FONT

# Section: Profit expense
ws["A20"] = "PROFIT EXPENSE FOR THE YEAR — MFRS 9"
ws["A20"].font = SUB_FONT
ws["A20"].fill = SUB_FILL
ws.merge_cells("A20:E20")

ws["A21"] = "Year start date"
ws["B21"] = "=EDATE(B4,-12)+1"
ws["B21"].number_format = DATE_FMT

ws["A22"] = "Profit recognised in P&L (RM)"
ws["B22"] = '=SUMIFS(Amortisation_MBB!E:E,Amortisation_MBB!B:B,">="&B21,Amortisation_MBB!B:B,"<="&B4)'
ws["B22"].number_format = NUM_FMT

# Section: Movement reconciliation (template)
ws["A24"] = "BORROWINGS MOVEMENT — MFRS 107 Reconciliation"
ws["A24"].font = SUB_FONT
ws["A24"].fill = SUB_FILL
ws.merge_cells("A24:E24")

mvt = [
    ("Opening balance (1 Apr)", None),
    ("Add: Drawdowns during year", None),
    ("Add: Profit accrued", None),
    ("Less: Profit paid", None),
    ("Less: Principal repayments", None),
    ("Closing balance (FYE)", None),
]
for i, (lbl, val) in enumerate(mvt, start=25):
    ws.cell(row=i, column=1, value=lbl)
    cell = ws.cell(row=i, column=2, value=val)
    cell.fill = INPUT_FILL
    cell.number_format = NUM_FMT

ws.column_dimensions['A'].width = 50
ws.column_dimensions['B'].width = 20

# =============================================================
# SHEET 6: Ref_TxnCodes
# =============================================================
ws = wb.create_sheet("Ref_TxnCodes")
ws["A1"] = "MAYBANK ISLAMIC — STATEMENT TRANSACTION CODE DICTIONARY"
ws["A1"].font = HDR_FONT
ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:D1")

codes = [
    ("Code Pattern", "Meaning", "DR/CR Effect", "Treatment in Schedule"),
    ("0xxx/N",        "DRAWDOWN DEBIT",          "Increases outstanding (DR)", "Add to principal — log in Drawdown_Log"),
    ("3xxx/",         "PROGRESSIVE PROFIT",      "Increases outstanding (DR)", "Profit accrual — flows to P&L on accrual basis"),
    ("0xxx/ BRH...",  "Settlement (profit/principal)", "Reduces outstanding (CR)", "Split into profit-paid and principal-paid"),
    ("0410/",         "Initial drawdown of CMTF-i 1", "DR", "First drawdown row"),
    ("0402/, 1003/",  "CMTF-i 2 progressive drawdowns", "DR", "Multi-tranche drawdowns"),
    ("0605/, 3006/",  "CMTF-i 3 progressive drawdowns", "DR", ""),
    ("2807/, 0309/",  "CMTF-i 4 progressive drawdowns", "DR", ""),
    ("0810/",         "CMTF-i 5 progressive drawdown", "DR", ""),
    ("BRH000003",     "Branch reference for CMTF-i 1-3 settlements", "n/a", "Used to match settlement to facility"),
    ("BRH000005",     "Branch reference for CMTF-i 4-5 settlements", "n/a", ""),
    ("DISPLAY 1",     "Last year transactions", "n/a", "Reading aid"),
    ("DISPLAY 2",     "Current year transactions", "n/a", "Reading aid"),
    ("Negative balance", "Bank's books — liability of borrower", "n/a", "Convert to positive when normalising for FS"),
]
for r, row in enumerate(codes, start=3):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BORDER
        if r == 3:
            cell.font = HDR_FONT
            cell.fill = HDR_FILL
            cell.alignment = CENTER

for col, w in zip("ABCD", [22, 38, 32, 50]):
    ws.column_dimensions[col].width = w

# Save
wb.save(OUT)
print(f"Saved: {OUT}")
