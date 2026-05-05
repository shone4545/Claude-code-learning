"""
Malaysia Commercial Banking Loan Repayment Template Builder — v2
  Bank 1: Maybank Islamic - CMTF-i (FARMIERA BREEDER FARM SDN. BHD.)
  Bank 2: Alliance Bank - Fixed Loan (FARMIERA AGRO RESOURCES SDN. BHD.)

MFRS-compliant disclosure: current/non-current principal, interest/profit
expense, MFRS 7 maturity profile, MFRS 107 movement reconciliation.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/Claude-code-learning/Loan_Repayment_Template_v2_MBB_ALB.xlsx"

# ---- Styles ----
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

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
    ("Version 2.0  (Maybank Islamic CMTF-i + Alliance Bank Fixed Loan)", False),
    ("", False),
    ("PURPOSE", True),
    ("Capture full repayment schedule per facility so the entity can:", False),
    ("  1. Recognise interest / profit expense on accrual basis (MFRS 9).", False),
    ("  2. Split principal outstanding into Current vs Non-current (MFRS 101).", False),
    ("  3. Disclose maturity profile <1y / 1-2y / 2-5y / >5y (MFRS 7).", False),
    ("  4. Reconcile borrowings movement for cash flow note (MFRS 107).", False),
    ("", False),
    ("COLOUR LEGEND", True),
    ("  YELLOW = manual input (from facility letter / bank statement)", False),
    ("  GREEN  = auto-calculated formula — do NOT overwrite", False),
    ("  ORANGE = total / cross-check row", False),
    ("  BLUE   = section / sheet header", False),
    ("", False),
    ("SHEETS", True),
    ("  1. Facility_Master    — one row per facility / tranche (key terms)", False),
    ("  2. Drawdown_Log       — log of each drawdown across all banks", False),
    ("  3. Amortisation_MBB   — Maybank Islamic CMTF-i forward schedule (Islamic, EPI)", False),
    ("  4. Amortisation_ALB   — Alliance Bank forward schedule (conventional, fixed principal)", False),
    ("  5. FS_Disclosure_MBB  — disclosure summary for Maybank facility", False),
    ("  6. FS_Disclosure_ALB  — disclosure summary for Alliance Bank facility", False),
    ("  7. FS_Aggregate       — group / consolidated borrowings disclosure", False),
    ("  8. Ref_TxnCodes       — Maybank + Alliance Bank statement code dictionary", False),
    ("", False),
    ("BANK MECHANICS RECAP", True),
    ("  MAYBANK ISLAMIC (CMTF-i)", False),
    ("    - Islamic — Commodity Murabahah; 'PROGRESSIVE PROFIT' = monthly profit accrual.", False),
    ("    - Multi-tranche (CMTF-i 1..5) under one account; CMTF-i 2-5 have 36-mth", False),
    ("      availability period (drawdown phase, profit-only servicing).", False),
    ("    - Equal monthly instalment inclusive of profit (EPI).", False),
    ("    - Statement balance shown NEGATIVE (bank's books) — flip to + for FS.", False),
    ("", False),
    ("  ALLIANCE BANK (Fixed Loan)", False),
    ("    - Conventional — INTEREST.", False),
    ("    - Daily rest, billed monthly: '*DR INTEREST BILLED' at month-end,", False),
    ("      '*CR INT (AUTO)' auto-deducted from current account next day.", False),
    ("    - Fixed principal RM70,000/month ('*CR PRIN AUTO').", False),
    ("    - One-shot drawdown via 'LOAN RELEASE'.", False),
    ("    - Statement balance shown DR (positive convention from borrower's view).", False),
    ("    - At FYE there is typically 1 month interest billed but not yet paid", False),
    ("      — that's accrued interest payable, separate line in FS.", False),
    ("", False),
    ("KEY ACCOUNTING TREATMENT", True),
    ("  Loan principal split:", False),
    ("    Current portion       = principal due within 12 mths post-FYE", False),
    ("    Non-current portion   = principal due > 12 mths post-FYE", False),
    ("  Interest / profit expense for the period = SUM of monthly accruals between", False),
    ("    (FYE - 12mth) + 1 day  to  FYE", False),
    ("  Accrued interest at FYE (Alliance Bank): last billed but unpaid amount", False),
    ("    sitting in loan account. Reclass from loan principal balance to accruals.", False),
    ("", False),
    ("EXTENSIBILITY — NEXT", True),
    ("  Add Amortisation_MBSB sheet by copying ALB structure.", False),
    ("  MBSB is typically Islamic (Tawarruq) — adjust terminology to 'profit'.", False),
]
for i, (text, is_header) in enumerate(readme, start=1):
    cell = ws.cell(row=i, column=1, value=text)
    if is_header:
        cell.font = HDR_FONT if i == 1 else SUB_FONT
        cell.fill = HDR_FILL if i == 1 else SUB_FILL
ws.column_dimensions['A'].width = 100

# =============================================================
# SHEET 2: Facility_Master
# =============================================================
ws = wb.create_sheet("Facility_Master")
fm_headers = [
    "Facility_ID", "Bank", "Account_No", "Borrower",
    "Facility_Type", "Tranche", "Facility_Limit_RM",
    "Drawdown_Date_First", "Tenure_Months", "Availability_Months", "Repayment_Months",
    "Maturity_Date", "Rate_pa", "Rest_Basis", "Instalment_RM", "Instalment_Type",
    "First_Instalment_Date", "Reporting_FYE", "Last_Stmt_Date", "Last_Stmt_Balance_RM",
    "Currency", "Notes"
]
widths_fm = [14, 22, 18, 32, 22, 12, 16, 16, 12, 14, 14, 14, 12, 12, 14, 22, 16, 14, 14, 18, 10, 50]
for col_idx, h in enumerate(fm_headers, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = widths_fm[col_idx-1]

facility_rows = [
    # MAYBANK CMTF-i 1
    ("MBB-CMTFi-1", "Maybank Islamic Bhd", "462393920113", "FARMIERA BREEDER FARM SDN. BHD.",
     "CMTF-i (Islamic TL)", "CMTF-i 1", 21000000, "2024-07-01", 180, 0, 180,
     None, 0.0567, "Monthly", 144660.00, "EPI (incl. profit)",
     "2024-08-01", "2026-03-31", "2024-12-31", 1079864.66, "MYR",
     "180 equal monthly instalments inclusive of profit. Statement bal sign-flipped."),
    # MAYBANK CMTF-i 2..5 — placeholders
    ("MBB-CMTFi-2", "Maybank Islamic Bhd", "462393920113", "FARMIERA BREEDER FARM SDN. BHD.",
     "CMTF-i (Islamic TL)", "CMTF-i 2", None, None, 216, 36, 180,
     None, 0.0567, "Monthly", None, "EPI (incl. profit)",
     None, "2026-03-31", "2024-12-31", None, "MYR",
     "216 mths inc. 36 mths availability period. Fill in actuals."),
    ("MBB-CMTFi-3", "Maybank Islamic Bhd", "462393920113", "FARMIERA BREEDER FARM SDN. BHD.",
     "CMTF-i (Islamic TL)", "CMTF-i 3", None, None, 216, 36, 180,
     None, 0.0567, "Monthly", None, "EPI (incl. profit)",
     None, "2026-03-31", "2024-12-31", None, "MYR",
     "216 mths inc. 36 mths availability period."),
    ("MBB-CMTFi-4", "Maybank Islamic Bhd", "462393920113", "FARMIERA BREEDER FARM SDN. BHD.",
     "CMTF-i (Islamic TL)", "CMTF-i 4", None, None, 216, 36, 180,
     None, 0.0567, "Monthly", None, "EPI (incl. profit)",
     None, "2026-03-31", "2024-12-31", None, "MYR",
     "216 mths inc. 36 mths availability period."),
    ("MBB-CMTFi-5", "Maybank Islamic Bhd", "462393920113", "FARMIERA BREEDER FARM SDN. BHD.",
     "CMTF-i (Islamic TL)", "CMTF-i 5", None, None, 216, 36, 180,
     None, 0.0567, "Monthly", None, "EPI (incl. profit)",
     None, "2026-03-31", "2024-12-31", None, "MYR",
     "216 mths inc. 36 mths availability period."),
    # ALLIANCE BANK Fixed Loan
    ("ALB-FL-1", "Alliance Bank Malaysia Bhd", "040100052028231", "FARMIERA AGRO RESOURCES SDN. BHD.",
     "Fixed Loan (Conventional TL)", "-", None, None, None, 0, None,
     None, 0.0520, "Daily (billed mthly)", 70000.00, "Fixed Principal + Variable Interest",
     "2025-04-01", "2025-12-31", "2025-12-31", 6953133.21, "MYR",
     "Per stmt 31/12/25 closing 6,953,133.21 incl. accrued int 28,406.56 (billed 31/12 not yet auto-deducted). Principal-only 6,924,726.65."),
]
for r, row in enumerate(facility_rows, start=2):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.fill = INPUT_FILL; cell.border = BORDER
        if c in (8, 17, 18, 19):
            cell.number_format = DATE_FMT
        elif c == 12:
            cell.number_format = DATE_FMT
        elif c in (7, 15, 20):
            cell.number_format = NUM_FMT
        elif c == 13:
            cell.number_format = PCT_FMT
    # maturity formula
    mat = ws.cell(row=r, column=12,
                  value=f'=IF(OR(H{r}="",I{r}=""),"",EDATE(H{r},I{r}))')
    mat.fill = CALC_FILL; mat.number_format = DATE_FMT; mat.border = BORDER
ws.freeze_panes = "B2"

# =============================================================
# SHEET 3: Drawdown_Log
# =============================================================
ws = wb.create_sheet("Drawdown_Log")
dl_headers = ["Facility_ID", "Drawdown_Date", "Drawdown_Amount_RM",
              "Stmt_Ref", "Description", "Cumulative_per_Facility_RM", "Notes"]
widths_dl = [16, 14, 18, 16, 38, 22, 30]
for col_idx, h in enumerate(dl_headers, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = widths_dl[col_idx-1]

drawdowns = [
    ("MBB-CMTFi-1", "2024-07-01", 1079864.66, "0410/N", "DRAWDOWN DEBIT", "First drawdown CMTF-i 1"),
    ("MBB-CMTFi-2", "2024-10-01", 1345971.14, "0402/N", "DRAWDOWN DEBIT", "Per stmt 0402"),
    ("MBB-CMTFi-2", "2024-10-01", 2283980.12, "1003/N", "DRAWDOWN DEBIT", "Per stmt 1003"),
    ("MBB-CMTFi-3", "2024-11-01", 1238287.64, "0605/N", "DRAWDOWN DEBIT", "Per stmt 0605"),
    ("MBB-CMTFi-3", "2024-11-01", 2284631.16, "3006/N", "DRAWDOWN DEBIT", "Per stmt 3006"),
    ("MBB-CMTFi-4", "2024-11-01", 1395804.56, "2807/N", "DRAWDOWN DEBIT", "Per stmt 2807"),
    ("MBB-CMTFi-4", "2024-12-01", 917380.84,  "0309/N", "DRAWDOWN DEBIT", "Per stmt 0309"),
    ("MBB-CMTFi-5", "2024-12-01", 1158197.81, "0810/N", "DRAWDOWN DEBIT", "Per stmt 0810"),
    ("ALB-FL-1",    "2025-01-24", 1410876.56, "LOAN RELEASE", "LOAN RELEASE", "Drawdown during FY25"),
]
for r, row in enumerate(drawdowns, start=2):
    for c, v in enumerate(row[:5], start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.fill = INPUT_FILL; cell.border = BORDER
        if c == 2: cell.number_format = DATE_FMT
        if c == 3: cell.number_format = NUM_FMT
    cum = ws.cell(row=r, column=6, value=f"=SUMIFS($C$2:C{r},$A$2:A{r},A{r})")
    cum.fill = CALC_FILL; cum.number_format = NUM_FMT; cum.border = BORDER
    nt = ws.cell(row=r, column=7, value=row[5])
    nt.fill = INPUT_FILL; nt.border = BORDER
ws.freeze_panes = "B2"

# =============================================================
# SHEET 4: Amortisation_MBB  (same structure as v1)
# =============================================================
ws = wb.create_sheet("Amortisation_MBB")
ws["A1"] = "MAYBANK ISLAMIC — CMTF-i AMORTISATION (Forward Projection)"
ws["A1"].font = HDR_FONT; ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:K1")

ws["A3"] = "Facility_ID:"
ws["B3"] = "MBB-CMTFi-1"; ws["B3"].fill = INPUT_FILL
ws["D3"] = "Reporting FYE:"
ws["E3"] = "=VLOOKUP(B3,Facility_Master!A:R,18,FALSE)"
ws["E3"].number_format = DATE_FMT; ws["E3"].fill = CALC_FILL
ws["G3"] = "Profit Rate p.a.:"
ws["H3"] = "=VLOOKUP(B3,Facility_Master!A:M,13,FALSE)"
ws["H3"].number_format = PCT_FMT; ws["H3"].fill = CALC_FILL

ws["A4"] = "Opening Principal:"
ws["B4"] = "=VLOOKUP(B3,Facility_Master!A:G,7,FALSE)"
ws["B4"].number_format = NUM_FMT; ws["B4"].fill = CALC_FILL
ws["D4"] = "Tenure (mths):"
ws["E4"] = "=VLOOKUP(B3,Facility_Master!A:I,9,FALSE)"
ws["E4"].fill = CALC_FILL
ws["G4"] = "Instalment RM:"
ws["H4"] = "=VLOOKUP(B3,Facility_Master!A:O,15,FALSE)"
ws["H4"].number_format = NUM_FMT; ws["H4"].fill = CALC_FILL

ws["A5"] = "First Instalment Date:"
ws["B5"] = "=VLOOKUP(B3,Facility_Master!A:Q,17,FALSE)"
ws["B5"].number_format = DATE_FMT; ws["B5"].fill = CALC_FILL
ws["D5"] = "Availability (mths):"
ws["E5"] = "=VLOOKUP(B3,Facility_Master!A:J,10,FALSE)"
ws["E5"].fill = CALC_FILL
ws["G5"] = "Maturity Date:"
ws["H5"] = "=VLOOKUP(B3,Facility_Master!A:L,12,FALSE)"
ws["H5"].number_format = DATE_FMT; ws["H5"].fill = CALC_FILL

mbb_hdr = ["Period", "Due_Date", "Phase", "Opening_Bal_RM", "Profit_RM",
           "Instalment_RM", "Principal_RM", "Closing_Bal_RM",
           "Months_to_FYE", "Bucket", "Cumulative_Profit_RM"]
sch_hdr_row = 8
for col_idx, h in enumerate(mbb_hdr, start=1):
    c = ws.cell(row=sch_hdr_row, column=col_idx, value=h)
    c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
for i, w in enumerate([8, 14, 16, 16, 14, 14, 14, 16, 12, 14, 18], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

n_periods = 180
start_row = 9
for p in range(1, n_periods + 1):
    r = start_row + p - 1
    ws.cell(row=r, column=1, value=p)
    ws.cell(row=r, column=2, value=f"=EDATE($B$5,A{r}-1)").number_format = DATE_FMT
    ws.cell(row=r, column=3, value=f'=IF(A{r}<=$E$5,"Availability","Repayment")')
    ws.cell(row=r, column=4, value=("=$B$4" if p == 1 else f"=H{r-1}")).number_format = NUM_FMT
    ws.cell(row=r, column=5, value=f"=ROUND(D{r}*$H$3/12,2)").number_format = NUM_FMT
    ws.cell(row=r, column=6,
            value=f'=IF(C{r}="Availability",E{r},IF(A{r}=$E$4,D{r}+E{r},$H$4))'
           ).number_format = NUM_FMT
    ws.cell(row=r, column=7, value=f"=MAX(F{r}-E{r},0)").number_format = NUM_FMT
    ws.cell(row=r, column=8, value=f"=MAX(D{r}-G{r},0)").number_format = NUM_FMT
    # Months from FYE to due date (informational, may be 0 for past/current)
    ws.cell(row=r, column=9,
            value=f'=IF(B{r}<=$E$3,0,DATEDIF($E$3,B{r},"m"))')
    # Bucket — uses EDATE thresholds for robustness (handles month-end edges)
    ws.cell(row=r, column=10,
            value=(f'=IF(B{r}<=$E$3,"Past/Current",'
                   f'IF(B{r}<=EDATE($E$3,12),"Current",'
                   f'IF(B{r}<=EDATE($E$3,24),"1-2 yrs",'
                   f'IF(B{r}<=EDATE($E$3,60),"2-5 yrs","Over 5 yrs"))))'))
    ws.cell(row=r, column=11,
            value=(f"=E{r}" if p == 1 else f"=K{r-1}+E{r}")).number_format = NUM_FMT
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = BORDER
ws.freeze_panes = "A9"

# =============================================================
# SHEET 5: Amortisation_ALB  (Alliance Bank — fixed principal + daily-rest interest)
# =============================================================
ws = wb.create_sheet("Amortisation_ALB")
ws["A1"] = "ALLIANCE BANK — FIXED LOAN AMORTISATION (Forward Projection)"
ws["A1"].font = HDR_FONT; ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:L1")

ws["A3"] = "Facility_ID:"
ws["B3"] = "ALB-FL-1"; ws["B3"].fill = INPUT_FILL
ws["D3"] = "Reporting FYE:"
ws["E3"] = "=VLOOKUP(B3,Facility_Master!A:R,18,FALSE)"
ws["E3"].number_format = DATE_FMT; ws["E3"].fill = CALC_FILL
ws["G3"] = "Interest Rate p.a.:"
ws["H3"] = "=VLOOKUP(B3,Facility_Master!A:M,13,FALSE)"
ws["H3"].number_format = PCT_FMT; ws["H3"].fill = CALC_FILL

ws["A4"] = "Schedule Start Bal (Principal-only) RM:"
# Principal-only opening = Last stmt balance LESS accrued interest at FYE
ws["B4"] = 6924726.65
ws["B4"].number_format = NUM_FMT; ws["B4"].fill = INPUT_FILL
ws["D4"] = "Schedule Start Date:"
ws["E4"] = "2026-01-01"
ws["E4"].number_format = DATE_FMT; ws["E4"].fill = INPUT_FILL
ws["G4"] = "Fixed Principal/mth:"
ws["H4"] = "=VLOOKUP(B3,Facility_Master!A:O,15,FALSE)"
ws["H4"].number_format = NUM_FMT; ws["H4"].fill = CALC_FILL

ws["A5"] = "Day-Count Basis:"
ws["B5"] = 365
ws["B5"].fill = INPUT_FILL
ws["D5"] = "Accrued Int at FYE RM:"
ws["E5"] = 28406.56
ws["E5"].number_format = NUM_FMT; ws["E5"].fill = INPUT_FILL
ws["G5"] = "Months projected:"
ws["H5"] = "=ROUNDUP(B4/H4,0)+2"
ws["H5"].fill = CALC_FILL

alb_hdr = ["Period", "Due_Date", "Days_in_Period",
           "Opening_Bal_RM", "Interest_RM", "Principal_RM",
           "Total_Payment_RM", "Closing_Bal_RM",
           "Months_to_FYE", "Bucket",
           "Cumulative_Interest_RM", "Cumulative_Principal_RM"]
sch_hdr_row = 8
for col_idx, h in enumerate(alb_hdr, start=1):
    c = ws.cell(row=sch_hdr_row, column=col_idx, value=h)
    c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
for i, w in enumerate([8, 14, 12, 16, 14, 14, 14, 16, 12, 14, 18, 18], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# 120 month projection (10 years — covers ~99 month payoff)
n_alb = 120
start_row = 9
for p in range(1, n_alb + 1):
    r = start_row + p - 1
    ws.cell(row=r, column=1, value=p)
    # Due date = EOMONTH(start_date, p-1) i.e. month-end
    ws.cell(row=r, column=2, value=f"=EOMONTH($E$4,A{r}-1)").number_format = DATE_FMT
    # Days in period
    ws.cell(row=r, column=3, value=(
        f"=DAY(B{r})" if p == 1 else f"=B{r}-B{r-1}"
    ))
    # Opening
    ws.cell(row=r, column=4, value=("=$B$4" if p == 1 else f"=H{r-1}")).number_format = NUM_FMT
    # Interest = opening * rate * days / 365
    ws.cell(row=r, column=5,
            value=f"=ROUND(D{r}*$H$3*C{r}/$B$5,2)").number_format = NUM_FMT
    # Principal — fixed unless balance smaller
    ws.cell(row=r, column=6,
            value=f"=MIN($H$4,D{r})").number_format = NUM_FMT
    # Total payment = interest + principal
    ws.cell(row=r, column=7, value=f"=E{r}+F{r}").number_format = NUM_FMT
    # Closing
    ws.cell(row=r, column=8, value=f"=MAX(D{r}-F{r},0)").number_format = NUM_FMT
    # Months from FYE to due date (informational)
    ws.cell(row=r, column=9,
            value=f'=IF(B{r}<=$E$3,0,DATEDIF($E$3,B{r},"m"))')
    # Bucket — uses EDATE thresholds (robust across month-end edges)
    ws.cell(row=r, column=10,
            value=(f'=IF(B{r}<=$E$3,"Past/Current",'
                   f'IF(B{r}<=EDATE($E$3,12),"Current",'
                   f'IF(B{r}<=EDATE($E$3,24),"1-2 yrs",'
                   f'IF(B{r}<=EDATE($E$3,60),"2-5 yrs","Over 5 yrs"))))'))
    # Cumulative interest
    ws.cell(row=r, column=11,
            value=(f"=E{r}" if p == 1 else f"=K{r-1}+E{r}")).number_format = NUM_FMT
    # Cumulative principal
    ws.cell(row=r, column=12,
            value=(f"=F{r}" if p == 1 else f"=L{r-1}+F{r}")).number_format = NUM_FMT
    for c in range(1, 13):
        ws.cell(row=r, column=c).border = BORDER

ws.freeze_panes = "A9"

# =============================================================
# SHEET 6: FS_Disclosure_MBB
# =============================================================
def build_fs_disclosure(ws, label, amort_sheet, principal_col, opening_bal_lookup_cell):
    ws[f"A1"] = f"FS DISCLOSURE — {label}"
    ws["A1"].font = HDR_FONT; ws["A1"].fill = HDR_FILL
    ws.merge_cells("A1:D1")
    ws["A3"] = "Facility:"
    ws["B3"] = f"={amort_sheet}!B3"
    ws["A4"] = "Reporting FYE:"
    ws["B4"] = f"={amort_sheet}!E3"
    ws["B4"].number_format = DATE_FMT

    ws["A6"] = "OUTSTANDING — MFRS 101"
    ws["A6"].font = SUB_FONT; ws["A6"].fill = SUB_FILL
    ws.merge_cells("A6:D6")
    ws["A7"] = "Closing principal at FYE (RM)"
    # Sum of all principal payments due after FYE = principal outstanding at FYE
    ws["B7"] = f'=SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!B:B,">"&B4)'
    ws["B7"].number_format = NUM_FMT
    ws["A8"] = "Current portion (next 12 mths) (RM)"
    ws["B8"] = f'=SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!J:J,"Current")'
    ws["B8"].number_format = NUM_FMT
    ws["A9"] = "Non-current portion (RM)"
    ws["B9"] = (f'=SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!J:J,"1-2 yrs")'
                f'+SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!J:J,"2-5 yrs")'
                f'+SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!J:J,"Over 5 yrs")')
    ws["B9"].number_format = NUM_FMT
    ws["A10"] = "Sub-total (current + non-current)"
    ws["B10"] = "=B8+B9"
    ws["B10"].number_format = NUM_FMT; ws["B10"].fill = TOTAL_FILL
    ws["A11"] = "Variance vs closing balance"
    ws["B11"] = "=B7-B10"
    ws["B11"].number_format = NUM_FMT

    ws["A13"] = "MATURITY PROFILE — MFRS 7"
    ws["A13"].font = SUB_FONT; ws["A13"].fill = SUB_FILL
    ws.merge_cells("A13:D13")
    profile = [
        ("Within 1 year (RM)", "Current"),
        ("1 to 2 years (RM)", "1-2 yrs"),
        ("2 to 5 years (RM)", "2-5 yrs"),
        ("Over 5 years (RM)", "Over 5 yrs"),
    ]
    for i, (lbl, b) in enumerate(profile, start=14):
        ws.cell(row=i, column=1, value=lbl)
        cell = ws.cell(row=i, column=2,
            value=f'=SUMIFS({amort_sheet}!{principal_col}:{principal_col},{amort_sheet}!J:J,"{b}")')
        cell.number_format = NUM_FMT
    ws["A18"] = "Total"
    ws["A18"].font = SUB_FONT
    ws["B18"] = "=SUM(B14:B17)"
    ws["B18"].number_format = NUM_FMT; ws["B18"].fill = TOTAL_FILL; ws["B18"].font = SUB_FONT

    ws["A20"] = "INTEREST / PROFIT P&L — MFRS 9"
    ws["A20"].font = SUB_FONT; ws["A20"].fill = SUB_FILL
    ws.merge_cells("A20:D20")
    ws["A21"] = "Year start"
    ws["B21"] = "=EDATE(B4,-12)+1"; ws["B21"].number_format = DATE_FMT
    ws["A22"] = "Charge to P&L (RM)"
    int_col = "E"
    ws["B22"] = (f'=SUMIFS({amort_sheet}!{int_col}:{int_col},'
                 f'{amort_sheet}!B:B,">="&B21,'
                 f'{amort_sheet}!B:B,"<="&B4)')
    ws["B22"].number_format = NUM_FMT

    ws["A24"] = "MOVEMENT — MFRS 107"
    ws["A24"].font = SUB_FONT; ws["A24"].fill = SUB_FILL
    ws.merge_cells("A24:D24")
    movement = [
        "Opening balance (start of year)",
        "Add: Drawdowns during year",
        "Add: Interest / profit accrued",
        "Less: Interest / profit paid",
        "Less: Principal repayments",
        "Closing balance (FYE)",
    ]
    for i, lbl in enumerate(movement, start=25):
        ws.cell(row=i, column=1, value=lbl)
        cell = ws.cell(row=i, column=2)
        cell.fill = INPUT_FILL; cell.number_format = NUM_FMT
        if "Closing" in lbl:
            cell.fill = TOTAL_FILL
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 22

ws_mbb = wb.create_sheet("FS_Disclosure_MBB")
build_fs_disclosure(ws_mbb, "MAYBANK ISLAMIC CMTF-i", "Amortisation_MBB", "G", "B4")

ws_alb = wb.create_sheet("FS_Disclosure_ALB")
build_fs_disclosure(ws_alb, "ALLIANCE BANK FIXED LOAN", "Amortisation_ALB", "F", "B4")

# Add Alliance-Bank-specific accrued interest line
ws_alb["A33"] = "ACCRUED INTEREST AT FYE (Alliance Bank specific)"
ws_alb["A33"].font = SUB_FONT; ws_alb["A33"].fill = SUB_FILL
ws_alb.merge_cells("A33:D33")
ws_alb["A34"] = "Accrued interest payable (RM)"
ws_alb["B34"] = "=Amortisation_ALB!E5"
ws_alb["B34"].number_format = NUM_FMT
ws_alb["A35"] = "  (reclass from loan principal balance to other payables)"
ws_alb["A35"].font = Font(italic=True, size=9, color="595959")

# =============================================================
# SHEET 8: FS_Aggregate
# =============================================================
ws = wb.create_sheet("FS_Aggregate")
ws["A1"] = "GROUP / CONSOLIDATED BORROWINGS — DISCLOSURE AGGREGATE"
ws["A1"].font = HDR_FONT; ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:E1")

ws["A3"] = "Note: aggregate assumes same FYE across facilities. Adjust if FYEs differ."
ws["A3"].font = Font(italic=True, color="595959", size=9)
ws.merge_cells("A3:E3")

agg_hdr = ["Line Item", "Maybank Islamic", "Alliance Bank", "TOTAL", "Comment"]
for c, h in enumerate(agg_hdr, start=1):
    cell = ws.cell(row=5, column=c, value=h)
    cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER

agg_rows = [
    ("Closing principal at FYE",       "=FS_Disclosure_MBB!B7", "=FS_Disclosure_ALB!B7"),
    ("Current portion",                "=FS_Disclosure_MBB!B8", "=FS_Disclosure_ALB!B8"),
    ("Non-current portion",            "=FS_Disclosure_MBB!B9", "=FS_Disclosure_ALB!B9"),
    ("",                               None, None),
    ("Maturity: within 1 yr",          "=FS_Disclosure_MBB!B14", "=FS_Disclosure_ALB!B14"),
    ("Maturity: 1-2 yrs",              "=FS_Disclosure_MBB!B15", "=FS_Disclosure_ALB!B15"),
    ("Maturity: 2-5 yrs",              "=FS_Disclosure_MBB!B16", "=FS_Disclosure_ALB!B16"),
    ("Maturity: over 5 yrs",           "=FS_Disclosure_MBB!B17", "=FS_Disclosure_ALB!B17"),
    ("",                               None, None),
    ("Interest / profit P&L",          "=FS_Disclosure_MBB!B22", "=FS_Disclosure_ALB!B22"),
    ("Accrued interest payable @ FYE", 0, "=FS_Disclosure_ALB!B34"),
]
for i, (lbl, mbb, alb) in enumerate(agg_rows, start=6):
    ws.cell(row=i, column=1, value=lbl)
    if mbb is not None:
        c = ws.cell(row=i, column=2, value=mbb); c.number_format = NUM_FMT
    if alb is not None:
        c = ws.cell(row=i, column=3, value=alb); c.number_format = NUM_FMT
    if mbb is not None and alb is not None:
        c = ws.cell(row=i, column=4, value=f"=B{i}+C{i}")
        c.number_format = NUM_FMT; c.fill = TOTAL_FILL
    for col in range(1, 6):
        ws.cell(row=i, column=col).border = BORDER

ws.column_dimensions['A'].width = 38
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 20
ws.column_dimensions['E'].width = 35

# =============================================================
# SHEET 9: Ref_TxnCodes
# =============================================================
ws = wb.create_sheet("Ref_TxnCodes")
ws["A1"] = "STATEMENT TRANSACTION CODE DICTIONARY"
ws["A1"].font = HDR_FONT; ws["A1"].fill = HDR_FILL
ws.merge_cells("A1:D1")

# Maybank section
ws["A3"] = "MAYBANK ISLAMIC (CMTF-i)"
ws["A3"].font = SUB_FONT; ws["A3"].fill = SUB_FILL
ws.merge_cells("A3:D3")
mbb_codes = [
    ("Code Pattern", "Meaning", "DR/CR Effect", "Treatment"),
    ("0xxx/N",        "DRAWDOWN DEBIT",          "Increases outstanding (DR)", "Add to principal — log in Drawdown_Log"),
    ("3xxx/",         "PROGRESSIVE PROFIT",      "Increases outstanding (DR)", "Profit accrual — flows to P&L"),
    ("0xxx/ BRH...",  "Settlement (profit/principal)", "Reduces outstanding (CR)", "Split profit-paid vs principal-paid"),
    ("0410/",         "Initial drawdown CMTF-i 1", "DR", "First drawdown row"),
    ("0402/, 1003/",  "CMTF-i 2 progressive drawdowns", "DR", "Multi-tranche"),
    ("0605/, 3006/",  "CMTF-i 3 progressive drawdowns", "DR", ""),
    ("2807/, 0309/",  "CMTF-i 4 progressive drawdowns", "DR", ""),
    ("0810/",         "CMTF-i 5 progressive drawdown",  "DR", ""),
    ("BRH000003",     "Branch ref CMTF-i 1-3 settlements", "n/a", "Match settlement to facility"),
    ("BRH000005",     "Branch ref CMTF-i 4-5 settlements", "n/a", ""),
    ("Negative balance", "Bank's books — borrower liability", "n/a", "Sign-flip when normalising for FS"),
]
for r, row in enumerate(mbb_codes, start=4):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=v); cell.border = BORDER
        if r == 4:
            cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER

# Alliance Bank section
sec_row = 4 + len(mbb_codes) + 1
ws.cell(row=sec_row, column=1, value="ALLIANCE BANK (Fixed Loans - Others)")
ws.cell(row=sec_row, column=1).font = SUB_FONT
ws.cell(row=sec_row, column=1).fill = SUB_FILL
ws.merge_cells(start_row=sec_row, start_column=1, end_row=sec_row, end_column=4)

alb_codes = [
    ("Code / Description", "Meaning", "DR/CR Effect", "Treatment"),
    ("BEGINNING BALANCE", "Statement opening balance", "DR shown",
     "Use as schedule starting principal"),
    ("LOAN RELEASE",      "Drawdown event",            "DR",
     "Log in Drawdown_Log"),
    ("*DR INTEREST BILLED","Month-end interest charge","DR",
     "Interest expense — flows to P&L on accrual"),
    ("*CR INT (AUTO)",    "Auto-deduction of interest from current account next day", "CR",
     "Interest paid — clears the accrual"),
    ("*CR PRIN (AUTO)",   "Auto-deduction of fixed principal", "CR",
     "Principal repayment — reduces loan principal"),
    ("DR balance",        "Debit balance = borrower liability", "n/a",
     "Already positive convention; no sign flip"),
    ("Note",              "Closing balance includes last interest billed but not yet auto-deducted", "n/a",
     "Reclass that amount to 'accrued interest payable' for FS"),
]
for r_off, row in enumerate(alb_codes):
    r = sec_row + 1 + r_off
    for c, v in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=v); cell.border = BORDER
        if r_off == 0:
            cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER

for col, w in zip("ABCD", [26, 42, 32, 50]):
    ws.column_dimensions[col].width = w

# Save
wb.save(OUT)
print(f"Saved: {OUT}")
print(f"Sheets: {wb.sheetnames}")
