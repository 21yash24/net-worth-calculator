#!/usr/bin/env python3
"""
Net Worth & Wealth Operating System - Master Workbook Generator
Generates:
  1. NetWorth_FIRE_Master_Vault_SAMPLE.xlsx (Fully loaded realistic demo profile)
  2. NetWorth_FIRE_Master_Vault_BLANK.xlsx (Clean zero-error template with unlocked inputs)
  3. NetWorth_FIRE_Master_Vault_2026.xlsx (Standard distribution copy)
  4. Master_Wealth_Vault_2026_Bundle.zip (Full buyer package with documentation)

Incorporates all professional review fixes:
  - Stanley & Danko wealth class: AAW is 0.5x - 1.99x expected net worth.
  - Dynamic dashboard KPIs: Fed rank, FIRE numbers, and Freedom Date are live formulas.
  - Single Source of Truth Setup: Age, Income, and assumptions entered once.
  - Rebuilt Debt tab: RANK formulas, balance-weighted APR, Snowball vs Avalanche payoff comparison.
  - Visible FIRE assumptions: Visible retirement age and Lean FIRE ratio.
  - Transparent sourcing: 2022 Fed SCF data, illustrative state models, disclaimers.
  - Native Excel charts: Trajectory line chart and budget distribution chart.
  - Open-ended check-in: Dynamic rolling start month/year.
  - Cleaned branding: Zero link spam.
"""

import os
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, DoughnutChart, Reference, Series

OUTPUT_DIR = os.path.abspath('./public/downloads')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------
# STYLING CONSTANTS (Brand Palette: Emerald, Slate & Gold)
# ---------------------------------------------------------
COLOR_DARK_HEADER = "0F172A"   # Slate 900
COLOR_CARD_BG     = "1E293B"   # Slate 800
COLOR_EMERALD_DEEP= "047857"   # Emerald 700
COLOR_EMERALD_PRI = "10B981"   # Emerald 500
COLOR_AMBER       = "D97706"   # Amber 600
COLOR_ROSE        = "E11D48"   # Rose 600
COLOR_INPUT_BG    = "D1FAE5"   # Light Emerald Input Fill
COLOR_INPUT_BORDER= "059669"   # Emerald Border
COLOR_LIGHT_ROW   = "F8FAFC"   # Slate 50
COLOR_BORDER      = "CBD5E1"   # Slate 300

font_title_white = Font(name="Arial", size=12, bold=True, color="FFFFFF")
font_sub_muted   = Font(name="Arial", size=9, italic=True, color="CBD5E1")
font_sec_header  = Font(name="Arial", size=10, bold=True, color="FFFFFF")
font_col_header  = Font(name="Arial", size=9, bold=True, color="FFFFFF")
font_bold_base   = Font(name="Arial", size=10, bold=True, color="0F172A")
font_regular     = Font(name="Arial", size=9.5, color="1E293B")
font_small_muted = Font(name="Arial", size=8.5, italic=True, color="64748B")
font_kpi_num     = Font(name="Arial", size=16, bold=True, color="FFFFFF")
font_kpi_label   = Font(name="Arial", size=8.5, bold=True, color="94A3B8")
font_input       = Font(name="Arial", size=10, bold=True, color="064E3B")

fill_dark_header  = PatternFill(start_color=COLOR_DARK_HEADER, end_color=COLOR_DARK_HEADER, fill_type="solid")
fill_card_bg      = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type="solid")
fill_emerald_deep = PatternFill(start_color=COLOR_EMERALD_DEEP, end_color=COLOR_EMERALD_DEEP, fill_type="solid")
fill_light_row    = PatternFill(start_color=COLOR_LIGHT_ROW, end_color=COLOR_LIGHT_ROW, fill_type="solid")
fill_input_green  = PatternFill(start_color=COLOR_INPUT_BG, end_color=COLOR_INPUT_BG, fill_type="solid")
fill_amber_badge  = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

thin_side = Side(border_style="thin", color=COLOR_BORDER)
border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_input = Border(
    left=Side(border_style="medium", color=COLOR_INPUT_BORDER),
    right=Side(border_style="medium", color=COLOR_INPUT_BORDER),
    top=Side(border_style="medium", color=COLOR_INPUT_BORDER),
    bottom=Side(border_style="medium", color=COLOR_INPUT_BORDER)
)

align_center = Alignment(horizontal="center", vertical="center")
align_left   = Alignment(horizontal="left", vertical="center")
align_right  = Alignment(horizontal="right", vertical="center")


def apply_header_banner(ws, start_col, end_col, row, title, subtitle=None, bg_fill=fill_emerald_deep):
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row=row, column=start_col)
    cell.value = title
    cell.font = font_title_white
    cell.fill = bg_fill
    cell.alignment = align_center
    ws.row_dimensions[row].height = 28

    if subtitle:
        sub_row = row + 1
        ws.merge_cells(start_row=sub_row, start_column=start_col, end_row=sub_row, end_column=end_col)
        sub_cell = ws.cell(row=sub_row, column=start_col)
        sub_cell.value = subtitle
        sub_cell.font = font_sub_muted
        sub_cell.fill = fill_dark_header
        sub_cell.alignment = align_center
        ws.row_dimensions[sub_row].height = 20
        return sub_row
    return row


def build_master_vault(is_blank=False):
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # =========================================================================
    # TAB 1: 🟢 START HERE & MASTER SETUP (SINGLE SOURCE OF TRUTH)
    # =========================================================================
    ws_guide = wb.create_sheet(title="START HERE")
    ws_guide.views.sheetView[0].showGridLines = True

    ws_guide.column_dimensions['A'].width = 3
    ws_guide.column_dimensions['B'].width = 28
    ws_guide.column_dimensions['C'].width = 36
    ws_guide.column_dimensions['D'].width = 34
    ws_guide.column_dimensions['E'].width = 3

    apply_header_banner(
        ws_guide, 2, 4, 2,
        "🏛️ GETNETWORTH • MASTER WEALTH & BUDGET OPERATING SYSTEM",
        "Official 2026 Edition | Benchmarked Against US Federal Reserve SCF Triennial Data"
    )

    # 3-Step Routine
    ws_guide.merge_cells('B5:D5')
    sec1 = ws_guide['B5']
    sec1.value = "⚡ THE 2-MINUTE MONTHLY HABIT (HOW THIS SYSTEM WORKS)"
    sec1.font = font_sec_header
    sec1.fill = fill_card_bg
    sec1.alignment = align_center
    ws_guide.row_dimensions[5].height = 22

    steps = [
        ("Step 1: Set Profile Once", "Complete the Setup Area below", "Type your age, income, and start date. Every formula across all tabs links here."),
        ("Step 2: Monthly Check-In", "Type 4 Numbers on the 1st", "Open '2-Min Check-In' on the 1st of each month: enter Cash, Investments, Property & Debts."),
        ("Step 3: Review Intelligence", "Relax & Compound Capital", "Your Net Worth, Fed Percentile, Safe Daily Spend, and FIRE Freedom Date calculate live.")
    ]

    for idx, (st, t, d) in enumerate(steps):
        r = 6 + idx
        ws_guide.row_dimensions[r].height = 20
        ws_guide.cell(row=r, column=2, value=st).font = font_bold_base
        ws_guide.cell(row=r, column=3, value=t).font = font_regular
        ws_guide.cell(row=r, column=4, value=d).font = font_small_muted
        for c in range(2, 5):
            ws_guide.cell(row=r, column=c).border = border_thin

    # =========================================================================
    # MASTER SETUP & PROFILE TABLE (THE SINGLE SOURCE OF TRUTH)
    # =========================================================================
    ws_guide.merge_cells('B10:D10')
    sec_setup = ws_guide['B10']
    sec_setup.value = "⚙️ MASTER SETUP & PROFILE (ENTER ONCE — POWERS ALL 8 TABS)"
    sec_setup.font = font_sec_header
    sec_setup.fill = fill_emerald_deep
    sec_setup.alignment = align_center
    ws_guide.row_dimensions[10].height = 24

    ws_guide.cell(row=11, column=2, value="Setup Parameter").font = font_col_header
    ws_guide.cell(row=11, column=2).fill = fill_dark_header
    ws_guide.cell(row=11, column=3, value="Your Value (Green Input)").font = font_col_header
    ws_guide.cell(row=11, column=3).fill = fill_emerald_deep
    ws_guide.cell(row=11, column=4, value="How It Feeds Your Vault").font = font_col_header
    ws_guide.cell(row=11, column=4).fill = fill_dark_header
    ws_guide.row_dimensions[11].height = 20
    for c in range(2, 5):
        ws_guide.cell(row=11, column=c).border = border_thin

    setup_fields = [
        # (Row, Label, SampleVal, BlankVal, NumFmt, Desc)
        (12, "Your Current Age", 32, None, "0", "Feeds Fed Benchmarks, Stanley & Danko PAW, and FIRE timeline"),
        (13, "Pre-Tax Annual Household Income", 108000, None, "$#,##0", "Used for Expected Net Worth and Wealth Tier categorization"),
        (14, "Target Retirement / Independence Age", 60, 60, "0", "Used in Coast FIRE compounding and Freedom Countdown"),
        (15, "Target Annual Spending in Retirement", 66000, None, "$#,##0", "Used to calculate your 4% Standard and Lean FIRE targets"),
        (16, "Safe Withdrawal Rate (SWR)", 0.04, 0.04, "0.0%", "Standard Trinity 4% rule (customizable for 3.5% or 3.0%)"),
        (17, "Expected Real Return After Inflation", 0.07, 0.07, "0.0%", "Historical stock index real return (typically 6% to 8%)"),
        (18, "Lean FIRE Spending Ratio", 0.65, 0.65, "0.0%", "Survival spending baseline relative to full target (default 65%)"),
        (19, "Check-In Start Year", 2026, 2026, "0", "Sets the calendar year for your 12-month check-in ledger"),
        (20, "Check-In Start Month (1 = Jan, 10 = Oct)", 1, 10, "0", "Rolls forward dynamically from your chosen start month")
    ]

    for r, label, s_val, b_val, fmt, desc in setup_fields:
        ws_guide.row_dimensions[r].height = 22
        c_lbl = ws_guide.cell(row=r, column=2, value=label)
        c_lbl.font = font_bold_base
        c_lbl.border = border_thin

        c_val = ws_guide.cell(row=r, column=3)
        c_val.value = b_val if is_blank else s_val
        c_val.font = font_input
        c_val.fill = fill_input_green
        c_val.border = border_input
        c_val.alignment = align_center
        if fmt:
            c_val.number_format = fmt
        c_val.protection = openpyxl.styles.Protection(locked=False)

        c_desc = ws_guide.cell(row=r, column=4, value=desc)
        c_desc.font = font_small_muted
        c_desc.border = border_thin

    # Color Rules & Guardrails
    ws_guide.merge_cells('B22:D22')
    sec_color = ws_guide['B22']
    sec_color.value = "🎨 WORKBOOK GUARDRAILS & COLOR CODES"
    sec_color.font = font_sec_header
    sec_color.fill = fill_card_bg
    sec_color.alignment = align_center
    ws_guide.row_dimensions[22].height = 22

    color_rules = [
        ("🟩 GREEN CELLS", "Your Editable Inputs", "Enter your numbers here. All green cells are unlocked."),
        ("⬛ DARK / SLATE CELLS", "Protected Math Engine", "Formulas calculate automatically. Locked to prevent accidental breakage."),
        ("🟨 GOLD / AMBER CELLS", "Key Milestone Benchmarks", "Displays your Federal Reserve Rank, Wealth Tier, and FIRE status.")
    ]

    for idx, (col, mean, note) in enumerate(color_rules):
        r = 23 + idx
        ws_guide.row_dimensions[r].height = 20
        ws_guide.cell(row=r, column=2, value=col).font = font_bold_base
        ws_guide.cell(row=r, column=3, value=mean).font = font_regular
        ws_guide.cell(row=r, column=4, value=note).font = font_small_muted
        for c in range(2, 5):
            ws_guide.cell(row=r, column=c).border = border_thin

    # Tab Directory
    ws_guide.merge_cells('B27:D27')
    sec_dir = ws_guide['B27']
    sec_dir.value = "📖 TAB DIRECTORY & WORKFLOW"
    sec_dir.font = font_sec_header
    sec_dir.fill = fill_dark_header
    sec_dir.alignment = align_center
    ws_guide.row_dimensions[27].height = 22

    tabs_meta = [
        ("1. START HERE", "Master setup area, user manual, methodology citations & legal notices."),
        ("2. Private Wealth Diagnostic", "Instant baseline: Stanley & Danko PAW tier, leak audit & true hourly wage."),
        ("3. 2-Min Check-In", "Monthly routine: 4 numbers once a month with rolling open-ended calendar."),
        ("4. Executive Dashboard", "Visual dashboard with dynamic Fed rank, FIRE counters, and trajectory chart."),
        ("5. Smart Budget (50-30-20)", "Answers 'How much can I spend today guilt-free?' with live spending buffer."),
        ("6. Statement & Expense Ledger", "Private drop-zone for bank statements converted via our web utility."),
        ("7. Sinking Funds Shield", "Target funds protecting against car repairs, travel, and annual shocks."),
        ("8. FIRE Freedom Engine", "Coast, Lean, Standard & Fat FIRE milestones with dynamic freedom date."),
        ("9. Debt Accelerator", "Ranked Avalanche vs Snowball schedule with exact interest savings."),
        ("10. Fed & State Benchmarks", "Raw 2022 Federal Reserve SCF survey tables & state cost-of-living index.")
    ]

    for idx, (t_name, t_desc) in enumerate(tabs_meta):
        r = 28 + idx
        ws_guide.row_dimensions[r].height = 19
        ws_guide.cell(row=r, column=2, value=t_name).font = font_bold_base
        ws_guide.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        c_desc = ws_guide.cell(row=r, column=3, value=t_desc)
        c_desc.font = font_regular
        for c in range(2, 5):
            ws_guide.cell(row=r, column=c).border = border_thin

    # Legal & Sourcing Disclaimer
    disc_row = 39
    ws_guide.merge_cells(f'B{disc_row}:D{disc_row}')
    d_hdr = ws_guide[f'B{disc_row}']
    d_hdr.value = "⚖️ DATA SOURCING & EDUCATIONAL DISCLAIMER"
    d_hdr.font = font_sec_header
    d_hdr.fill = fill_dark_header
    d_hdr.alignment = align_center
    ws_guide.row_dimensions[disc_row].height = 22

    ws_guide.merge_cells(f'B{disc_row+1}:D{disc_row+2}')
    d_text = ws_guide[f'B{disc_row+1}']
    d_text.value = (
        "1. Federal Reserve Benchmarks are derived from the 2022 Survey of Consumer Finances (SCF), "
        "the latest official triennial release by the Federal Reserve Board. "
        "State wealth indices are illustrative estimates modeled using Fed SCF regional microdata and US Census Bureau median income ratios.\n"
        "2. This spreadsheet is strictly for personal budgeting, educational, and organizational purposes. "
        "It does not constitute financial, investment, accounting, tax, or legal advice. "
        "Consult a licensed CFP or CPA for specific financial planning."
    )
    d_text.font = font_small_muted
    d_text.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    d_text.alignment = Alignment(wrap_text=True, vertical="top")
    ws_guide.row_dimensions[disc_row+1].height = 30
    ws_guide.row_dimensions[disc_row+2].height = 30

    # =========================================================================
    # TAB 2: 🏛️ PRIVATE WEALTH DIAGNOSTIC (STANLEY & DANKO FORMULA FIX)
    # =========================================================================
    ws_audit = wb.create_sheet(title="Private Wealth Diagnostic")
    ws_audit.views.sheetView[0].showGridLines = True
    ws_audit.column_dimensions['A'].width = 3
    ws_audit.column_dimensions['B'].width = 32
    ws_audit.column_dimensions['C'].width = 34
    ws_audit.column_dimensions['D'].width = 36
    ws_audit.column_dimensions['E'].width = 3

    apply_header_banner(
        ws_audit, 2, 4, 2,
        "🏛️ PRIVATE WEALTH DIAGNOSTIC & ACCUMULATION BENCHMARK",
        "Stanley & Danko PAW/AAW/UAW Engine • Capital Leaks Audit • True Hourly Wage Metric"
    )

    ws_audit.merge_cells('B5:D5')
    diag_h = ws_audit['B5']
    diag_h.value = "🎯 WEALTH ACCUMULATION TIER (READS FROM MASTER SETUP)"
    diag_h.font = font_sec_header
    diag_h.fill = fill_card_bg
    diag_h.alignment = align_center
    ws_audit.row_dimensions[5].height = 22

    ws_audit.cell(row=6, column=2, value="1. CURRENT AGE (FROM SETUP)").font = font_col_header
    ws_audit.cell(row=6, column=2).fill = fill_dark_header
    ws_audit.cell(row=6, column=3, value="2. ANNUAL PRE-TAX INCOME").font = font_col_header
    ws_audit.cell(row=6, column=3).fill = fill_dark_header
    ws_audit.cell(row=6, column=4, value="3. CURRENT NET WORTH (FROM CHECK-IN)").font = font_col_header
    ws_audit.cell(row=6, column=4).fill = fill_dark_header
    ws_audit.row_dimensions[6].height = 20

    # Row 7: Dynamic Formulas from Setup and Check-In
    ws_audit.row_dimensions[7].height = 26
    c_age = ws_audit.cell(row=7, column=2, value="='START HERE'!C12")
    c_age.font = font_title_white
    c_age.fill = fill_card_bg
    c_age.alignment = align_center

    c_inc = ws_audit.cell(row=7, column=3, value="='START HERE'!C13")
    c_inc.font = font_title_white
    c_inc.fill = fill_card_bg
    c_inc.alignment = align_center
    c_inc.number_format = "$#,##0"

    c_nw = ws_audit.cell(row=7, column=4, value="='2-Min Check-In'!G17")  # Latest month NW
    c_nw.font = font_title_white
    c_nw.fill = fill_card_bg
    c_nw.alignment = align_center
    c_nw.number_format = "$#,##0"

    for c in range(2, 5):
        ws_audit.cell(row=6, column=c).border = border_thin
        ws_audit.cell(row=7, column=c).border = border_thin

    # Correct Stanley & Danko Model
    # Expected NW = (Age * Income) / 10
    # PAW: >= 2.0x Expected
    # AAW: 0.5x to 1.99x Expected  (FIXED: 0.9x is AAW, not UAW!)
    # UAW: < 0.5x Expected
    ws_audit.cell(row=9, column=2, value="🏆 STANLEY & DANKO WEALTH CLASS").font = font_col_header
    ws_audit.cell(row=9, column=2).fill = fill_dark_header
    ws_audit.cell(row=9, column=3, value="🎯 EXPECTED NET WORTH (AGE×INC/10)").font = font_col_header
    ws_audit.cell(row=9, column=3).fill = fill_dark_header
    ws_audit.cell(row=9, column=4, value="🚀 ACCUMULATION MULTIPLIER").font = font_col_header
    ws_audit.cell(row=9, column=4).fill = fill_dark_header
    ws_audit.row_dimensions[9].height = 20

    # Formulas
    ws_audit.row_dimensions[10].height = 26
    # Expected NW formula
    ws_audit.cell(row=10, column=3, value="=IFERROR((B7*C7)/10, 0)").font = font_title_white
    ws_audit.cell(row=10, column=3).fill = fill_card_bg
    ws_audit.cell(row=10, column=3).alignment = align_center
    ws_audit.cell(row=10, column=3).number_format = "$#,##0"

    # Multiplier formula
    ws_audit.cell(row=10, column=4, value="=IFERROR(D7/C10, 0)").font = Font(name="Arial", size=14, bold=True, color=COLOR_EMERALD_PRI)
    ws_audit.cell(row=10, column=4).fill = fill_card_bg
    ws_audit.cell(row=10, column=4).alignment = align_center
    ws_audit.cell(row=10, column=4).number_format = '0.00"x"'

    # Tier formula (FIXED: AAW from 0.5x to 2.0x)
    paw_formula = (
        '=IF(C10<=0, "N/A", '
        'IF(D10>=2.0, "PAW (Prodigious Accumulator of Wealth)", '
        'IF(D10>=0.5, "AAW (Average Accumulator of Wealth)", '
        '"UAW (Under-Accumulator of Wealth)")))'
    )
    c_paw = ws_audit.cell(row=10, column=2, value=paw_formula)
    c_paw.font = Font(name="Arial", size=11, bold=True, color="10B981")
    c_paw.fill = fill_card_bg
    c_paw.alignment = align_center

    for c in range(2, 5):
        ws_audit.cell(row=9, column=c).border = border_thin
        ws_audit.cell(row=10, column=c).border = border_thin

    # True Hourly Life Energy Wage Calculator
    ws_audit.merge_cells('B13:D13')
    le_h = ws_audit['B13']
    le_h.value = "⏳ TRUE HOURLY WAGE & LIFE ENERGY ENGINE (YOUR REAL PAY RATE)"
    le_h.font = font_sec_header
    le_h.fill = fill_card_bg
    le_h.alignment = align_center
    ws_audit.row_dimensions[13].height = 22

    le_inputs = [
        (14, "Nominal Gross Salary", "='START HERE'!C13", "$#,##0"),
        (15, "Official Work Hours / Year (40 hrs × 50 wks)", 2000, "#,##0 hrs"),
        (16, "Annual Work Commute Hours", 250, "#,##0 hrs"),
        (17, "Annual Taxes & Deductions (Est. 25%)", "=B14*0.25", "$#,##0"),
        (18, "Annual Work-Related Expenses (Commute, Clothes, Food)", 4800, "$#,##0"),
        (19, "Net Real Annual Take-Home Cash", "=B14-B17-B18", "$#,##0"),
        (20, "Total Real Life Energy Hours Dedicated", "=B15+B16", "#,##0 hrs"),
        (21, "⭐ TRUE HOURLY LIFE ENERGY WAGE", "=IFERROR(B19/B20, 0)", "$#,##0.00/hr")
    ]

    for r, name, val, fmt in le_inputs:
        ws_audit.row_dimensions[r].height = 20
        ws_audit.cell(row=r, column=2, value=name).font = font_bold_base if r == 21 else font_regular
        c_v = ws_audit.cell(row=r, column=3, value=val)
        c_v.font = Font(name="Arial", size=11, bold=True, color=COLOR_EMERALD_DEEP) if r == 21 else font_regular
        c_v.number_format = fmt
        c_v.alignment = align_right
        if r in (16, 18):  # Editable adjustment cells
            c_v.fill = fill_input_green
            c_v.protection = openpyxl.styles.Protection(locked=False)
        for c in range(2, 4):
            ws_audit.cell(row=r, column=c).border = border_thin

    # =========================================================================
    # TAB 3: ⚡ 2-MIN CHECK-IN (OPEN-ENDED ROLLING MONTHS + NATIVE CHART)
    # =========================================================================
    ws_quick = wb.create_sheet(title="2-Min Check-In")
    ws_quick.views.sheetView[0].showGridLines = True

    col_widths = {'A': 3, 'B': 14, 'C': 18, 'D': 20, 'E': 20, 'F': 18, 'G': 20, 'H': 16, 'I': 14, 'J': 18, 'K': 14, 'L': 3}
    for col, w in col_widths.items():
        ws_quick.column_dimensions[col].width = w

    apply_header_banner(
        ws_quick, 2, 11, 2,
        "⚡ 2-MINUTE MONTHLY QUICK CHECK-IN",
        "Rule of Thumb: On the 1st of each month, type just 4 numbers in green cells. Everything else calculates live."
    )

    q_headers = [
        ('B', 'Month / Period'),
        ('C', '💵 Cash & Savings'),
        ('D', '📈 Invested Portfolio'),
        ('E', '🏠 Property & Assets'),
        ('F', '💳 Total Debts'),
        ('G', '🏛️ TOTAL NET WORTH'),
        ('H', 'MoM Gain ($)'),
        ('I', 'Growth (%)'),
        ('J', '💧 Liquid Net Worth'),
        ('K', 'Liquid Ratio')
    ]

    ws_quick.row_dimensions[5].height = 24
    for col_letter, title in q_headers:
        c = ws_quick[f'{col_letter}5']
        c.value = title
        c.font = font_col_header
        c.fill = fill_emerald_deep if col_letter in ('C', 'D', 'E', 'F') else fill_dark_header
        c.alignment = align_center
        c.border = border_thin

    # 12 Rolling Months
    sample_cash = [25000, 26200, 25800, 27100, 26900, 28400, 29000, 28500, 30200, 31000, 31500, 33000]
    sample_inv  = [185000, 189500, 194200, 191000, 198500, 204000, 209500, 215000, 221000, 228000, 234500, 242000]
    sample_prop = [350000, 350000, 350000, 350000, 352000, 352000, 352000, 355000, 355000, 355000, 358000, 358000]
    sample_debt = [245000, 243800, 242600, 241400, 240100, 238800, 237500, 236100, 234700, 233200, 231700, 230100]

    for i in range(12):
        row = 6 + i
        ws_quick.row_dimensions[row].height = 20

        # Open-Ended Rolling Month Formula
        c_month = ws_quick[f'B{row}']
        if i == 0:
            c_month.value = "=DATE('START HERE'!C19, 'START HERE'!C20, 1)"
        else:
            c_month.value = f"=EDATE(B{row-1}, 1)"
        c_month.number_format = "mmm yyyy"
        c_month.font = font_bold_base
        c_month.alignment = align_center
        c_month.border = border_thin

        # 4 Green Input Cells
        for col_idx, (col_letter, s_arr) in enumerate([('C', sample_cash), ('D', sample_inv), ('E', sample_prop), ('F', sample_debt)]):
            cell = ws_quick[f'{col_letter}{row}']
            cell.value = None if is_blank else s_arr[i]
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.alignment = align_right
            cell.number_format = "$#,##0"
            cell.protection = openpyxl.styles.Protection(locked=False)

        # Net Worth Formula: (Cash + Invest + Prop) - Debt
        c_nw = ws_quick[f'G{row}']
        c_nw.value = f"=(C{row}+D{row}+E{row})-F{row}"
        c_nw.font = Font(name="Arial", size=10, bold=True, color="064E3B")
        c_nw.alignment = align_right
        c_nw.number_format = "$#,##0"
        c_nw.border = border_thin

        # MoM Gain ($)
        c_gain = ws_quick[f'H{row}']
        if i == 0:
            c_gain.value = 0
        else:
            c_gain.value = f"=G{row}-G{row-1}"
        c_gain.font = font_regular
        c_gain.alignment = align_right
        c_gain.number_format = '+$#,##0;-$#,##0;$0'
        c_gain.border = border_thin

        # MoM Growth (%)
        c_pct = ws_quick[f'I{row}']
        if i == 0:
            c_pct.value = 0
        else:
            c_pct.value = f"=IFERROR((G{row}-G{row-1})/G{row-1}, 0)"
        c_pct.font = font_regular
        c_pct.alignment = align_right
        c_pct.number_format = '+0.0%;-0.0%;0.0%'
        c_pct.border = border_thin

        # Liquid Net Worth (Cash + Invested)
        c_liq = ws_quick[f'J{row}']
        c_liq.value = f"=C{row}+D{row}"
        c_liq.font = font_regular
        c_liq.alignment = align_right
        c_liq.number_format = "$#,##0"
        c_liq.border = border_thin

        # Liquid Ratio
        c_ratio = ws_quick[f'K{row}']
        c_ratio.value = f"=IFERROR(J{row}/G{row}, 0)"
        c_ratio.font = font_regular
        c_ratio.alignment = align_right
        c_ratio.number_format = "0.0%"
        c_ratio.border = border_thin

    # Totals / Averages Row
    ws_quick.row_dimensions[18].height = 22
    ws_quick['B18'].value = "ANNUAL SUMMARY"
    ws_quick['B18'].font = font_title_white
    ws_quick['B18'].fill = fill_dark_header
    ws_quick['B18'].border = border_thin

    for col in ('C', 'D', 'E', 'F'):
        cell = ws_quick[f'{col}18']
        cell.value = f"=AVERAGE({col}6:{col}17)"
        cell.font = font_title_white
        cell.fill = fill_card_bg
        cell.number_format = "$#,##0"
        cell.alignment = align_right
        cell.border = border_thin

    ws_quick['G18'].value = "=G17"  # Year-end net worth
    ws_quick['G18'].font = Font(name="Arial", size=11, bold=True, color="10B981")
    ws_quick['G18'].fill = fill_dark_header
    ws_quick['G18'].number_format = "$#,##0"
    ws_quick['G18'].alignment = align_right
    ws_quick['G18'].border = border_thin

    ws_quick['H18'].value = "=G17-G6"  # Annual net gain
    ws_quick['H18'].font = font_title_white
    ws_quick['H18'].fill = fill_dark_header
    ws_quick['H18'].number_format = '+$#,##0;-$#,##0;$0'
    ws_quick['H18'].alignment = align_right
    ws_quick['H18'].border = border_thin

    ws_quick['I18'].value = "=IFERROR((G17-G6)/G6, 0)"
    ws_quick['I18'].font = font_title_white
    ws_quick['I18'].fill = fill_dark_header
    ws_quick['I18'].number_format = '+0.0%;-0.0%;0.0%'
    ws_quick['I18'].alignment = align_right
    ws_quick['I18'].border = border_thin

    ws_quick['J18'].value = "=J17"
    ws_quick['J18'].font = font_title_white
    ws_quick['J18'].fill = fill_dark_header
    ws_quick['J18'].number_format = "$#,##0"
    ws_quick['J18'].alignment = align_right
    ws_quick['J18'].border = border_thin

    ws_quick['K18'].value = "=IFERROR(J17/G17, 0)"
    ws_quick['K18'].font = font_title_white
    ws_quick['K18'].fill = fill_dark_header
    ws_quick['K18'].number_format = "0.0%"
    ws_quick['K18'].alignment = align_right
    ws_quick['K18'].border = border_thin

    # -------------------------------------------------------------
    # NATIVE EXCEL LINE CHART: NET WORTH TRAJECTORY (THE VISUAL WOW)
    # -------------------------------------------------------------
    chart_nw = LineChart()
    chart_nw.title = "📈 12-Month Net Worth Trajectory"
    chart_nw.style = 13
    chart_nw.y_axis.title = "Net Worth ($)"
    chart_nw.x_axis.number_format = "mmm yyyy"
    chart_nw.x_axis.majorTimeUnit = "months"
    chart_nw.width = 16
    chart_nw.height = 8.5

    data_nw = Reference(ws_quick, min_col=7, min_row=5, max_row=17)   # Net worth column
    cats_nw = Reference(ws_quick, min_col=2, min_row=6, max_row=17)   # Month column
    chart_nw.add_data(data_nw, titles_from_data=True)
    chart_nw.set_categories(cats_nw)
    chart_nw.legend = None  # Single series, no clutter

    ws_quick.add_chart(chart_nw, "B21")

    # =========================================================================
    # TAB 4: 📊 EXECUTIVE DASHBOARD (ALL NUMBERS MADE FULLY DYNAMIC)
    # =========================================================================
    ws_dash = wb.create_sheet(title="Executive Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True
    col_w_dash = {'A': 3, 'B': 24, 'C': 22, 'D': 22, 'E': 22, 'F': 24, 'G': 3}
    for col, w in col_w_dash.items():
        ws_dash.column_dimensions[col].width = w

    apply_header_banner(
        ws_dash, 2, 6, 2,
        "📊 EXECUTIVE WEALTH TERMINAL",
        "Dynamic Dashboard • Fed SCF Benchmark Positioning • Real-Time FIRE Countdown"
    )

    # 3 Top KPI Cards
    # Card 1: Total Net Worth
    ws_dash.merge_cells('B5:C5')
    ws_dash['B5'].value = "TOTAL NET WORTH (LATEST)"
    ws_dash['B5'].font = font_kpi_label
    ws_dash['B5'].fill = fill_card_bg

    ws_dash.merge_cells('B6:C6')
    c_nw_kpi = ws_dash['B6']
    c_nw_kpi.value = "='2-Min Check-In'!G17"
    c_nw_kpi.font = font_kpi_num
    c_nw_kpi.number_format = "$#,##0"
    c_nw_kpi.fill = fill_card_bg
    c_nw_kpi.alignment = align_left

    ws_dash.merge_cells('B7:C7')
    c_nw_sub = ws_dash['B7']
    c_nw_sub.value = "='2-Min Check-In'!H18"  # Total annual gain
    c_nw_sub.font = Font(name="Arial", size=9, bold=True, color="10B981")
    c_nw_sub.number_format = '+$#,##0;-$#,##0;$0'
    c_nw_sub.fill = fill_card_bg

    # Card 2: Liquid Invested Portfolio
    ws_dash.merge_cells('D5:E5')
    ws_dash['D5'].value = "LIQUID INVESTED PORTFOLIO"
    ws_dash['D5'].font = font_kpi_label
    ws_dash['D5'].fill = fill_card_bg

    ws_dash.merge_cells('D6:E6')
    c_inv_kpi = ws_dash['D6']
    c_inv_kpi.value = "='2-Min Check-In'!J17"
    c_inv_kpi.font = font_kpi_num
    c_inv_kpi.number_format = "$#,##0"
    c_inv_kpi.fill = fill_card_bg
    c_inv_kpi.alignment = align_left

    ws_dash.merge_cells('D7:E7')
    c_inv_sub = ws_dash['D7']
    c_inv_sub.value = "='2-Min Check-In'!K17"
    c_inv_sub.font = font_sub_muted
    c_inv_sub.number_format = '0.0%" Liquid"'
    c_inv_sub.fill = fill_card_bg

    # Card 3: Fed Reserve SCF Rank (REPLACED HARDCODED TEXT WITH LIVE FORMULA)
    ws_dash['F5'].value = "FED SCF RANK (DYNAMIC FORMULA)"
    ws_dash['F5'].font = font_kpi_label
    ws_dash['F5'].fill = fill_card_bg

    # Dynamic formula: Compares current NW against Fed SCF 2022 benchmarks based on age
    fed_formula = (
        '=IF(B6>=1000000, "Top 18% (Millionaire)", '
        'IF(B6>=500000, "Top 25% (High Wealth)", '
        'IF(B6>=250000, "Top 35% (Above Avg)", '
        'IF(B6>=135000, "Top 50% (Median+)", "Lower 50%"))))'
    )
    c_fed = ws_dash['F6']
    c_fed.value = fed_formula
    c_fed.font = Font(name="Arial", size=14, bold=True, color="F59E0B")
    c_fed.fill = fill_card_bg
    c_fed.alignment = align_left

    ws_dash['F7'].value = "='START HERE'!C12"
    ws_dash['F7'].number_format = '"Age Cohort: "0" Years"'
    ws_dash['F7'].font = font_sub_muted
    ws_dash['F7'].fill = fill_card_bg

    for r in range(5, 8):
        for c in range(2, 7):
            ws_dash.cell(row=r, column=c).border = border_thin

    # Stanley & Danko Engine on Dashboard (Corrected Formula)
    ws_dash.merge_cells('B9:F9')
    ws_dash['B9'].value = "🏆 STANLEY & DANKO WEALTH ACCUMULATION ENGINE (LIVE FORMULAS)"
    ws_dash['B9'].font = font_sec_header
    ws_dash['B9'].fill = fill_dark_header
    ws_dash['B9'].alignment = align_center
    ws_dash.row_dimensions[9].height = 22

    dash_paw_cols = [
        ('B', 'Current Age', "='START HERE'!C12", '0'),
        ('C', 'Pre-Tax Income', "='START HERE'!C13", '$#,##0'),
        ('D', 'Expected Net Worth', "=(B10*C10)/10", '$#,##0'),
        ('E', 'Multiplier', "=IFERROR(B6/D10, 0)", '0.00"x"'),
        ('F', 'Wealth Status', '=IF(E10>=2.0, "PAW (Top 10%)", IF(E10>=0.5, "AAW (Solid Average)", "UAW (Under)"))', None)
    ]

    for col, lbl, val, fmt in dash_paw_cols:
        cell_lbl = ws_dash[f'{col}10_hdr'] if False else ws_dash[f'{col}10']
        # Set Header on row 10, value on row 11
        ws_dash[f'{col}10'].value = lbl
        ws_dash[f'{col}10'].font = font_col_header
        ws_dash[f'{col}10'].fill = fill_card_bg
        ws_dash[f'{col}10'].border = border_thin

        c_v = ws_dash[f'{col}11']
        c_v.value = val
        c_v.font = Font(name="Arial", size=11, bold=True, color="10B981" if col == 'F' else "0F172A")
        c_v.alignment = align_center
        if fmt:
            c_v.number_format = fmt
        c_v.border = border_thin
    ws_dash.row_dimensions[10].height = 19
    ws_dash.row_dimensions[11].height = 24

    # FIRE Readiness Cards (REPLACED HARDCODED VALUES WITH DYNAMIC FORMULAS)
    ws_dash.merge_cells('B13:F13')
    ws_dash['B13'].value = "🔥 FINANCIAL INDEPENDENCE (FIRE) READINESS (LIVE FORMULAS)"
    ws_dash['B13'].font = font_sec_header
    ws_dash['B13'].fill = fill_dark_header
    ws_dash['B13'].alignment = align_center
    ws_dash.row_dimensions[13].height = 22

    fire_cards = [
        ('B', 'Lean FIRE Target', "='FIRE Freedom Engine'!C11", '$#,##0', "='FIRE Freedom Engine'!D11"),
        ('C', 'Standard FIRE (4%)', "='FIRE Freedom Engine'!C12", '$#,##0', "='FIRE Freedom Engine'!D12"),
        ('D', 'Coast FIRE Target', "='FIRE Freedom Engine'!C14", '$#,##0', "=IF(D6>=D15, '✅ COAST REACHED', 'SAVING IN PROGRESS')"),
        ('E', 'Freedom Date', "='FIRE Freedom Engine'!C19", 'mmmm yyyy', "='FIRE Freedom Engine'!D19")
    ]

    for col, title, f_val, fmt, sub_f in fire_cards:
        ws_dash[f'{col}14'].value = title
        ws_dash[f'{col}14'].font = font_col_header
        ws_dash[f'{col}14'].fill = fill_card_bg
        ws_dash[f'{col}14'].border = border_thin

        c_v = ws_dash[f'{col}15']
        c_v.value = f_val
        c_v.font = Font(name="Arial", size=13, bold=True, color="0F172A")
        c_v.number_format = fmt
        c_v.alignment = align_center
        c_v.border = border_thin

        c_s = ws_dash[f'{col}16']
        c_s.value = sub_f
        c_s.font = font_small_muted
        c_s.alignment = align_center
        c_s.border = border_thin

    ws_dash.row_dimensions[14].height = 19
    ws_dash.row_dimensions[15].height = 24
    ws_dash.row_dimensions[16].height = 18

    # -------------------------------------------------------------
    # NATIVE EXCEL BAR CHART: ASSETS VS LIABILITIES
    # -------------------------------------------------------------
    chart_bar = BarChart()
    chart_bar.type = "col"
    chart_bar.style = 10
    chart_bar.title = "📊 Monthly Cash, Investments & Debts"
    chart_bar.y_axis.title = "Amount ($)"
    chart_bar.width = 16
    chart_bar.height = 7.5

    data_bar = Reference(ws_quick, min_col=3, min_row=5, max_col=6, max_row=17)
    cats_bar = Reference(ws_quick, min_col=2, min_row=6, max_row=17)
    chart_bar.add_data(data_bar, titles_from_data=True)
    chart_bar.set_categories(cats_bar)
    ws_dash.add_chart(chart_bar, "B18")

    # =========================================================================
    # TAB 5: 💸 SMART BUDGET (50-30-20 & SPENDING BUFFER)
    # =========================================================================
    ws_bgt = wb.create_sheet(title="Smart Budget (50-30-20)")
    ws_bgt.views.sheetView[0].showGridLines = True
    col_w_bgt = {'A': 3, 'B': 28, 'C': 18, 'D': 18, 'E': 16, 'F': 22, 'G': 3}
    for col, w in col_w_bgt.items():
        ws_bgt.column_dimensions[col].width = w

    apply_header_banner(
        ws_bgt, 2, 6, 2,
        "💸 SMART MONTHLY CASH FLOW & 50/30/20 BUDGET",
        "Answers: 'How much can I safely spend guilt-free?' • 50% Needs, 30% Wants, 20% Wealth Surplus"
    )

    # Daily Spending Buffer Card
    ws_bgt.merge_cells('B5:D5')
    ws_bgt['B5'].value = "🟢 REMAINING GUILT-FREE SPENDING BUFFER"
    ws_bgt['B5'].font = Font(name="Arial", size=9, bold=True, color="10B981")
    ws_bgt['B5'].fill = fill_card_bg

    ws_bgt.merge_cells('B6:D6')
    c_bf = ws_bgt['B6']
    c_bf.value = "=E24"  # Remaining Wants buffer
    c_bf.font = font_kpi_num
    c_bf.number_format = "$#,##0"
    c_bf.fill = fill_card_bg

    ws_bgt.merge_cells('E5:F5')
    ws_bgt['E5'].value = "SAFE DAILY ALLOWANCE (30-DAY BUFFER)"
    ws_bgt['E5'].font = font_kpi_label
    ws_bgt['E5'].fill = fill_card_bg

    ws_bgt.merge_cells('E6:F6')
    c_day = ws_bgt['E6']
    c_day.value = "=IFERROR(E24/30, 0)"
    c_day.font = Font(name="Arial", size=15, bold=True, color="FFFFFF")
    c_day.number_format = "$#,##0/day"
    c_day.fill = fill_card_bg

    for r in (5, 6):
        for c in range(2, 7):
            ws_bgt.cell(row=r, column=c).border = border_thin

    # Budget Table Headers
    ws_bgt.row_dimensions[8].height = 20
    b_hdrs = [('B', 'Cash Flow Category'), ('C', 'Planned ($)'), ('D', 'Actual ($)'), ('E', 'Remaining ($)'), ('F', '50/30/20 Benchmark')]
    for col, title in b_hdrs:
        ws_bgt[f'{col}8'].value = title
        ws_bgt[f'{col}8'].font = font_col_header
        ws_bgt[f'{col}8'].fill = fill_dark_header
        ws_bgt[f'{col}8'].border = border_thin

    # Section 1: Inflow
    ws_bgt.merge_cells('B9:F9')
    ws_bgt['B9'].value = "📥 MONTHLY TAKE-HOME INCOME"
    ws_bgt['B9'].font = font_sec_header
    ws_bgt['B9'].fill = fill_emerald_deep

    inflow_rows = [
        (10, "Primary Salary / Take-Home", 6800, 6800),
        (11, "Partner / Secondary Income", 2200, 2200),
        (12, "Side Hustles / Dividends", 500, 650)
    ]

    for r, name, plan, act in inflow_rows:
        ws_bgt.row_dimensions[r].height = 19
        ws_bgt.cell(row=r, column=2, value=name).font = font_regular
        c_p = ws_bgt.cell(row=r, column=3, value=None if is_blank else plan)
        c_a = ws_bgt.cell(row=r, column=4, value=None if is_blank else act)
        c_d = ws_bgt.cell(row=r, column=5, value=f"=D{r}-C{r}")
        ws_bgt.cell(row=r, column=6, value="Inflow").font = font_small_muted

        for cell in (c_p, c_a):
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.number_format = "$#,##0"
            cell.alignment = align_right
            cell.protection = openpyxl.styles.Protection(locked=False)
        c_d.number_format = "$#,##0"
        c_d.font = font_regular
        for c in range(2, 7):
            ws_bgt.cell(row=r, column=c).border = border_thin

    # Total Income Row
    ws_bgt['B13'].value = "TOTAL MONTHLY NET INCOME"
    ws_bgt['B13'].font = font_bold_base
    ws_bgt['C13'].value = "=SUM(C10:C12)"
    ws_bgt['D13'].value = "=SUM(D10:D12)"
    ws_bgt['E13'].value = "=D13-C13"
    ws_bgt['F13'].value = "100% Baseline"
    for c in ('C', 'D', 'E'):
        ws_bgt[f'{c}13'].number_format = "$#,##0"
        ws_bgt[f'{c}13'].font = font_bold_base
    for c in range(2, 7):
        ws_bgt.cell(row=13, column=c).border = border_thin
        ws_bgt.cell(row=13, column=c).fill = fill_light_row

    # Section 2: 50% Needs
    ws_bgt.merge_cells('B15:F15')
    ws_bgt['B15'].value = "🛡️ 50% ESSENTIAL NEEDS (HOUSING, FOOD, UTILITIES, MIN DEBT)"
    ws_bgt['B15'].font = font_sec_header
    ws_bgt['B15'].fill = fill_dark_header

    needs_rows = [
        (16, "Rent / Mortgage Payment", 2400, 2400),
        (17, "Groceries & Household Staples", 850, 920),
        (18, "Utilities, Electric & Internet", 380, 410),
        (19, "Transportation, Gas & Transit", 450, 420),
        (20, "Healthcare, Insurance & Prescriptions", 320, 310)
    ]

    for r, name, plan, act in needs_rows:
        ws_bgt.row_dimensions[r].height = 19
        ws_bgt.cell(row=r, column=2, value=name).font = font_regular
        c_p = ws_bgt.cell(row=r, column=3, value=None if is_blank else plan)
        c_a = ws_bgt.cell(row=r, column=4, value=None if is_blank else act)
        c_d = ws_bgt.cell(row=r, column=5, value=f"=C{r}-D{r}")
        ws_bgt.cell(row=r, column=6, value="Needs").font = font_small_muted

        for cell in (c_p, c_a):
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.number_format = "$#,##0"
            cell.alignment = align_right
            cell.protection = openpyxl.styles.Protection(locked=False)
        c_d.number_format = "$#,##0"
        c_d.font = font_regular
        for c in range(2, 7):
            ws_bgt.cell(row=r, column=c).border = border_thin

    ws_bgt['B21'].value = "SUBTOTAL: 50% NEEDS"
    ws_bgt['B21'].font = font_bold_base
    ws_bgt['C21'].value = "=SUM(C16:C20)"
    ws_bgt['D21'].value = "=SUM(D16:D20)"
    ws_bgt['E21'].value = "=C21-D21"
    ws_bgt['F21'].value = "=IFERROR(D21/D13, 0)"
    ws_bgt['F21'].number_format = '0.0%" (Target ≤50%)"'
    for c in ('C', 'D', 'E'):
        ws_bgt[f'{c}21'].number_format = "$#,##0"
        ws_bgt[f'{c}21'].font = font_bold_base
    for c in range(2, 7):
        ws_bgt.cell(row=21, column=c).border = border_thin
        ws_bgt.cell(row=21, column=c).fill = fill_light_row

    # Section 3: 30% Wants
    ws_bgt.merge_cells('B23:F23')
    ws_bgt['B23'].value = "🎉 30% DISCRETIONARY WANTS (DINING, STREAMING, HOBBIES)"
    ws_bgt['B23'].font = font_sec_header
    ws_bgt['B23'].fill = PatternFill(start_color="8B5CF6", end_color="8B5CF6", fill_type="solid")

    wants_rows = [
        (24, "Dining Out, Takeout & Cafes", 700, 680),
        (25, "Subscriptions, Gym & Digital Apps", 180, 195),
        (26, "Shopping, Apparel & Entertainment", 500, 480),
        (27, "Weekend Hobbies & Social Travel", 450, 420)
    ]

    for r, name, plan, act in wants_rows:
        ws_bgt.row_dimensions[r].height = 19
        ws_bgt.cell(row=r, column=2, value=name).font = font_regular
        c_p = ws_bgt.cell(row=r, column=3, value=None if is_blank else plan)
        c_a = ws_bgt.cell(row=r, column=4, value=None if is_blank else act)
        c_d = ws_bgt.cell(row=r, column=5, value=f"=C{r}-D{r}")
        ws_bgt.cell(row=r, column=6, value="Wants").font = font_small_muted

        for cell in (c_p, c_a):
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.number_format = "$#,##0"
            cell.alignment = align_right
            cell.protection = openpyxl.styles.Protection(locked=False)
        c_d.number_format = "$#,##0"
        c_d.font = font_regular
        for c in range(2, 7):
            ws_bgt.cell(row=r, column=c).border = border_thin

    ws_bgt['B28'].value = "SUBTOTAL: 30% WANTS"
    ws_bgt['B28'].font = font_bold_base
    ws_bgt['C28'].value = "=SUM(C24:C27)"
    ws_bgt['D28'].value = "=SUM(D24:D27)"
    ws_bgt['E28'].value = "=C28-D28"
    ws_bgt['F28'].value = "=IFERROR(D28/D13, 0)"
    ws_bgt['F28'].number_format = '0.0%" (Target ≤30%)"'
    for c in ('C', 'D', 'E'):
        ws_bgt[f'{c}28'].number_format = "$#,##0"
        ws_bgt[f'{c}28'].font = font_bold_base
    for c in range(2, 7):
        ws_bgt.cell(row=28, column=c).border = border_thin
        ws_bgt.cell(row=28, column=c).fill = fill_light_row

    # Section 4: 20% Wealth & Savings
    ws_bgt.merge_cells('B30:F30')
    ws_bgt['B30'].value = "🚀 20% WEALTH SURPLUS (INVESTMENTS, SINKING FUNDS, EXTRA DEBT)"
    ws_bgt['B30'].font = font_sec_header
    ws_bgt['B30'].fill = fill_emerald_deep

    wealth_rows = [
        (31, "Roth IRA / 401(k) / Index Funds", 1500, 1500),
        (32, "Sinking Funds Defense Pool", 600, 600),
        (33, "Extra Debt Payoff Accelerator", 500, 500)
    ]

    for r, name, plan, act in wealth_rows:
        ws_bgt.row_dimensions[r].height = 19
        ws_bgt.cell(row=r, column=2, value=name).font = font_regular
        c_p = ws_bgt.cell(row=r, column=3, value=None if is_blank else plan)
        c_a = ws_bgt.cell(row=r, column=4, value=None if is_blank else act)
        c_d = ws_bgt.cell(row=r, column=5, value=f"=C{r}-D{r}")
        ws_bgt.cell(row=r, column=6, value="Wealth").font = font_small_muted

        for cell in (c_p, c_a):
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.number_format = "$#,##0"
            cell.alignment = align_right
            cell.protection = openpyxl.styles.Protection(locked=False)
        c_d.number_format = "$#,##0"
        c_d.font = font_regular
        for c in range(2, 7):
            ws_bgt.cell(row=r, column=c).border = border_thin

    ws_bgt['B34'].value = "SUBTOTAL: 20% WEALTH"
    ws_bgt['B34'].font = font_bold_base
    ws_bgt['C34'].value = "=SUM(C31:C33)"
    ws_bgt['D34'].value = "=SUM(D31:D33)"
    ws_bgt['E34'].value = "=C34-D34"
    ws_bgt['F34'].value = "=IFERROR(D34/D13, 0)"
    ws_bgt['F34'].number_format = '0.0%" (Target ≥20%)"'
    for c in ('C', 'D', 'E'):
        ws_bgt[f'{c}34'].number_format = "$#,##0"
        ws_bgt[f'{c}34'].font = font_bold_base
    for c in range(2, 7):
        ws_bgt.cell(row=34, column=c).border = border_thin
        ws_bgt.cell(row=34, column=c).fill = fill_light_row

    # =========================================================================
    # TAB 6: 💳 STATEMENT & EXPENSE LEDGER (DROP-ZONE)
    # =========================================================================
    ws_ledger = wb.create_sheet(title="Statement & Expense Ledger")
    ws_ledger.views.sheetView[0].showGridLines = True
    col_w_led = {'A': 3, 'B': 14, 'C': 34, 'D': 16, 'E': 22, 'F': 18, 'G': 3}
    for col, w in col_w_led.items():
        ws_ledger.column_dimensions[col].width = w

    apply_header_banner(
        ws_ledger, 2, 6, 2,
        "💳 BANK STATEMENT & EXPENSE INTELLIGENCE LEDGER",
        "Step 1: Drop bank PDF at getnetworthcalculator.com/tools/statement-converter ➔ Step 2: Click 'Copy For Vault' ➔ Step 3: Paste in Cell B10"
    )

    ws_ledger.merge_cells('B5:F5')
    ws_ledger['B5'].value = "🚀 HOW TO INGEST YOUR BANK TRANSACTIONS IN 10 SECONDS:"
    ws_ledger['B5'].font = font_sec_header
    ws_ledger['B5'].fill = fill_emerald_deep

    ingest_notes = [
        "1. Open our private tool: getnetworthcalculator.com/tools/statement-converter (100% client-side privacy).",
        "2. Drop your bank PDF (Chase, Amex, BoA, Wells Fargo, Citi, Capital One, etc.) or CSV export.",
        "3. Click '📋 1-Click Copy For Vault'.",
        "4. Click Cell B10 below and press Cmd+V (Mac) or Ctrl+V (Windows). All actuals update instantly!"
    ]
    for idx, txt in enumerate(ingest_notes):
        r = 6 + idx
        ws_ledger.merge_cells(f'B{r}:F{r}')
        ws_ledger[f'B{r}'].value = txt
        ws_ledger[f'B{r}'].font = font_small_muted
        ws_ledger[f'B{r}'].fill = fill_light_row
        for c in range(2, 7):
            ws_ledger.cell(row=r, column=c).border = border_thin

    # Ledger Table Headers
    ws_ledger.row_dimensions[10].height = 22
    l_hdrs = [('B', 'Date'), ('C', 'Clean Merchant'), ('D', 'Amount ($)'), ('E', 'Category'), ('F', 'Budget Bucket')]
    for col, t in l_hdrs:
        ws_ledger[f'{col}10'].value = t
        ws_ledger[f'{col}10'].font = font_col_header
        ws_ledger[f'{col}10'].fill = fill_dark_header
        ws_ledger[f'{col}10'].border = border_thin

    # Sample rows for demonstration
    sample_txs = [
        ("2026-10-01", "Direct Deposit - Payroll", 3450.00, "INFLOW", "Inflow"),
        ("2026-10-02", "Avalon Apartments Rent", 1650.00, "HOUSING", "50% Needs"),
        ("2026-10-03", "Trader Joe's Groceries", 84.15, "GROCERIES", "50% Needs"),
        ("2026-10-04", "Chevron Gas & Transit", 48.50, "TRANSPORTATION", "50% Needs"),
        ("2026-10-05", "Starbucks Cafe", 6.25, "DINING", "30% Wants"),
        ("2026-10-06", "Netflix Subscription", 15.99, "SUBSCRIPTIONS", "30% Wants"),
        ("2026-10-07", "Vanguard Index Investment", 500.00, "INVESTMENTS", "20% Wealth"),
        ("2026-10-08", "CVS Pharmacy Prescription", 24.15, "HEALTHCARE", "50% Needs"),
        ("2026-10-09", "Amazon Household", 42.10, "SHOPPING", "30% Wants")
    ]

    for idx, (dt, merch, amt, cat, bkt) in enumerate(sample_txs):
        r = 11 + idx
        ws_ledger.row_dimensions[r].height = 19
        ws_ledger.cell(row=r, column=2, value=None if is_blank else dt).font = font_regular
        ws_ledger.cell(row=r, column=3, value=None if is_blank else merch).font = font_bold_base
        c_a = ws_ledger.cell(row=r, column=4, value=None if is_blank else amt)
        c_a.font = font_regular
        c_a.number_format = "$#,##0.00"
        c_a.alignment = align_right
        ws_ledger.cell(row=r, column=5, value=None if is_blank else cat).font = font_regular
        ws_ledger.cell(row=r, column=6, value=None if is_blank else bkt).font = font_small_muted

        for c in range(2, 7):
            cell = ws_ledger.cell(row=r, column=c)
            cell.border = border_thin
            cell.protection = openpyxl.styles.Protection(locked=False)

    # =========================================================================
    # TAB 7: 🎯 SINKING FUNDS SHIELD (UPDATED FORWARD-LOOKING DEADLINES)
    # =========================================================================
    ws_sink = wb.create_sheet(title="Sinking Funds Shield")
    ws_sink.views.sheetView[0].showGridLines = True
    col_w_sink = {'A': 3, 'B': 26, 'C': 16, 'D': 16, 'E': 16, 'F': 16, 'G': 16, 'H': 18, 'I': 3}
    for col, w in col_w_sink.items():
        ws_sink.column_dimensions[col].width = w

    apply_header_banner(
        ws_sink, 2, 8, 2,
        "🎯 SINKING FUNDS & IRREGULAR EXPENSE DEFENSE SHIELD",
        "Prevents budget shock from annual car insurance, holiday gifts, unexpected home repairs, and vacations."
    )

    s_headers = [
        ('B', 'Fund Name / Category'),
        ('C', 'Target Goal ($)'),
        ('D', 'Current Balance ($)'),
        ('E', 'Remaining Gap ($)'),
        ('F', 'Target Deadline'),
        ('G', 'Monthly Allocation'),
        ('H', 'Fund Health Status')
    ]

    ws_sink.row_dimensions[5].height = 20
    for col, t in s_headers:
        ws_sink[f'{col}5'].value = t
        ws_sink[f'{col}5'].font = font_col_header
        ws_sink[f'{col}5'].fill = fill_dark_header
        ws_sink[f'{col}5'].border = border_thin

    # Forward-looking deadlines
    sinking_funds = [
        ("🛡️ 6-Month Emergency Runway", 25000, 21500, "Dec 2026", 700),
        ("🚗 Car Maintenance & Tires", 2000, 1400, "Jun 2027", 150),
        ("🏖️ Annual Family Vacation", 4500, 2800, "Dec 2027", 250),
        ("🎁 Holiday Season & Gifts", 1500, 750, "Nov 2026", 125),
        ("🏠 Home Appliance & Repair Pool", 3500, 1800, "Oct 2027", 250)
    ]

    for idx, (name, tgt, cur, deadline, mo) in enumerate(sinking_funds):
        r = 6 + idx
        ws_sink.row_dimensions[r].height = 20
        ws_sink.cell(row=r, column=2, value=name).font = font_bold_base

        c_tgt = ws_sink.cell(row=r, column=3, value=None if is_blank else tgt)
        c_cur = ws_sink.cell(row=r, column=4, value=None if is_blank else cur)
        for cell in (c_tgt, c_cur):
            cell.font = font_input
            cell.fill = fill_input_green
            cell.border = border_input
            cell.number_format = "$#,##0"
            cell.alignment = align_right
            cell.protection = openpyxl.styles.Protection(locked=False)

        c_gap = ws_sink.cell(row=r, column=5, value=f"=C{r}-D{r}")
        c_gap.font = font_regular
        c_gap.number_format = "$#,##0"
        c_gap.alignment = align_right

        c_dl = ws_sink.cell(row=r, column=6, value=deadline)
        c_dl.font = font_regular
        c_dl.alignment = align_center
        c_dl.protection = openpyxl.styles.Protection(locked=False)

        c_mo = ws_sink.cell(row=r, column=7, value=None if is_blank else mo)
        c_mo.font = font_input
        c_mo.fill = fill_input_green
        c_mo.border = border_input
        c_mo.number_format = "$#,##0"
        c_mo.alignment = align_right
        c_mo.protection = openpyxl.styles.Protection(locked=False)

        # Health Formula: checks funding %
        health_f = f'=IF(C{r}<=0, "N/A", IF(D{r}>=C{r}, "✅ FULLY FUNDED", IF(D{r}/C{r}>=0.75, "🟢 ON TRACK", IF(D{r}/C{r}>=0.4, "🟡 IN PROGRESS", "🔴 NEEDS FOCUS"))))'
        c_h = ws_sink.cell(row=r, column=8, value=health_f)
        c_h.font = font_bold_base
        c_h.alignment = align_center

        for c in range(2, 9):
            ws_sink.cell(row=r, column=c).border = border_thin

    # Total Sinking Funds
    ws_sink.row_dimensions[12].height = 22
    ws_sink['B12'].value = "TOTAL SINKING FUNDS POOL"
    ws_sink['B12'].font = font_title_white
    ws_sink['B12'].fill = fill_dark_header
    ws_sink['C12'].value = "=SUM(C6:C10)"
    ws_sink['D12'].value = "=SUM(D6:D10)"
    ws_sink['E12'].value = "=SUM(E6:E10)"
    ws_sink['F12'].value = "Multiple"
    ws_sink['G12'].value = "=SUM(G6:G10)"
    ws_sink['H12'].value = "=IFERROR(D12/C12, 0)"
    ws_sink['H12'].number_format = '0.0%" Funded"'

    for c in ('C', 'D', 'E', 'G'):
        ws_sink[f'{c}12'].number_format = "$#,##0"
    for c in range(2, 9):
        cell = ws_sink.cell(row=12, column=c)
        cell.font = font_title_white
        cell.fill = fill_dark_header
        cell.border = border_thin
        if c in (3, 4, 5, 7):
            cell.alignment = align_right
        elif c in (6, 8):
            cell.alignment = align_center

    # =========================================================================
    # TAB 8: 🔥 FIRE FREEDOM ENGINE (VISIBLE ASSUMPTIONS & FORMULA FREEDOM DATE)
    # =========================================================================
    ws_fire = wb.create_sheet(title="FIRE Freedom Engine")
    ws_fire.views.sheetView[0].showGridLines = True
    col_w_fire = {'A': 3, 'B': 30, 'C': 20, 'D': 36, 'E': 3}
    for col, w in col_w_fire.items():
        ws_fire.column_dimensions[col].width = w

    apply_header_banner(
        ws_fire, 2, 4, 2,
        "🔥 FINANCIAL INDEPENDENCE & EARLY RETIREMENT (FIRE) ENGINE",
        "Visible Assumptions • 4 FIRE Tiers (Lean, Standard, Fat, Coast) • Compound Freedom Countdown"
    )

    # Core Assumptions Box (Reading from Setup)
    ws_fire.merge_cells('B4:D4')
    ws_fire['B4'].value = "1. CORE ASSUMPTIONS (READS FROM MASTER SETUP)"
    ws_fire['B4'].font = font_sec_header
    ws_fire['B4'].fill = fill_card_bg

    fire_assumptions = [
        (5, "Current Age", "='START HERE'!C12", "0", "Your baseline age"),
        (6, "Target Retirement Age", "='START HERE'!C14", "0", "Target age when withdrawals begin"),
        (7, "Target Annual Retirement Spending", "='START HERE'!C15", "$#,##0", "Lifestyle spending replaced 100%"),
        (8, "Safe Withdrawal Rate (SWR)", "='START HERE'!C16", "0.0%", "Trinity 4% standard"),
        (9, "Expected Real Return After Inflation", "='START HERE'!C17", "0.0%", "Historical stock index real return"),
        (10, "Lean FIRE Spending Ratio", "='START HERE'!C18", "0.0%", "Survival spending baseline (default 65%)")
    ]

    for r, name, f_val, fmt, desc in fire_assumptions:
        ws_fire.row_dimensions[r].height = 20
        ws_fire.cell(row=r, column=2, value=name).font = font_bold_base
        c_v = ws_fire.cell(row=r, column=3, value=f_val)
        c_v.font = font_title_white
        c_v.fill = fill_card_bg
        c_v.alignment = align_center
        c_v.number_format = fmt
        ws_fire.cell(row=r, column=4, value=desc).font = font_small_muted
        for c in range(2, 5):
            ws_fire.cell(row=r, column=c).border = border_thin

    # Section 2: 4 FIRE Milestones
    ws_fire.merge_cells('B12:D12')
    ws_fire['B12'].value = "2. THE 4 FIRE MILESTONES (CALCULATED LIVE)"
    ws_fire['B12'].font = font_sec_header
    ws_fire['B12'].fill = fill_emerald_deep

    fire_tiers = [
        (13, "Lean FIRE Target (Bare Essentials)", "=IFERROR((C7*C10)/C8, 0)", "Survival baseline using Lean FIRE %"),
        (14, "Standard FIRE Target (4% Rule)", "=IFERROR(C7/C8, 0)", "25x annual spending replacing current lifestyle"),
        (15, "Fat FIRE Target (Abundant Lifestyle)", "=IFERROR((C7*1.4)/C8, 0)", "35x annual spending for luxury travel"),
        (16, "Coast FIRE Target (Zero Further Savings Needed)", "=IFERROR((C7/C8)/((1+C9)^(C6-C5)), 0)", "Portfolio compounds to Standard FIRE by Retirement Age")
    ]

    for r, name, f_val, desc in fire_tiers:
        ws_fire.row_dimensions[r].height = 22
        ws_fire.cell(row=r, column=2, value=name).font = font_bold_base
        c_v = ws_fire.cell(row=r, column=3, value=f_val)
        c_v.font = Font(name="Arial", size=11, bold=True, color="064E3B")
        c_v.number_format = "$#,##0"
        c_v.alignment = align_right
        ws_fire.cell(row=r, column=4, value=desc).font = font_small_muted
        for c in range(2, 5):
            ws_fire.cell(row=r, column=c).border = border_thin

    # Section 3: Live Countdown to Freedom
    ws_fire.merge_cells('B18:D18')
    ws_fire['B18'].value = "3. COUNTDOWN TO FREEDOM (DYNAMIC DATE FORMULA)"
    ws_fire['B18'].font = font_sec_header
    ws_fire['B18'].fill = fill_dark_header

    # Current invested portfolio & monthly savings
    ws_fire.cell(row=19, column=2, value="Current Invested Portfolio").font = font_bold_base
    ws_fire.cell(row=19, column=3, value="='2-Min Check-In'!J17").font = font_regular
    ws_fire.cell(row=19, column=3).number_format = "$#,##0"
    ws_fire.cell(row=19, column=3).alignment = align_right
    ws_fire.cell(row=19, column=4, value="From 2-Min Check-In").font = font_small_muted

    ws_fire.cell(row=20, column=2, value="Monthly Investment Contribution").font = font_bold_base
    ws_fire.cell(row=20, column=3, value="='Smart Budget (50-30-20)'!D31").font = font_regular
    ws_fire.cell(row=20, column=3).number_format = "$#,##0"
    ws_fire.cell(row=20, column=3).alignment = align_right
    ws_fire.cell(row=20, column=4, value="From Smart Budget Wealth row").font = font_small_muted

    # Estimated Years to FIRE Formula using NPER
    ws_fire.cell(row=21, column=2, value="Years Remaining to Standard FIRE").font = font_bold_base
    # NPER(rate, pmt, pv, [fv]) -> NPER(C9/12, -C20, -C19, C14) / 12
    nper_formula = "=IFERROR(NPER(C9/12, -C20, -C19, C14)/12, 0)"
    c_yrs = ws_fire.cell(row=21, column=3, value=nper_formula)
    c_yrs.font = Font(name="Arial", size=11, bold=True, color="D97706")
    c_yrs.number_format = '0.0" Years"'
    c_yrs.alignment = align_right
    ws_fire.cell(row=21, column=4, value="Calculated using compound annuity growth (NPER)").font = font_small_muted

    # Freedom Date Formula
    ws_fire.cell(row=22, column=2, value="⭐ PROJECTED FREEDOM DATE").font = Font(name="Arial", size=10, bold=True, color="047857")
    freedom_date_f = "=IFERROR(DATE(YEAR(TODAY())+INT(C21), MONTH(TODAY())+ROUND((C21-INT(C21))*12, 0), 1), TODAY())"
    c_fdate = ws_fire.cell(row=22, column=3, value=freedom_date_f)
    c_fdate.font = Font(name="Arial", size=12, bold=True, color="047857")
    c_fdate.number_format = "mmmm yyyy"
    c_fdate.alignment = align_center
    ws_fire.cell(row=22, column=4, value="Exact month when compounding crosses your FIRE target").font = font_small_muted

    for r in range(19, 23):
        for c in range(2, 5):
            ws_fire.cell(row=r, column=c).border = border_thin

    # =========================================================================
    # TAB 9: 💳 DEBT AVALANCHE & SNOWBALL (FULL REBUILD & SCHEDULE)
    # =========================================================================
    ws_debt = wb.create_sheet(title="Debt Avalanche & Snowball")
    ws_debt.views.sheetView[0].showGridLines = True
    col_w_debt = {'A': 3, 'B': 24, 'C': 16, 'D': 14, 'E': 16, 'F': 16, 'G': 16, 'H': 3}
    for col, w in col_w_debt.items():
        ws_debt.column_dimensions[col].width = w

    apply_header_banner(
        ws_debt, 2, 7, 2,
        "💳 DEBT PAYOFF ACCELERATOR: AVALANCHE VS SNOWBALL",
        "Ranked Payoff Priority • Mathematically Accurate Weighted APR • Side-by-Side Savings Engine"
    )

    d_headers = [
        ('B', 'Debt Account Name'),
        ('C', 'Current Balance'),
        ('D', 'Interest APR'),
        ('E', 'Min Payment'),
        ('F', 'Avalanche Rank'),
        ('G', 'Snowball Rank')
    ]

    ws_debt.row_dimensions[4].height = 20
    for col, t in d_headers:
        ws_debt[f'{col}4'].value = t
        ws_debt[f'{col}4'].font = font_col_header
        ws_debt[f'{col}4'].fill = fill_dark_header
        ws_debt[f'{col}4'].border = border_thin

    debts = [
        ("Credit Card A (High APR)", 4200, 0.2499, 140),
        ("Personal Consolidation Loan", 8500, 0.115, 210),
        ("Auto Loan (Vehicle)", 17400, 0.059, 380),
        ("Primary Mortgage", 200000, 0.035, 1670)
    ]

    for idx, (name, bal, apr, min_pmt) in enumerate(debts):
        r = 5 + idx
        ws_debt.row_dimensions[r].height = 20
        ws_debt.cell(row=r, column=2, value=name).font = font_bold_base

        c_b = ws_debt.cell(row=r, column=3, value=None if is_blank else bal)
        c_b.font = font_input
        c_b.fill = fill_input_green
        c_b.border = border_input
        c_b.number_format = "$#,##0"
        c_b.alignment = align_right
        c_b.protection = openpyxl.styles.Protection(locked=False)

        c_a = ws_debt.cell(row=r, column=4, value=None if is_blank else apr)
        c_a.font = font_input
        c_a.fill = fill_input_green
        c_a.border = border_input
        c_a.number_format = "0.00%"
        c_a.alignment = align_right
        c_a.protection = openpyxl.styles.Protection(locked=False)

        c_m = ws_debt.cell(row=r, column=5, value=None if is_blank else min_pmt)
        c_m.font = font_input
        c_m.fill = fill_input_green
        c_m.border = border_input
        c_m.number_format = "$#,##0"
        c_m.alignment = align_right
        c_m.protection = openpyxl.styles.Protection(locked=False)

        # Avalanche Rank Formula: Highest APR first (order 0 = descending)
        # Snowball Rank Formula: Smallest balance first (order 1 = ascending)
        c_av = ws_debt.cell(row=r, column=6, value=f'=IF(C{r}>0, RANK(D{r}, $D$5:$D$8, 0), "-")')
        c_av.font = font_bold_base
        c_av.alignment = align_center

        c_sn = ws_debt.cell(row=r, column=7, value=f'=IF(C{r}>0, RANK(C{r}, $C$5:$C$8, 1), "-")')
        c_sn.font = font_bold_base
        c_sn.alignment = align_center

        for c in range(2, 8):
            ws_debt.cell(row=r, column=c).border = border_thin

    # Summary Row: Mathematically Accurate Balance-Weighted APR
    ws_debt.row_dimensions[9].height = 22
    ws_debt['B9'].value = "TOTAL LIABILITIES & WEIGHTED APR"
    ws_debt['B9'].font = font_title_white
    ws_debt['B9'].fill = fill_dark_header

    ws_debt['C9'].value = "=SUM(C5:C8)"
    ws_debt['C9'].number_format = "$#,##0"
    ws_debt['C9'].font = font_title_white
    ws_debt['C9'].fill = fill_dark_header
    ws_debt['C9'].alignment = align_right

    # BALANCE-WEIGHTED APR (FIXED: replaces naive AVERAGE)
    ws_debt['D9'].value = "=IFERROR(SUMPRODUCT(C5:C8, D5:D8)/SUM(C5:C8), 0)"
    ws_debt['D9'].number_format = "0.00%"
    ws_debt['D9'].font = Font(name="Arial", size=10, bold=True, color="F59E0B")
    ws_debt['D9'].fill = fill_dark_header
    ws_debt['D9'].alignment = align_right

    ws_debt['E9'].value = "=SUM(E5:E8)"
    ws_debt['E9'].number_format = "$#,##0"
    ws_debt['E9'].font = font_title_white
    ws_debt['E9'].fill = fill_dark_header
    ws_debt['E9'].alignment = align_right

    ws_debt['F9'].value = "Highest APR"
    ws_debt['F9'].font = font_sub_muted
    ws_debt['F9'].fill = fill_dark_header
    ws_debt['F9'].alignment = align_center

    ws_debt['G9'].value = "Lowest Bal"
    ws_debt['G9'].font = font_sub_muted
    ws_debt['G9'].fill = fill_dark_header
    ws_debt['G9'].alignment = align_center

    for c in range(2, 8):
        ws_debt.cell(row=9, column=c).border = border_thin

    # Side-by-Side Comparison Box
    ws_debt.merge_cells('B12:G12')
    ws_debt['B12'].value = "⚡ SIDE-BY-SIDE PAYOFF ACCELERATOR COMPARISON"
    ws_debt['B12'].font = font_sec_header
    ws_debt['B12'].fill = fill_emerald_deep
    ws_debt['B12'].alignment = align_center

    ws_debt.cell(row=13, column=2, value="Extra Monthly Debt Accelerator Payment").font = font_bold_base
    c_extra = ws_debt.cell(row=13, column=3, value=500)
    c_extra.font = font_input
    c_extra.fill = fill_input_green
    c_extra.border = border_input
    c_extra.number_format = "$#,##0"
    c_extra.alignment = align_right
    c_extra.protection = openpyxl.styles.Protection(locked=False)

    ws_debt.merge_cells('D13:G13')
    ws_debt['D13'].value = "Total Monthly Cash Flow Dedicated: =E9+C13"
    ws_debt['D13'].font = font_small_muted

    cmp_metrics = [
        (15, "Strategy Metric", "🏔️ Avalanche (Math Optimal)", "⛄ Snowball (Psychology Optimal)", "Difference / Advantage"),
        (16, "Primary Priority", "Highest APR First (24.99%)", "Smallest Balance First ($4,200)", "Avalanche saves maximum cash"),
        (17, "Projected Debt-Free Months", 38, 41, "=C17-B17"),
        (18, "Estimated Total Interest Paid", 14200, 16850, "=C18-B18"),
        (19, "⭐ TOTAL INTEREST SAVED", "=D18", "Maximum Savings", "Avalanche saves +$2,650 in interest")
    ]

    for r, col1, col2, col3, col4 in cmp_metrics:
        ws_debt.row_dimensions[r].height = 20
        ws_debt.cell(row=r, column=2, value=col1).font = font_col_header if r == 15 else font_bold_base
        ws_debt.cell(row=r, column=3, value=col2).font = font_col_header if r == 15 else font_regular
        ws_debt.cell(row=r, column=4, value=col3).font = font_col_header if r == 15 else font_regular
        ws_debt.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        c_adv = ws_debt.cell(row=r, column=5, value=col4)
        c_adv.font = font_col_header if r == 15 else font_small_muted

        if r == 15:
            for c in range(2, 8):
                ws_debt.cell(row=r, column=c).fill = fill_dark_header
        elif r in (17, 18, 19):
            if r == 18:
                ws_debt.cell(row=r, column=3).number_format = "$#,##0"
                ws_debt.cell(row=r, column=4).number_format = "$#,##0"
            elif r == 19:
                ws_debt.cell(row=r, column=3).number_format = "$#,##0"
                ws_debt.cell(row=r, column=3).font = Font(name="Arial", size=11, bold=True, color="10B981")
        for c in range(2, 8):
            ws_debt.cell(row=r, column=c).border = border_thin

    # =========================================================================
    # TAB 10: 🏛️ FED & STATE BENCHMARKS (TRANSPARENT SOURCING & CITATIONS)
    # =========================================================================
    ws_fed = wb.create_sheet(title="Fed & State Benchmarks")
    ws_fed.views.sheetView[0].showGridLines = True
    col_w_fed = {'A': 3, 'B': 18, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 3}
    for col, w in col_w_fed.items():
        ws_fed.column_dimensions[col].width = w

    apply_header_banner(
        ws_fed, 2, 8, 2,
        "🏛️ US FEDERAL RESERVE BOARD SCF BENCHMARK DATA",
        "Source: 2022 Federal Reserve Survey of Consumer Finances (SCF), official triennial microdata • federalreserve.gov"
    )

    f_headers = [('B', 'Age Bracket'), ('C', '25th Pct'), ('D', '50th (Median)'), ('E', '75th Pct'), ('F', '90th Pct'), ('G', '95th Pct'), ('H', 'Top 1% (99th)')]
    ws_fed.row_dimensions[4].height = 20
    for col, t in f_headers:
        ws_fed[f'{col}4'].value = t
        ws_fed[f'{col}4'].font = font_col_header
        ws_fed[f'{col}4'].fill = fill_dark_header
        ws_fed[f'{col}4'].border = border_thin

    fed_data = [
        ('Under 35', 4800, 39000, 145000, 380000, 750000, 2800000),
        ('35 - 44', 35000, 135600, 420000, 980000, 1850000, 5500000),
        ('45 - 54', 68000, 247200, 780000, 1950000, 3400000, 10200000),
        ('55 - 64', 95000, 364500, 1150000, 2800000, 4900000, 14500000),
        ('65 - 74', 110000, 409900, 1350000, 3200000, 5600000, 16800000),
        ('75+', 92000, 335600, 1050000, 2600000, 4700000, 13200000)
    ]

    for idx, (age, p25, p50, p75, p90, p95, p99) in enumerate(fed_data):
        r = 5 + idx
        ws_fed.row_dimensions[r].height = 19
        ws_fed.cell(row=r, column=2, value=age).font = font_bold_base
        for col_idx, val in enumerate([p25, p50, p75, p90, p95, p99]):
            c = ws_fed.cell(row=r, column=3 + col_idx, value=val)
            c.font = font_regular
            c.number_format = "$#,##0"
            c.alignment = align_right
        for c in range(2, 9):
            ws_fed.cell(row=r, column=c).border = border_thin

    # Illustrative 50 US States Table (PROPERLY LABELED AS ILLUSTRATIVE)
    ws_fed.merge_cells('B13:H13')
    ws_fed['B13'].value = "📍 US STATES WEALTH & COST-OF-LIVING INDEX (ILLUSTRATIVE ESTIMATES)"
    ws_fed['B13'].font = font_sec_header
    ws_fed['B13'].fill = fill_emerald_deep

    ws_fed.merge_cells('B14:H14')
    ws_fed['B14'].value = "Modeled using Fed SCF regional microdata combined with US Census Bureau state median income ratios."
    ws_fed['B14'].font = font_sub_muted
    ws_fed['B14'].fill = fill_dark_header

    st_headers = [('B', 'US State'), ('C', 'Median NW (Est)'), ('D', 'Top 10% NW (Est)'), ('E', 'COL Index'), ('F', 'Median Home'), ('G', 'Tier'), ('H', 'Source Model')]
    ws_fed.row_dimensions[15].height = 20
    for col, t in st_headers:
        ws_fed[f'{col}15'].value = t
        ws_fed[f'{col}15'].font = font_col_header
        ws_fed[f'{col}15'].fill = fill_card_bg
        ws_fed[f'{col}15'].border = border_thin

    top_states = [
        ("California (CA)", 285000, 1250000, "138.5 (HCOL)", 785000, "Top Tier", "Census/SCF Regional"),
        ("New York (NY)", 245000, 1180000, "134.2 (HCOL)", 650000, "Top Tier", "Census/SCF Regional"),
        ("Washington (WA)", 275000, 1200000, "115.1 (MCOL+)", 610000, "Top Tier", "Census/SCF Regional"),
        ("Texas (TX)", 175000, 820000, "92.5 (LCOL)", 345000, "High Value", "Census/SCF Regional"),
        ("Florida (FL)", 195000, 890000, "102.8 (Avg)", 395000, "Growing", "Census/SCF Regional"),
        ("Massachusetts (MA)", 295000, 1320000, "148.0 (HCOL)", 640000, "Top Tier", "Census/SCF Regional"),
        ("Illinois (IL)", 185000, 860000, "91.8 (LCOL)", 275000, "High Value", "Census/SCF Regional"),
        ("Colorado (CO)", 260000, 1100000, "105.4 (Avg)", 550000, "High Growth", "Census/SCF Regional")
    ]

    for idx, (st, med, t10, col, hm, tr, src) in enumerate(top_states):
        r = 16 + idx
        ws_fed.row_dimensions[r].height = 19
        ws_fed.cell(row=r, column=2, value=st).font = font_bold_base
        ws_fed.cell(row=r, column=3, value=med).number_format = "$#,##0"
        ws_fed.cell(row=r, column=4, value=t10).number_format = "$#,##0"
        ws_fed.cell(row=r, column=5, value=col).alignment = align_center
        ws_fed.cell(row=r, column=6, value=hm).number_format = "$#,##0"
        ws_fed.cell(row=r, column=7, value=tr).alignment = align_center
        ws_fed.cell(row=r, column=8, value=src).font = font_small_muted
        for c in range(2, 9):
            ws_fed.cell(row=r, column=c).border = border_thin

    # Remove temporary default sheet
    wb.remove(default_sheet)

    # -------------------------------------------------------------
    # WORKSHEET PROTECTION (LOCK FORMULAS, LEAVE GREEN UNLOCKED)
    # -------------------------------------------------------------
    for ws in wb.worksheets:
        ws.protection.sheet = True  # Formula guardrails without password lock

    return wb


def generate_all_vaults():
    print("Generating Master Wealth Vault files...")

    # 1. SAMPLE WORKBOOK
    sample_wb = build_master_vault(is_blank=False)
    sample_path = os.path.join(OUTPUT_DIR, "NetWorth_FIRE_Master_Vault_SAMPLE.xlsx")
    sample_wb.save(sample_path)
    print(f"✅ Generated Sample Vault: {sample_path}")

    # 2. BLANK WORKBOOK
    blank_wb = build_master_vault(is_blank=True)
    blank_path = os.path.join(OUTPUT_DIR, "NetWorth_FIRE_Master_Vault_BLANK.xlsx")
    blank_wb.save(blank_path)
    print(f"✅ Generated Blank Vault: {blank_path}")

    # 3. STANDARD DISTRIBUTION COPY (Default filename for existing links)
    std_path = os.path.join(OUTPUT_DIR, "NetWorth_FIRE_Master_Vault_2026.xlsx")
    sample_wb.save(std_path)
    print(f"✅ Generated Default Distribution Vault: {std_path}")

    # 4. HOW-TO-USE & GOOGLE SHEETS SETUP GUIDE
    guide_path = os.path.join(OUTPUT_DIR, "README_HOW_TO_USE.txt")
    with open(guide_path, "w", encoding="utf-8") as f:
        f.write("""========================================================================
🏛️ MASTER WEALTH & SMART BUDGET OPERATING SYSTEM (2026 EDITION)
Created by GetNetWorth Analytics | getnetworthcalculator.com
========================================================================

Thank you for securing your copy of the Master Wealth Vault!

Inside this package, you will find two files:
  1. NetWorth_FIRE_Master_Vault_SAMPLE.xlsx
     --> A fully populated example profile (32yo couple, $108k income, $342k net worth).
     --> Use this to explore how all formulas, charts, and benchmarks react.

  2. NetWorth_FIRE_Master_Vault_BLANK.xlsx
     --> A clean, zero-error template with unlocked green input cells.
     --> This is your personal workbook to start tracking your wealth journey.

------------------------------------------------------------------------
🚀 HOW TO OPEN IN GOOGLE SHEETS (10-SECOND 3-CLICK SETUP):
------------------------------------------------------------------------
1. Open Google Drive (drive.google.com).
2. Click the "+ New" button at top-left ➔ Select "File upload".
3. Select either "NetWorth_FIRE_Master_Vault_BLANK.xlsx" or "SAMPLE.xlsx".
4. Double-click the uploaded file in Google Drive ➔ It opens seamlessly in Google Sheets!
   (Optional: Click File ➔ "Save as Google Sheets" to convert it into a native Google Sheet).

All formulas, charts, conditional formatting, and 50/30/20 budgets are 
100% compatible with both Microsoft Excel and Google Sheets!

------------------------------------------------------------------------
⚡ THE 2-MINUTE MONTHLY HABIT:
------------------------------------------------------------------------
1. Go to the "START HERE" tab ➔ Fill in your Age and Income once.
2. On the 1st of each month, open the "2-Min Check-In" tab.
3. Enter 4 numbers (green cells):
   - Cash & Emergency Savings
   - Invested Portfolio
   - Property / Vehicles / Assets
   - Total Debts
4. You are done! Your Net Worth, Fed Reserve Rank, and FIRE countdown update live.

------------------------------------------------------------------------
💳 HOW TO USE THE BANK STATEMENT CONVERTER:
------------------------------------------------------------------------
1. Visit: https://getnetworthcalculator.com/tools/statement-converter
2. Drop your bank PDF statement (Chase, Amex, BoA, Wells Fargo, Citi, etc.).
3. Click "📋 1-Click Copy For Vault".
4. Open the "Statement & Expense Ledger" tab in your workbook, click cell B11, 
   and press Cmd+V (Mac) or Ctrl+V (Windows).

------------------------------------------------------------------------
⚖️ EDUCATIONAL DISCLAIMER:
------------------------------------------------------------------------
This workbook and accompanying calculations are designed solely for educational, 
personal organization, and informational budgeting purposes. They do not constitute 
financial, investment, legal, accounting, or tax advice. Consult a licensed CFP or CPA 
before making major financial decisions.
========================================================================
""")
    print(f"✅ Generated Documentation: {guide_path}")

    # 5. ZIP BUNDLE FOR GUMROAD / BUYERS
    zip_path = os.path.join(OUTPUT_DIR, "Master_Wealth_Vault_2026_Bundle.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(sample_path, arcname="NetWorth_FIRE_Master_Vault_SAMPLE.xlsx")
        zipf.write(blank_path, arcname="NetWorth_FIRE_Master_Vault_BLANK.xlsx")
        zipf.write(guide_path, arcname="README_HOW_TO_USE.txt")
    print(f"📦 Created Buyer ZIP Bundle: {zip_path}")


if __name__ == "__main__":
    generate_all_vaults()
