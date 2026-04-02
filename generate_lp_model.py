#!/usr/bin/env python3
"""
Nuveen Housing LP Financial Model Generator

Generates a professional Excel financial model for a Limited Partnership
investing in LIHTC loans, municipal bonds, JVs, securitization residuals
(B-Pieces), and cash. Styled with Nuveen corporate branding.

Usage: python3 generate_lp_model.py
Output: LP_Financial_Model.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, NamedStyle, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# =============================================================================
# NUVEEN BRANDING CONSTANTS (from Nuveen_PPT-Template_Universal-Print.pptx)
# =============================================================================

DARK_TEAL = "00313C"
WHITE = "FFFFFF"
LIGHT_GRAY = "BABCBC"
ORANGE = "FF8D19"
LIME_GREEN = "C3D62E"
TEAL_GREEN = "00AD97"
MEDIUM_GRAY = "747476"

INPUT_BG = "FFF3E0"       # Light orange/peach for input cells
CALC_BG = "F0F1F1"        # Very light gray for calculated cells
SUBTOTAL_BG = "E8E8E8"    # Slightly darker gray for subtotal rows

FONT_NAME = "Georgia"
NUM_YEARS = 20

ASSET_CLASSES = [
    "LIHTC Loans",
    "Municipal Bonds",
    "JVs",
    "Securitization Residuals (B-Pieces)",
    "Cash",
]
NUM_ASSETS = len(ASSET_CLASSES)

# =============================================================================
# DEFAULT INPUTS
# =============================================================================

DEFAULT_AUM = 500_000_000
DEFAULT_MGMT_FEE_BPS = 50
DEFAULT_PERF_FEE_BPS = 2000    # 20% of excess return
DEFAULT_HURDLE_BPS = 400       # 4% hurdle
DEFAULT_OTHER_EXP_BPS = 15
DEFAULT_LEVERAGE_PCT = 0.25
DEFAULT_LEVERAGE_COST_BPS = 350

# Year 1 defaults (same for all years as starting point)
DEFAULT_ALLOCATIONS = [0.35, 0.25, 0.15, 0.15, 0.10]
DEFAULT_YIELDS = [0.065, 0.045, 0.080, 0.075, 0.020]

# =============================================================================
# STYLE HELPERS
# =============================================================================

def make_font(bold=False, color=DARK_TEAL, size=10, italic=False):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size, italic=italic)

def make_fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

def make_border(bottom=None, top=None, left=None, right=None):
    def side(style):
        return Side(style=style, color=MEDIUM_GRAY) if style else None
    return Border(
        bottom=side(bottom), top=side(top),
        left=side(left), right=side(right),
    )

HEADER_FONT = make_font(bold=True, color=WHITE, size=11)
HEADER_FILL = make_fill(DARK_TEAL)
SUBHEADER_FONT = make_font(bold=True, color=DARK_TEAL, size=10)
SUBHEADER_FILL = make_fill(LIGHT_GRAY)
BODY_FONT = make_font()
INPUT_FONT = make_font(color="1A237E")  # Dark blue for input text
INPUT_FILL = make_fill(INPUT_BG)
CALC_FILL = make_fill(CALC_BG)
SUBTOTAL_FILL = make_fill(SUBTOTAL_BG)
BOLD_FONT = make_font(bold=True)
TITLE_FONT = make_font(bold=True, color=WHITE, size=14)
THIN_BORDER = make_border(bottom="thin", top="thin", left="thin", right="thin")
BOTTOM_BORDER = make_border(bottom="medium")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

FMT_CURRENCY = '$#,##0'
FMT_CURRENCY_DEC = '$#,##0.00'
FMT_PCT = '0.00%'
FMT_BPS = '0" bps"'
FMT_NUMBER = '#,##0'
FMT_PCT_INPUT = '0.0%'

# =============================================================================
# CELL STYLING UTILITIES
# =============================================================================

def style_header_row(ws, row, start_col, end_col, merge=False):
    """Apply dark teal header styling to a row."""
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

def style_subheader_row(ws, row, start_col, end_col):
    """Apply subheader styling (light gray background, bold teal text)."""
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.border = THIN_BORDER

def style_input_cell(cell, fmt=None):
    """Style a cell as an editable input."""
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = THIN_BORDER
    cell.alignment = RIGHT
    if fmt:
        cell.number_format = fmt

def style_calc_cell(cell, fmt=None, bold=False):
    """Style a cell as a calculated formula cell."""
    cell.font = BOLD_FONT if bold else BODY_FONT
    cell.fill = CALC_FILL
    cell.border = THIN_BORDER
    cell.alignment = RIGHT
    if fmt:
        cell.number_format = fmt

def style_subtotal_cell(cell, fmt=None):
    """Style a subtotal/total row cell."""
    cell.font = BOLD_FONT
    cell.fill = SUBTOTAL_FILL
    cell.border = Border(
        bottom=Side(style="medium", color=DARK_TEAL),
        top=Side(style="thin", color=MEDIUM_GRAY),
        left=Side(style="thin", color=MEDIUM_GRAY),
        right=Side(style="thin", color=MEDIUM_GRAY),
    )
    cell.alignment = RIGHT
    if fmt:
        cell.number_format = fmt

def style_label_cell(cell, bold=False, indent=0):
    """Style a row label cell."""
    cell.font = BOLD_FONT if bold else BODY_FONT
    cell.alignment = Alignment(
        horizontal="left", vertical="center", indent=indent
    )
    cell.border = THIN_BORDER


# =============================================================================
# SHEET 1: SUMMARY
# =============================================================================

def build_summary_sheet(wb):
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_properties.tabColor = DARK_TEAL

    # Column widths
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 18

    end_col = 8  # Column H

    # ── Title Row ──
    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    title_cell = ws.cell(row=row, column=1, value="Nuveen Housing LP Financial Model")
    title_cell.font = TITLE_FONT
    title_cell.fill = HEADER_FILL
    title_cell.alignment = CENTER
    for c in range(1, end_col + 1):
        ws.cell(row=row, column=c).fill = HEADER_FILL

    # ── Fund Parameters Section ──
    row = 3
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    ws.cell(row=row, column=1, value="Fund Parameters").font = HEADER_FONT
    style_header_row(ws, row, 1, end_col)

    params = [
        (5, "Initial AUM", DEFAULT_AUM, FMT_CURRENCY),
        (6, "Management Fee", DEFAULT_MGMT_FEE_BPS, FMT_BPS),
        (7, "Performance Fee (% of excess return)", DEFAULT_PERF_FEE_BPS, FMT_BPS),
        (8, "Hurdle Rate", DEFAULT_HURDLE_BPS, FMT_BPS),
        (9, "Other Expenses", DEFAULT_OTHER_EXP_BPS, FMT_BPS),
        (10, "Leverage Percentage", DEFAULT_LEVERAGE_PCT, FMT_PCT_INPUT),
        (11, "Leverage Cost", DEFAULT_LEVERAGE_COST_BPS, FMT_BPS),
    ]
    for r, label, default, fmt in params:
        label_cell = ws.cell(row=r, column=1, value=label)
        style_label_cell(label_cell)
        # Merge label across A-B
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        inp = ws.cell(row=r, column=3, value=default)
        style_input_cell(inp, fmt)

    # ── Summary Statistics Section ──
    row = 13
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    ws.cell(row=row, column=1, value="Summary Statistics").font = HEADER_FONT
    style_header_row(ws, row, 1, end_col)

    # Helper: sum a row across all years in Annual Model (cols C-V)
    def _am_sum(am_row):
        parts = [f"'Annual Model'!{get_column_letter(2+y)}{am_row}"
                 for y in range(1, NUM_YEARS + 1)]
        return "=" + "+".join(parts)

    # Last year column letter in Annual Model
    last_yr_cl = get_column_letter(2 + NUM_YEARS)

    # ── Return Metrics sub-section ──
    row = 14
    ws.cell(row=row, column=1, value="Return Metrics").font = SUBHEADER_FONT
    style_subheader_row(ws, row, 1, end_col)
    ws.cell(row=row, column=1).alignment = LEFT

    # Precompute comma-joined cell references for AVERAGE/MAX/MIN formulas
    def _am_refs(am_row):
        return ",".join(
            "'Annual Model'!" + get_column_letter(2 + y) + str(am_row)
            for y in range(1, NUM_YEARS + 1)
        )

    net_yield_refs = _am_refs(47)
    gross_yield_refs = _am_refs(20)
    end_aum_refs = _am_refs(48)

    stats = [
        (15, "Net IRR (Avg Annual Net Yield)",
         f"=AVERAGE({net_yield_refs})",
         FMT_PCT),
        (16, "Gross IRR (Avg Annual Gross Yield)",
         f"=AVERAGE({gross_yield_refs})",
         FMT_PCT),
        (17, "MOIC (Multiple on Invested Capital)",
         f"='Annual Model'!{last_yr_cl}48/Summary!$C$5",
         '0.00x'),
        (18, "DPI (Distributions to Paid-In)",
         f"=({_am_sum(46)[1:]})/Summary!$C$5",
         '0.00x'),
        (19, "Total Value to Paid-In (TVPI)",
         f"=('Annual Model'!{last_yr_cl}48+{_am_sum(46)[1:]})/Summary!$C$5",
         '0.00x'),
    ]

    for r, label, formula, fmt in stats:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        label_cell = ws.cell(row=r, column=1, value=label)
        style_label_cell(label_cell, indent=1)
        val_cell = ws.cell(row=r, column=3, value=formula)
        style_calc_cell(val_cell, fmt, bold=True)

    # ── Capital Metrics sub-section ──
    row = 21
    ws.cell(row=row, column=1, value="Capital Metrics").font = SUBHEADER_FONT
    style_subheader_row(ws, row, 1, end_col)
    ws.cell(row=row, column=1).alignment = LEFT

    cap_stats = [
        (22, "Initial Committed Capital",
         "=Summary!$C$5", FMT_CURRENCY),
        (23, "Ending AUM (Year 20)",
         f"='Annual Model'!{last_yr_cl}48", FMT_CURRENCY),
        (24, "Total AUM Growth",
         f"='Annual Model'!{last_yr_cl}48/Summary!$C$5-1", FMT_PCT),
        (25, "Total Cumulative Net Income",
         _am_sum(46), FMT_CURRENCY),
        (26, "Total Cumulative Gross Income",
         _am_sum(36), FMT_CURRENCY),
        (27, "Peak AUM",
         f"=MAX({end_aum_refs})",
         FMT_CURRENCY),
    ]

    for r, label, formula, fmt in cap_stats:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        label_cell = ws.cell(row=r, column=1, value=label)
        style_label_cell(label_cell, indent=1)
        val_cell = ws.cell(row=r, column=3, value=formula)
        style_calc_cell(val_cell, fmt, bold=True)

    # ── Fee & Expense Metrics sub-section ──
    row = 29
    ws.cell(row=row, column=1, value="Fee & Expense Metrics").font = SUBHEADER_FONT
    style_subheader_row(ws, row, 1, end_col)
    ws.cell(row=row, column=1).alignment = LEFT

    fee_stats = [
        (30, "Total Management Fees",
         _am_sum(39), FMT_CURRENCY),
        (31, "Total Performance Fees",
         _am_sum(41), FMT_CURRENCY),
        (32, "Total Other Expenses",
         _am_sum(42), FMT_CURRENCY),
        (33, "Total All-In Fees & Expenses",
         _am_sum(43), FMT_CURRENCY),
        (34, "Fee Drag (Fees / Gross Income)",
         f"=IF(C26=0,0,C33/C26)", FMT_PCT),
        (35, "Net-to-Gross Ratio",
         f"=IF(C26=0,0,C25/C26)", FMT_PCT),
    ]

    for r, label, formula, fmt in fee_stats:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        label_cell = ws.cell(row=r, column=1, value=label)
        style_label_cell(label_cell, indent=1)
        val_cell = ws.cell(row=r, column=3, value=formula)
        style_calc_cell(val_cell, fmt, bold=True)

    # ── Yield & Risk Metrics sub-section ──
    row = 37
    ws.cell(row=row, column=1, value="Yield & Risk Metrics").font = SUBHEADER_FONT
    style_subheader_row(ws, row, 1, end_col)
    ws.cell(row=row, column=1).alignment = LEFT

    yield_stats = [
        (38, "Year 1 Net Yield",
         "='Annual Model'!C47", FMT_PCT),
        (39, f"Year {NUM_YEARS} Net Yield",
         f"='Annual Model'!{last_yr_cl}47", FMT_PCT),
        (40, "Best Year Net Yield",
         f"=MAX({net_yield_refs})",
         FMT_PCT),
        (41, "Worst Year Net Yield",
         f"=MIN({net_yield_refs})",
         FMT_PCT),
        (42, "Year 1 Gross Yield",
         "='Annual Model'!C20", FMT_PCT),
        (43, "Gross-to-Net Spread (Avg)",
         f"=C16-C15", FMT_PCT),
        (44, "Leverage Contribution (Year 1)",
         "=IF('Annual Model'!C28=0,0,'Annual Model'!C34/'Annual Model'!C28)",
         FMT_PCT),
    ]

    for r, label, formula, fmt in yield_stats:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        label_cell = ws.cell(row=r, column=1, value=label)
        style_label_cell(label_cell, indent=1)
        val_cell = ws.cell(row=r, column=3, value=formula)
        style_calc_cell(val_cell, fmt, bold=True)

    # ── Legend ──
    row = 46
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    legend = ws.cell(row=row, column=1,
                     value="Input cells are highlighted in orange. Modify inputs to update the model.")
    legend.font = make_font(italic=True, color=MEDIUM_GRAY, size=9)

    swatch = ws.cell(row=row, column=5)
    swatch.fill = INPUT_FILL
    swatch.value = "= Input"
    swatch.font = make_font(size=9, color="1A237E")
    swatch.border = THIN_BORDER

    return ws


# =============================================================================
# SHEET 2: ANNUAL MODEL
# =============================================================================

def build_annual_model_sheet(wb):
    ws = wb.create_sheet("Annual Model")
    ws.sheet_properties.tabColor = ORANGE

    # Column widths
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 8
    for yr in range(1, NUM_YEARS + 1):
        col_letter = get_column_letter(2 + yr)
        ws.column_dimensions[col_letter].width = 16

    first_col = 3  # Column C = Year 1
    last_col = 2 + NUM_YEARS  # Column V = Year 20
    last_col_letter = get_column_letter(last_col)

    # ── Title Row ──
    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    title = ws.cell(row=row, column=1, value="Annual Financial Model")
    title.font = TITLE_FONT
    title.fill = HEADER_FILL
    title.alignment = CENTER
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).fill = HEADER_FILL

    # ── INPUTS SECTION ──
    row = 3
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.cell(row=row, column=1, value="INPUTS").font = HEADER_FONT
    style_header_row(ws, row, 1, last_col)

    # Year labels (row 4)
    row = 4
    ws.cell(row=row, column=1, value="").font = SUBHEADER_FONT
    ws.cell(row=row, column=1).fill = SUBHEADER_FILL
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=2).fill = SUBHEADER_FILL
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cell = ws.cell(row=row, column=col, value=f"Year {yr}")
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # Sub-header: Asset Allocation
    row = 5
    ws.cell(row=row, column=1, value="Asset Allocation").font = SUBHEADER_FONT
    ws.cell(row=row, column=2, value="%").font = make_font(color=MEDIUM_GRAY, size=9)
    style_subheader_row(ws, row, 1, last_col)
    ws.cell(row=row, column=1).alignment = LEFT

    # Asset allocation input rows (rows 6-10)
    for i, name in enumerate(ASSET_CLASSES):
        r = 6 + i
        label = ws.cell(row=r, column=1, value=name)
        style_label_cell(label, indent=1)
        ws.cell(row=r, column=2).border = THIN_BORDER

        for yr in range(1, NUM_YEARS + 1):
            col = 2 + yr
            cell = ws.cell(row=r, column=col, value=DEFAULT_ALLOCATIONS[i])
            style_input_cell(cell, FMT_PCT)

    # Total Allocation (row 11)
    row = 11
    label = ws.cell(row=row, column=1, value="Total Allocation")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"=SUM({cl}6:{cl}10)"
        style_subtotal_cell(cell, FMT_PCT)

    # Conditional formatting: highlight red if total != 100%
    red_fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
    red_font = Font(name=FONT_NAME, bold=True, color="B71C1C")
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell_ref = f"{cl}11"
        ws.conditional_formatting.add(
            cell_ref,
            CellIsRule(operator="notEqual", formula=["1"], fill=red_fill, font=red_font)
        )

    # Sub-header: Expected Yields
    row = 12
    ws.cell(row=row, column=1, value="Expected Yields").font = SUBHEADER_FONT
    ws.cell(row=row, column=2, value="%").font = make_font(color=MEDIUM_GRAY, size=9)
    style_subheader_row(ws, row, 1, last_col)
    ws.cell(row=row, column=1).alignment = LEFT

    # Yield input rows (rows 13-17)
    for i, name in enumerate(ASSET_CLASSES):
        r = 13 + i
        label = ws.cell(row=r, column=1, value=name)
        style_label_cell(label, indent=1)
        ws.cell(row=r, column=2).border = THIN_BORDER

        for yr in range(1, NUM_YEARS + 1):
            col = 2 + yr
            cell = ws.cell(row=r, column=col, value=DEFAULT_YIELDS[i])
            style_input_cell(cell, FMT_PCT)

    # ── PORTFOLIO SECTION ──
    row = 19
    ws.merge_cells(start_row=row - 1, start_column=1, end_row=row - 1, end_column=last_col)
    ws.cell(row=row - 1, column=1, value="PORTFOLIO PERFORMANCE").font = HEADER_FONT
    style_header_row(ws, row - 1, 1, last_col)

    # Beginning AUM (row 19)
    label = ws.cell(row=row, column=1, value="Beginning AUM")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2, value="$").font = make_font(color=MEDIUM_GRAY, size=9)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        if yr == 1:
            cell.value = "=Summary!$C$5"
        else:
            prev_cl = get_column_letter(col - 1)
            cell.value = f"={prev_cl}48"
        style_calc_cell(cell, FMT_CURRENCY, bold=True)

    # Gross Portfolio Yield (row 20)
    row = 20
    label = ws.cell(row=row, column=1, value="Gross Portfolio Yield")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2, value="%").font = make_font(color=MEDIUM_GRAY, size=9)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"=SUMPRODUCT({cl}6:{cl}10,{cl}13:{cl}17)"
        style_calc_cell(cell, FMT_PCT, bold=True)

    # ── Income by Asset Class ──
    row = 22
    ws.cell(row=row, column=1, value="Income by Asset Class").font = SUBHEADER_FONT
    ws.cell(row=row, column=2, value="$").font = make_font(color=MEDIUM_GRAY, size=9)
    style_subheader_row(ws, row, 1, last_col)
    ws.cell(row=row, column=1).alignment = LEFT

    for i, name in enumerate(ASSET_CLASSES):
        r = 23 + i
        label = ws.cell(row=r, column=1, value=name)
        style_label_cell(label, indent=1)
        ws.cell(row=r, column=2).border = THIN_BORDER

        alloc_row = 6 + i
        yield_row = 13 + i
        for yr in range(1, NUM_YEARS + 1):
            col = 2 + yr
            cl = get_column_letter(col)
            cell = ws.cell(row=r, column=col)
            cell.value = f"={cl}19*{cl}{alloc_row}*{cl}{yield_row}"
            style_calc_cell(cell, FMT_CURRENCY)

    # Total Asset Income (row 28)
    row = 28
    label = ws.cell(row=row, column=1, value="Total Asset Income")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"=SUM({cl}23:{cl}27)"
        style_subtotal_cell(cell, FMT_CURRENCY)

    # ── Leverage Section ──
    row = 30
    ws.cell(row=row, column=1, value="Leverage").font = SUBHEADER_FONT
    ws.cell(row=row, column=2, value="$").font = make_font(color=MEDIUM_GRAY, size=9)
    style_subheader_row(ws, row, 1, last_col)
    ws.cell(row=row, column=1).alignment = LEFT

    # Leveraged AUM (row 31)
    row = 31
    label = ws.cell(row=row, column=1, value="Leveraged AUM")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}19*Summary!$C$10"
        style_calc_cell(cell, FMT_CURRENCY)

    # Leveraged Income (row 32)
    row = 32
    label = ws.cell(row=row, column=1, value="Leveraged Income")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}31*{cl}20"
        style_calc_cell(cell, FMT_CURRENCY)

    # Leverage Cost (row 33)
    row = 33
    label = ws.cell(row=row, column=1, value="Leverage Cost")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}31*Summary!$C$11/10000"
        style_calc_cell(cell, FMT_CURRENCY)

    # Net Leverage Income (row 34)
    row = 34
    label = ws.cell(row=row, column=1, value="Net Leverage Income")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}32-{cl}33"
        style_subtotal_cell(cell, FMT_CURRENCY)

    # ── Total Gross Income ──
    row = 36
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    label = ws.cell(row=row, column=1, value="TOTAL GROSS INCOME")
    label.font = make_font(bold=True, color=WHITE, size=10)
    label.fill = make_fill(TEAL_GREEN)
    label.alignment = LEFT
    label.border = THIN_BORDER
    ws.cell(row=row, column=2).fill = make_fill(TEAL_GREEN)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}28+{cl}34"
        cell.font = make_font(bold=True, color=WHITE)
        cell.fill = make_fill(TEAL_GREEN)
        cell.number_format = FMT_CURRENCY
        cell.alignment = RIGHT
        cell.border = THIN_BORDER

    # ── Fees & Expenses Section ──
    row = 38
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.cell(row=row, column=1, value="FEES & EXPENSES").font = HEADER_FONT
    style_header_row(ws, row, 1, last_col)

    # Management Fee (row 39)
    row = 39
    label = ws.cell(row=row, column=1, value="Management Fee")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}19*Summary!$C$6/10000"
        style_calc_cell(cell, FMT_CURRENCY)

    # Performance Fee Basis (row 40)
    row = 40
    label = ws.cell(row=row, column=1, value="Return Above Hurdle")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"=MAX(0,{cl}36-{cl}19*Summary!$C$8/10000)"
        style_calc_cell(cell, FMT_CURRENCY)

    # Performance Fee (row 41)
    row = 41
    label = ws.cell(row=row, column=1, value="Performance Fee")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}40*Summary!$C$7/10000"
        style_calc_cell(cell, FMT_CURRENCY)

    # Other Expenses (row 42)
    row = 42
    label = ws.cell(row=row, column=1, value="Other Expenses")
    style_label_cell(label, indent=1)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}19*Summary!$C$9/10000"
        style_calc_cell(cell, FMT_CURRENCY)

    # Total Fees & Expenses (row 43)
    row = 43
    label = ws.cell(row=row, column=1, value="Total Fees & Expenses")
    style_label_cell(label, bold=True)
    ws.cell(row=row, column=2).border = THIN_BORDER
    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}39+{cl}41+{cl}42"
        style_subtotal_cell(cell, FMT_CURRENCY)

    # ── Returns Section ──
    row = 45
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.cell(row=row, column=1, value="RETURNS").font = HEADER_FONT
    style_header_row(ws, row, 1, last_col)

    # Net Income (row 46)
    row = 46
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    label = ws.cell(row=row, column=1, value="NET INCOME")
    label.font = make_font(bold=True, color=WHITE, size=10)
    label.fill = make_fill(DARK_TEAL)
    label.alignment = LEFT
    label.border = THIN_BORDER
    ws.cell(row=row, column=2).fill = make_fill(DARK_TEAL)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}36-{cl}43"
        cell.font = make_font(bold=True, color=WHITE)
        cell.fill = make_fill(DARK_TEAL)
        cell.number_format = FMT_CURRENCY
        cell.alignment = RIGHT
        cell.border = THIN_BORDER

    # Net Yield / Return to LPs (row 47)
    row = 47
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    label = ws.cell(row=row, column=1, value="NET YIELD (Return to LPs)")
    label.font = make_font(bold=True, color=WHITE, size=10)
    label.fill = make_fill(DARK_TEAL)
    label.alignment = LEFT
    label.border = THIN_BORDER
    ws.cell(row=row, column=2).fill = make_fill(DARK_TEAL)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"=IF({cl}19=0,0,{cl}46/{cl}19)"
        cell.font = make_font(bold=True, color=WHITE)
        cell.fill = make_fill(DARK_TEAL)
        cell.number_format = FMT_PCT
        cell.alignment = RIGHT
        cell.border = THIN_BORDER

    # Ending AUM (row 48)
    row = 48
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    label = ws.cell(row=row, column=1, value="ENDING AUM")
    label.font = make_font(bold=True, color=WHITE, size=10)
    label.fill = make_fill(ORANGE)
    label.alignment = LEFT
    label.border = THIN_BORDER
    ws.cell(row=row, column=2).fill = make_fill(ORANGE)
    ws.cell(row=row, column=2).border = THIN_BORDER

    for yr in range(1, NUM_YEARS + 1):
        col = 2 + yr
        cl = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = f"={cl}19+{cl}46"
        cell.font = make_font(bold=True, color=WHITE)
        cell.fill = make_fill(ORANGE)
        cell.number_format = FMT_CURRENCY
        cell.alignment = RIGHT
        cell.border = THIN_BORDER

    # Freeze panes: freeze columns A-B and rows 1-4
    ws.freeze_panes = "C5"

    return ws


# =============================================================================
# MAIN
# =============================================================================

def main():
    wb = Workbook()

    build_summary_sheet(wb)
    build_annual_model_sheet(wb)

    output_file = "LP_Financial_Model.xlsx"
    wb.save(output_file)
    print(f"Financial model generated: {output_file}")

if __name__ == "__main__":
    main()
