#!/usr/bin/env python3
"""
Cash Management Controller - Excel Workbook Generator
Apple-Inspired Modern Design

Features:
- Payment Calendar with day boxes and status indicators
- Time Deposit Manager
- Comprehensive Cash Flow Positions (Daily/Weekly/Monthly)
- All data from internal sheet sources - No external connections
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment,
    NamedStyle, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import (
    ColorScaleRule, FormulaRule, DataBarRule, IconSetRule
)
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from datetime import datetime, timedelta
from calendar import monthrange
import calendar


# ============================================================================
# APPLE-INSPIRED COLOR PALETTE
# ============================================================================
class Colors:
    """Modern Apple-inspired color palette"""
    # Primary Colors
    WHITE = "FFFFFF"
    BLACK = "1D1D1F"
    DARK_GRAY = "86868B"
    LIGHT_GRAY = "F5F5F7"

    # Accent Colors (iOS Style)
    BLUE = "007AFF"
    GREEN = "34C759"
    RED = "FF3B30"
    ORANGE = "FF9500"
    YELLOW = "FFCC00"
    PURPLE = "AF52DE"
    TEAL = "5AC8FA"

    # Soft Backgrounds
    SOFT_BLUE = "E3F2FD"
    SOFT_GREEN = "E8F5E9"
    SOFT_RED = "FFEBEE"
    SOFT_ORANGE = "FFF3E0"
    SOFT_YELLOW = "FFFDE7"
    SOFT_PURPLE = "F3E5F5"

    # Status Colors
    STATUS_PAID = "34C759"
    STATUS_PENDING = "FF9500"
    STATUS_OVERDUE = "FF3B30"
    STATUS_SCHEDULED = "007AFF"
    STATUS_CANCELLED = "86868B"


# ============================================================================
# STYLE DEFINITIONS
# ============================================================================
class Styles:
    """Pre-defined styles for consistent Apple-like appearance"""

    # Borders
    thin_border = Border(
        left=Side(style='thin', color=Colors.LIGHT_GRAY),
        right=Side(style='thin', color=Colors.LIGHT_GRAY),
        top=Side(style='thin', color=Colors.LIGHT_GRAY),
        bottom=Side(style='thin', color=Colors.LIGHT_GRAY)
    )

    card_border = Border(
        left=Side(style='medium', color=Colors.LIGHT_GRAY),
        right=Side(style='medium', color=Colors.LIGHT_GRAY),
        top=Side(style='medium', color=Colors.LIGHT_GRAY),
        bottom=Side(style='medium', color=Colors.LIGHT_GRAY)
    )

    # Fonts
    title_font = Font(name='Calibri', size=28, bold=True, color=Colors.BLACK)
    header_font = Font(name='Calibri', size=14, bold=True, color=Colors.BLACK)
    subheader_font = Font(name='Calibri', size=12, bold=True, color=Colors.DARK_GRAY)
    body_font = Font(name='Calibri', size=11, color=Colors.BLACK)
    small_font = Font(name='Calibri', size=9, color=Colors.DARK_GRAY)
    link_font = Font(name='Calibri', size=11, color=Colors.BLUE, underline='single')

    # Fills
    white_fill = PatternFill(start_color=Colors.WHITE, end_color=Colors.WHITE, fill_type='solid')
    light_gray_fill = PatternFill(start_color=Colors.LIGHT_GRAY, end_color=Colors.LIGHT_GRAY, fill_type='solid')
    blue_fill = PatternFill(start_color=Colors.BLUE, end_color=Colors.BLUE, fill_type='solid')
    soft_blue_fill = PatternFill(start_color=Colors.SOFT_BLUE, end_color=Colors.SOFT_BLUE, fill_type='solid')

    # Alignments
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')


def create_workbook():
    """Create the main workbook with all sheets"""
    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Create sheets in order
    create_dashboard_sheet(wb)
    create_payment_calendar_sheet(wb)
    create_time_deposit_sheet(wb)
    create_cash_flow_positions_sheet(wb)
    create_payments_data_sheet(wb)
    create_deposits_data_sheet(wb)
    create_cash_transactions_sheet(wb)
    create_settings_sheet(wb)

    return wb


def style_header_row(ws, row, start_col, end_col, fill_color=Colors.LIGHT_GRAY):
    """Style a header row with consistent formatting"""
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Styles.header_font
        cell.fill = fill
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border


def create_card(ws, start_row, start_col, end_row, end_col, title="", fill_color=Colors.WHITE):
    """Create a card-like container with Apple styling"""
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill

            # Set border based on position
            left = Side(style='medium', color=Colors.LIGHT_GRAY) if col == start_col else Side(style='thin', color=Colors.LIGHT_GRAY)
            right = Side(style='medium', color=Colors.LIGHT_GRAY) if col == end_col else Side(style='thin', color=Colors.LIGHT_GRAY)
            top = Side(style='medium', color=Colors.LIGHT_GRAY) if row == start_row else Side(style='thin', color=Colors.LIGHT_GRAY)
            bottom = Side(style='medium', color=Colors.LIGHT_GRAY) if row == end_row else Side(style='thin', color=Colors.LIGHT_GRAY)

            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    if title:
        title_cell = ws.cell(row=start_row, column=start_col, value=title)
        title_cell.font = Styles.header_font
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)


# ============================================================================
# DASHBOARD SHEET
# ============================================================================
def create_dashboard_sheet(wb):
    """Create the main dashboard with KPIs and overview"""
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False

    # Set column widths
    column_widths = [3, 20, 18, 18, 18, 18, 18, 18, 3]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Title
    ws.merge_cells('B2:H2')
    title_cell = ws['B2']
    title_cell.value = "Cash Management Controller"
    title_cell.font = Styles.title_font
    title_cell.alignment = Alignment(horizontal='left', vertical='center')

    # Subtitle with date
    ws.merge_cells('B3:H3')
    ws['B3'].value = f"Dashboard Overview • {datetime.now().strftime('%B %d, %Y')}"
    ws['B3'].font = Styles.subheader_font

    # ========== KPI CARDS ROW ==========
    kpi_row = 5

    # KPI 1: Total Cash Position
    create_card(ws, kpi_row, 2, kpi_row + 3, 3)
    ws.cell(row=kpi_row, column=2, value="Total Cash Position").font = Styles.small_font
    ws.merge_cells(start_row=kpi_row, start_column=2, end_row=kpi_row, end_column=3)
    ws.cell(row=kpi_row + 1, column=2, value="=CashTransactions!J2").font = Font(name='Calibri', size=24, bold=True, color=Colors.BLUE)
    ws.cell(row=kpi_row + 1, column=2).number_format = '"$"#,##0'
    ws.merge_cells(start_row=kpi_row + 1, start_column=2, end_row=kpi_row + 2, end_column=3)
    ws.cell(row=kpi_row + 3, column=2, value="Current Balance").font = Styles.small_font

    # KPI 2: Pending Payments
    create_card(ws, kpi_row, 4, kpi_row + 3, 5)
    ws.cell(row=kpi_row, column=4, value="Pending Payments").font = Styles.small_font
    ws.merge_cells(start_row=kpi_row, start_column=4, end_row=kpi_row, end_column=5)
    ws.cell(row=kpi_row + 1, column=4, value='=SUMIF(PaymentsData!F:F,"Pending",PaymentsData!D:D)').font = Font(name='Calibri', size=24, bold=True, color=Colors.ORANGE)
    ws.cell(row=kpi_row + 1, column=4).number_format = '"$"#,##0'
    ws.merge_cells(start_row=kpi_row + 1, start_column=4, end_row=kpi_row + 2, end_column=5)
    ws.cell(row=kpi_row + 3, column=4, value='=COUNTIF(PaymentsData!F:F,"Pending")&" payments"').font = Styles.small_font

    # KPI 3: Time Deposits
    create_card(ws, kpi_row, 6, kpi_row + 3, 7)
    ws.cell(row=kpi_row, column=6, value="Active Time Deposits").font = Styles.small_font
    ws.merge_cells(start_row=kpi_row, start_column=6, end_row=kpi_row, end_column=7)
    ws.cell(row=kpi_row + 1, column=6, value='=SUMIF(DepositsData!G:G,"Active",DepositsData!D:D)').font = Font(name='Calibri', size=24, bold=True, color=Colors.GREEN)
    ws.cell(row=kpi_row + 1, column=6).number_format = '"$"#,##0'
    ws.merge_cells(start_row=kpi_row + 1, start_column=6, end_row=kpi_row + 2, end_column=7)
    ws.cell(row=kpi_row + 3, column=6, value='=COUNTIF(DepositsData!G:G,"Active")&" deposits"').font = Styles.small_font

    # KPI 4: Overdue Payments
    create_card(ws, kpi_row, 8, kpi_row + 3, 8)
    ws.cell(row=kpi_row, column=8, value="Overdue").font = Styles.small_font
    ws.cell(row=kpi_row + 1, column=8, value='=COUNTIF(PaymentsData!F:F,"Overdue")').font = Font(name='Calibri', size=24, bold=True, color=Colors.RED)
    ws.merge_cells(start_row=kpi_row + 1, start_column=8, end_row=kpi_row + 2, end_column=8)
    ws.cell(row=kpi_row + 3, column=8, value="Items").font = Styles.small_font

    # ========== UPCOMING PAYMENTS SECTION ==========
    upcoming_row = 10
    ws.cell(row=upcoming_row, column=2, value="Upcoming Payments (Next 7 Days)").font = Styles.header_font

    # Headers
    headers = ["Date", "Payee", "Amount", "Category", "Status"]
    for i, header in enumerate(headers, 2):
        cell = ws.cell(row=upcoming_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.BLUE, end_color=Colors.BLUE, fill_type='solid')
        cell.alignment = Styles.center_align

    # Dynamic formulas for upcoming payments (first 5)
    for i in range(5):
        row = upcoming_row + 2 + i
        ws.cell(row=row, column=2, value=f"=IFERROR(INDEX(PaymentsData!B:B,SMALL(IF((PaymentsData!B:B>=TODAY())*(PaymentsData!B:B<=TODAY()+7),ROW(PaymentsData!B:B)),{i+1})),\"\")").number_format = 'MMM DD'
        ws.cell(row=row, column=3, value=f"=IFERROR(INDEX(PaymentsData!C:C,SMALL(IF((PaymentsData!B:B>=TODAY())*(PaymentsData!B:B<=TODAY()+7),ROW(PaymentsData!B:B)),{i+1})),\"\")")
        ws.cell(row=row, column=4, value=f"=IFERROR(INDEX(PaymentsData!D:D,SMALL(IF((PaymentsData!B:B>=TODAY())*(PaymentsData!B:B<=TODAY()+7),ROW(PaymentsData!B:B)),{i+1})),\"\")").number_format = '"$"#,##0.00'
        ws.cell(row=row, column=5, value=f"=IFERROR(INDEX(PaymentsData!E:E,SMALL(IF((PaymentsData!B:B>=TODAY())*(PaymentsData!B:B<=TODAY()+7),ROW(PaymentsData!B:B)),{i+1})),\"\")")
        ws.cell(row=row, column=6, value=f"=IFERROR(INDEX(PaymentsData!F:F,SMALL(IF((PaymentsData!B:B>=TODAY())*(PaymentsData!B:B<=TODAY()+7),ROW(PaymentsData!B:B)),{i+1})),\"\")")

        for col in range(2, 7):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font

    # ========== MATURING DEPOSITS SECTION ==========
    deposit_row = 10
    ws.cell(row=deposit_row, column=7, value="Maturing Deposits (Next 30 Days)").font = Styles.header_font
    ws.merge_cells(start_row=deposit_row, start_column=7, end_row=deposit_row, end_column=8)

    # Headers
    deposit_headers = ["Maturity", "Amount"]
    for i, header in enumerate(deposit_headers):
        cell = ws.cell(row=deposit_row + 1, column=7 + i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.GREEN, end_color=Colors.GREEN, fill_type='solid')
        cell.alignment = Styles.center_align

    # Deposit entries
    for i in range(5):
        row = deposit_row + 2 + i
        ws.cell(row=row, column=7, value=f"=IFERROR(INDEX(DepositsData!F:F,SMALL(IF((DepositsData!F:F>=TODAY())*(DepositsData!F:F<=TODAY()+30)*(DepositsData!G:G=\"Active\"),ROW(DepositsData!F:F)),{i+1})),\"\")").number_format = 'MMM DD'
        ws.cell(row=row, column=8, value=f"=IFERROR(INDEX(DepositsData!D:D,SMALL(IF((DepositsData!F:F>=TODAY())*(DepositsData!F:F<=TODAY()+30)*(DepositsData!G:G=\"Active\"),ROW(DepositsData!F:F)),{i+1})),\"\")").number_format = '"$"#,##0'

        for col in range(7, 9):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font

    # ========== CASH FLOW SUMMARY SECTION ==========
    cf_row = 19
    ws.cell(row=cf_row, column=2, value="Cash Flow Summary").font = Styles.header_font

    # Create summary table
    summary_headers = ["Period", "Inflows", "Outflows", "Net Flow", "End Balance"]
    for i, header in enumerate(summary_headers, 2):
        cell = ws.cell(row=cf_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.DARK_GRAY, end_color=Colors.DARK_GRAY, fill_type='solid')
        cell.alignment = Styles.center_align

    periods = ["Today", "This Week", "This Month", "Next 30 Days"]
    for i, period in enumerate(periods):
        row = cf_row + 2 + i
        ws.cell(row=row, column=2, value=period).font = Styles.body_font
        ws.cell(row=row, column=3, value=f"=CashFlowPositions!D{4+i}").number_format = '"$"#,##0'
        ws.cell(row=row, column=4, value=f"=CashFlowPositions!E{4+i}").number_format = '"$"#,##0'
        ws.cell(row=row, column=5, value=f"=CashFlowPositions!F{4+i}").number_format = '"$"#,##0'
        ws.cell(row=row, column=6, value=f"=CashFlowPositions!G{4+i}").number_format = '"$"#,##0'

        for col in range(2, 7):
            ws.cell(row=row, column=col).border = Styles.thin_border

    # Add conditional formatting for net flow
    ws.conditional_formatting.add('E21:E24',
        FormulaRule(formula=['E21>=0'], fill=PatternFill(start_color=Colors.SOFT_GREEN, end_color=Colors.SOFT_GREEN, fill_type='solid')))
    ws.conditional_formatting.add('E21:E24',
        FormulaRule(formula=['E21<0'], fill=PatternFill(start_color=Colors.SOFT_RED, end_color=Colors.SOFT_RED, fill_type='solid')))

    # ========== NAVIGATION LINKS ==========
    nav_row = 27
    ws.cell(row=nav_row, column=2, value="Quick Navigation").font = Styles.header_font

    nav_items = [
        ("Payment Calendar", "PaymentCalendar!A1"),
        ("Time Deposits", "TimeDeposits!A1"),
        ("Cash Flow Positions", "CashFlowPositions!A1"),
        ("Settings", "Settings!A1")
    ]

    for i, (name, link) in enumerate(nav_items):
        cell = ws.cell(row=nav_row + 1, column=2 + i, value=name)
        cell.font = Styles.link_font
        cell.hyperlink = link
        cell.alignment = Styles.center_align


# ============================================================================
# PAYMENT CALENDAR SHEET
# ============================================================================
def create_payment_calendar_sheet(wb):
    """Create the payment calendar with day boxes and status indicators"""
    ws = wb.create_sheet("PaymentCalendar")
    ws.sheet_view.showGridLines = False

    # Set column widths for calendar grid (7 columns for days + navigation)
    ws.column_dimensions['A'].width = 3
    for col in range(2, 9):  # B through H
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions['I'].width = 3
    ws.column_dimensions['J'].width = 25  # Legend column

    # Set row heights
    for row in range(1, 50):
        ws.row_dimensions[row].height = 18

    # Title
    ws.merge_cells('B2:H2')
    ws['B2'].value = "Payment Calendar"
    ws['B2'].font = Styles.title_font

    # Month/Year selector (current month)
    today = datetime.now()
    ws['B4'].value = "Selected Month:"
    ws['B4'].font = Styles.subheader_font
    ws['C4'].value = today.strftime('%B %Y')
    ws['C4'].font = Font(name='Calibri', size=14, bold=True, color=Colors.BLUE)

    # Create data validation for month selection
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]

    # Day headers
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for i, day in enumerate(days):
        cell = ws.cell(row=6, column=2 + i, value=day)
        cell.font = Font(name='Calibri', size=11, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.BLUE, end_color=Colors.BLUE, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Create calendar grid (6 weeks x 7 days)
    cal = calendar.Calendar(firstweekday=6)  # Sunday start
    year = today.year
    month = today.month

    # Get the days for the current month
    month_days = list(cal.monthdayscalendar(year, month))

    start_row = 7
    box_height = 5  # rows per day box

    for week_num, week in enumerate(month_days):
        week_start_row = start_row + (week_num * box_height)

        for day_num, day in enumerate(week):
            col = 2 + day_num

            # Create day box
            day_start_row = week_start_row
            day_end_row = week_start_row + box_height - 1

            # Merge cells for day number
            if day != 0:
                # Day number cell
                day_cell = ws.cell(row=day_start_row, column=col, value=day)
                day_cell.font = Font(name='Calibri', size=12, bold=True, color=Colors.BLACK)
                day_cell.alignment = Alignment(horizontal='right', vertical='top')

                # Check if this is today
                if day == today.day and month == today.month:
                    day_cell.fill = PatternFill(start_color=Colors.SOFT_BLUE, end_color=Colors.SOFT_BLUE, fill_type='solid')

                # Payment indicator cells (3 slots per day)
                for slot in range(1, 4):
                    slot_row = day_start_row + slot
                    slot_cell = ws.cell(row=slot_row, column=col)
                    # Formula to show payment if exists for this day
                    slot_cell.value = f'=IFERROR(INDEX(PaymentsData!C:C,SMALL(IF(DAY(PaymentsData!B:B)={day},IF(MONTH(PaymentsData!B:B)=MONTH(C4),ROW(PaymentsData!B:B))),{slot}))&" - $"&TEXT(INDEX(PaymentsData!D:D,SMALL(IF(DAY(PaymentsData!B:B)={day},IF(MONTH(PaymentsData!B:B)=MONTH(C4),ROW(PaymentsData!B:B))),{slot})),"#,##0"),"")'
                    slot_cell.font = Styles.small_font
                    slot_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            # Apply borders to entire day box
            for row in range(day_start_row, day_end_row + 1):
                cell = ws.cell(row=row, column=col)

                left = Side(style='thin', color=Colors.LIGHT_GRAY)
                right = Side(style='thin', color=Colors.LIGHT_GRAY)
                top = Side(style='thin', color=Colors.LIGHT_GRAY) if row == day_start_row else None
                bottom = Side(style='thin', color=Colors.LIGHT_GRAY) if row == day_end_row else None

                cell.border = Border(left=left, right=right, top=top, bottom=bottom)

                if day == 0:
                    cell.fill = PatternFill(start_color=Colors.LIGHT_GRAY, end_color=Colors.LIGHT_GRAY, fill_type='solid')

    # ========== STATUS LEGEND ==========
    legend_row = 7
    ws.cell(row=legend_row, column=10, value="Status Legend").font = Styles.header_font

    statuses = [
        ("Paid", Colors.STATUS_PAID),
        ("Pending", Colors.STATUS_PENDING),
        ("Scheduled", Colors.STATUS_SCHEDULED),
        ("Overdue", Colors.STATUS_OVERDUE),
        ("Cancelled", Colors.STATUS_CANCELLED)
    ]

    for i, (status, color) in enumerate(statuses):
        row = legend_row + 1 + i
        # Color indicator
        indicator = ws.cell(row=row, column=10, value="●")
        indicator.font = Font(name='Calibri', size=14, color=color)
        indicator.alignment = Alignment(horizontal='center', vertical='center')
        # Status name
        ws.cell(row=row, column=11, value=status).font = Styles.body_font

    # ========== DAILY TOTALS SECTION ==========
    totals_row = legend_row + 8
    ws.cell(row=totals_row, column=10, value="Monthly Totals").font = Styles.header_font
    ws.merge_cells(start_row=totals_row, start_column=10, end_row=totals_row, end_column=11)

    ws.cell(row=totals_row + 1, column=10, value="Scheduled:").font = Styles.body_font
    ws.cell(row=totals_row + 1, column=11, value='=SUMIF(PaymentsData!F:F,"Scheduled",PaymentsData!D:D)').number_format = '"$"#,##0'

    ws.cell(row=totals_row + 2, column=10, value="Pending:").font = Styles.body_font
    ws.cell(row=totals_row + 2, column=11, value='=SUMIF(PaymentsData!F:F,"Pending",PaymentsData!D:D)').number_format = '"$"#,##0'

    ws.cell(row=totals_row + 3, column=10, value="Paid:").font = Styles.body_font
    ws.cell(row=totals_row + 3, column=11, value='=SUMIF(PaymentsData!F:F,"Paid",PaymentsData!D:D)').number_format = '"$"#,##0'

    ws.cell(row=totals_row + 4, column=10, value="Overdue:").font = Styles.body_font
    ws.cell(row=totals_row + 4, column=11, value='=SUMIF(PaymentsData!F:F,"Overdue",PaymentsData!D:D)').number_format = '"$"#,##0'
    ws.cell(row=totals_row + 4, column=11).font = Font(name='Calibri', size=11, color=Colors.RED)


# ============================================================================
# TIME DEPOSIT MANAGER SHEET
# ============================================================================
def create_time_deposit_sheet(wb):
    """Create the Time Deposit Manager sheet"""
    ws = wb.create_sheet("TimeDeposits")
    ws.sheet_view.showGridLines = False

    # Set column widths
    column_widths = [3, 15, 18, 15, 18, 12, 15, 15, 18, 15, 3]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Title
    ws.merge_cells('B2:J2')
    ws['B2'].value = "Time Deposit Manager"
    ws['B2'].font = Styles.title_font

    # Subtitle
    ws.merge_cells('B3:J3')
    ws['B3'].value = "Track and manage your term deposits, CDs, and fixed-income investments"
    ws['B3'].font = Styles.subheader_font

    # ========== SUMMARY CARDS ==========
    card_row = 5

    # Card 1: Total Deposits
    create_card(ws, card_row, 2, card_row + 2, 3)
    ws.cell(row=card_row, column=2, value="Total Deposits").font = Styles.small_font
    ws.cell(row=card_row + 1, column=2, value='=SUMIF(DepositsData!G:G,"Active",DepositsData!D:D)').font = Font(name='Calibri', size=20, bold=True, color=Colors.GREEN)
    ws.cell(row=card_row + 1, column=2).number_format = '"$"#,##0'
    ws.merge_cells(start_row=card_row + 1, start_column=2, end_row=card_row + 1, end_column=3)

    # Card 2: Average Rate
    create_card(ws, card_row, 4, card_row + 2, 5)
    ws.cell(row=card_row, column=4, value="Average Rate").font = Styles.small_font
    ws.cell(row=card_row + 1, column=4, value='=AVERAGEIF(DepositsData!G:G,"Active",DepositsData!E:E)').font = Font(name='Calibri', size=20, bold=True, color=Colors.BLUE)
    ws.cell(row=card_row + 1, column=4).number_format = '0.00%'
    ws.merge_cells(start_row=card_row + 1, start_column=4, end_row=card_row + 1, end_column=5)

    # Card 3: Expected Interest
    create_card(ws, card_row, 6, card_row + 2, 7)
    ws.cell(row=card_row, column=6, value="Expected Interest").font = Styles.small_font
    ws.cell(row=card_row + 1, column=6, value='=SUMIF(DepositsData!G:G,"Active",DepositsData!H:H)').font = Font(name='Calibri', size=20, bold=True, color=Colors.TEAL)
    ws.cell(row=card_row + 1, column=6).number_format = '"$"#,##0'
    ws.merge_cells(start_row=card_row + 1, start_column=6, end_row=card_row + 1, end_column=7)

    # Card 4: Maturing Soon
    create_card(ws, card_row, 8, card_row + 2, 9)
    ws.cell(row=card_row, column=8, value="Maturing (30 days)").font = Styles.small_font
    ws.cell(row=card_row + 1, column=8, value='=SUMPRODUCT((DepositsData!F2:F100>=TODAY())*(DepositsData!F2:F100<=TODAY()+30)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').font = Font(name='Calibri', size=20, bold=True, color=Colors.ORANGE)
    ws.cell(row=card_row + 1, column=8).number_format = '"$"#,##0'
    ws.merge_cells(start_row=card_row + 1, start_column=8, end_row=card_row + 1, end_column=9)

    # ========== DEPOSIT TABLE ==========
    table_row = 9
    ws.cell(row=table_row, column=2, value="Active Time Deposits").font = Styles.header_font

    # Headers
    headers = ["ID", "Bank/Institution", "Principal", "Rate", "Start Date", "Maturity", "Days Left", "Interest", "Status"]
    for i, header in enumerate(headers, 2):
        cell = ws.cell(row=table_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.GREEN, end_color=Colors.GREEN, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Data rows (linked to DepositsData)
    for i in range(15):
        row = table_row + 2 + i
        data_row = i + 2

        ws.cell(row=row, column=2, value=f"=DepositsData!A{data_row}")  # ID
        ws.cell(row=row, column=3, value=f"=DepositsData!B{data_row}")  # Bank
        ws.cell(row=row, column=4, value=f"=DepositsData!D{data_row}").number_format = '"$"#,##0.00'  # Principal
        ws.cell(row=row, column=5, value=f"=DepositsData!E{data_row}").number_format = '0.00%'  # Rate
        ws.cell(row=row, column=6, value=f"=DepositsData!C{data_row}").number_format = 'YYYY-MM-DD'  # Start
        ws.cell(row=row, column=7, value=f"=DepositsData!F{data_row}").number_format = 'YYYY-MM-DD'  # Maturity
        ws.cell(row=row, column=8, value=f'=IF(DepositsData!F{data_row}="","",MAX(0,DepositsData!F{data_row}-TODAY()))')  # Days Left
        ws.cell(row=row, column=9, value=f"=DepositsData!H{data_row}").number_format = '"$"#,##0.00'  # Interest
        ws.cell(row=row, column=10, value=f"=DepositsData!G{data_row}")  # Status

        for col in range(2, 11):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font
            ws.cell(row=row, column=col).alignment = Styles.center_align

    # Conditional formatting for days left
    ws.conditional_formatting.add(f'H{table_row + 2}:H{table_row + 16}',
        FormulaRule(formula=[f'H{table_row + 2}<=30'], fill=PatternFill(start_color=Colors.SOFT_ORANGE, end_color=Colors.SOFT_ORANGE, fill_type='solid')))
    ws.conditional_formatting.add(f'H{table_row + 2}:H{table_row + 16}',
        FormulaRule(formula=[f'H{table_row + 2}<=7'], fill=PatternFill(start_color=Colors.SOFT_RED, end_color=Colors.SOFT_RED, fill_type='solid')))

    # Status conditional formatting
    ws.conditional_formatting.add(f'J{table_row + 2}:J{table_row + 16}',
        FormulaRule(formula=[f'J{table_row + 2}="Active"'], fill=PatternFill(start_color=Colors.SOFT_GREEN, end_color=Colors.SOFT_GREEN, fill_type='solid')))
    ws.conditional_formatting.add(f'J{table_row + 2}:J{table_row + 16}',
        FormulaRule(formula=[f'J{table_row + 2}="Matured"'], fill=PatternFill(start_color=Colors.SOFT_BLUE, end_color=Colors.SOFT_BLUE, fill_type='solid')))

    # ========== MATURITY LADDER ==========
    ladder_row = 28
    ws.cell(row=ladder_row, column=2, value="Maturity Ladder").font = Styles.header_font

    ladder_headers = ["Period", "Count", "Amount", "% of Total"]
    for i, header in enumerate(ladder_headers, 2):
        cell = ws.cell(row=ladder_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.DARK_GRAY)
        cell.fill = Styles.light_gray_fill
        cell.alignment = Styles.center_align

    periods = [
        ("0-30 Days", '=SUMPRODUCT((DepositsData!F2:F100>=TODAY())*(DepositsData!F2:F100<=TODAY()+30)*(DepositsData!G2:G100="Active"))'),
        ("31-60 Days", '=SUMPRODUCT((DepositsData!F2:F100>TODAY()+30)*(DepositsData!F2:F100<=TODAY()+60)*(DepositsData!G2:G100="Active"))'),
        ("61-90 Days", '=SUMPRODUCT((DepositsData!F2:F100>TODAY()+60)*(DepositsData!F2:F100<=TODAY()+90)*(DepositsData!G2:G100="Active"))'),
        ("91-180 Days", '=SUMPRODUCT((DepositsData!F2:F100>TODAY()+90)*(DepositsData!F2:F100<=TODAY()+180)*(DepositsData!G2:G100="Active"))'),
        ("180+ Days", '=SUMPRODUCT((DepositsData!F2:F100>TODAY()+180)*(DepositsData!G2:G100="Active"))')
    ]

    for i, (period, count_formula) in enumerate(periods):
        row = ladder_row + 2 + i
        ws.cell(row=row, column=2, value=period).font = Styles.body_font
        ws.cell(row=row, column=3, value=count_formula)

        # Amount formula based on period
        if "0-30" in period:
            ws.cell(row=row, column=4, value='=SUMPRODUCT((DepositsData!F2:F100>=TODAY())*(DepositsData!F2:F100<=TODAY()+30)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').number_format = '"$"#,##0'
        elif "31-60" in period:
            ws.cell(row=row, column=4, value='=SUMPRODUCT((DepositsData!F2:F100>TODAY()+30)*(DepositsData!F2:F100<=TODAY()+60)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').number_format = '"$"#,##0'
        elif "61-90" in period:
            ws.cell(row=row, column=4, value='=SUMPRODUCT((DepositsData!F2:F100>TODAY()+60)*(DepositsData!F2:F100<=TODAY()+90)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').number_format = '"$"#,##0'
        elif "91-180" in period:
            ws.cell(row=row, column=4, value='=SUMPRODUCT((DepositsData!F2:F100>TODAY()+90)*(DepositsData!F2:F100<=TODAY()+180)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').number_format = '"$"#,##0'
        else:
            ws.cell(row=row, column=4, value='=SUMPRODUCT((DepositsData!F2:F100>TODAY()+180)*(DepositsData!G2:G100="Active")*DepositsData!D2:D100)').number_format = '"$"#,##0'

        ws.cell(row=row, column=5, value=f'=IFERROR(D{row}/SUMIF(DepositsData!G:G,"Active",DepositsData!D:D),0)').number_format = '0.0%'

        for col in range(2, 6):
            ws.cell(row=row, column=col).border = Styles.thin_border


# ============================================================================
# CASH FLOW POSITIONS SHEET
# ============================================================================
def create_cash_flow_positions_sheet(wb):
    """Create comprehensive cash flow positions (daily, weekly, monthly)"""
    ws = wb.create_sheet("CashFlowPositions")
    ws.sheet_view.showGridLines = False

    # Set column widths
    column_widths = [3, 18, 18, 18, 18, 18, 18, 18, 3]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Title
    ws.merge_cells('B2:H2')
    ws['B2'].value = "Cash Flow Positions"
    ws['B2'].font = Styles.title_font

    ws.merge_cells('B3:H3')
    ws['B3'].value = "Comprehensive view of cash inflows, outflows, and positions"
    ws['B3'].font = Styles.subheader_font

    # ========== PERIOD SUMMARY TABLE ==========
    summary_row = 5
    ws.cell(row=summary_row, column=2, value="Position Summary").font = Styles.header_font

    headers = ["Period", "Start Balance", "Inflows", "Outflows", "Net Flow", "End Balance", "Change %"]
    for i, header in enumerate(headers, 2):
        cell = ws.cell(row=summary_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.BLUE, end_color=Colors.BLUE, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Period rows with formulas
    period_data = [
        ("Today", "TODAY()", "TODAY()"),
        ("This Week", "TODAY()-WEEKDAY(TODAY(),2)+1", "TODAY()-WEEKDAY(TODAY(),2)+7"),
        ("This Month", "DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "EOMONTH(TODAY(),0)"),
        ("Next 30 Days", "TODAY()", "TODAY()+30"),
        ("Next 60 Days", "TODAY()", "TODAY()+60"),
        ("Next 90 Days", "TODAY()", "TODAY()+90")
    ]

    for i, (period, start_date, end_date) in enumerate(period_data):
        row = summary_row + 2 + i

        ws.cell(row=row, column=2, value=period).font = Styles.body_font

        # Start Balance (from CashTransactions running balance)
        ws.cell(row=row, column=3, value="=CashTransactions!J2").number_format = '"$"#,##0'

        # Inflows
        ws.cell(row=row, column=4, value=f'=SUMPRODUCT((CashTransactions!B2:B500>={start_date})*(CashTransactions!B2:B500<={end_date})*(CashTransactions!D2:D500>0)*CashTransactions!D2:D500)+SUMPRODUCT((DepositsData!F2:F100>={start_date})*(DepositsData!F2:F100<={end_date})*(DepositsData!G2:G100="Active")*(DepositsData!D2:D100+DepositsData!H2:H100))').number_format = '"$"#,##0'

        # Outflows
        ws.cell(row=row, column=5, value=f'=SUMPRODUCT((PaymentsData!B2:B500>={start_date})*(PaymentsData!B2:B500<={end_date})*(PaymentsData!F2:F500<>"Cancelled")*PaymentsData!D2:D500)+ABS(SUMPRODUCT((CashTransactions!B2:B500>={start_date})*(CashTransactions!B2:B500<={end_date})*(CashTransactions!D2:D500<0)*CashTransactions!D2:D500))').number_format = '"$"#,##0'

        # Net Flow
        ws.cell(row=row, column=6, value=f'=D{row}-E{row}').number_format = '"$"#,##0'

        # End Balance
        ws.cell(row=row, column=7, value=f'=C{row}+F{row}').number_format = '"$"#,##0'

        # Change %
        ws.cell(row=row, column=8, value=f'=IFERROR(F{row}/C{row},0)').number_format = '0.0%'

        for col in range(2, 9):
            ws.cell(row=row, column=col).border = Styles.thin_border

    # Conditional formatting for net flow
    ws.conditional_formatting.add(f'F{summary_row + 2}:F{summary_row + 7}',
        FormulaRule(formula=[f'F{summary_row + 2}>=0'], fill=PatternFill(start_color=Colors.SOFT_GREEN, end_color=Colors.SOFT_GREEN, fill_type='solid')))
    ws.conditional_formatting.add(f'F{summary_row + 2}:F{summary_row + 7}',
        FormulaRule(formula=[f'F{summary_row + 2}<0'], fill=PatternFill(start_color=Colors.SOFT_RED, end_color=Colors.SOFT_RED, fill_type='solid')))

    # ========== DAILY POSITIONS (Next 14 Days) ==========
    daily_row = 14
    ws.cell(row=daily_row, column=2, value="Daily Positions (Next 14 Days)").font = Styles.header_font

    daily_headers = ["Date", "Day", "Inflows", "Outflows", "Net", "Running Balance"]
    for i, header in enumerate(daily_headers, 2):
        cell = ws.cell(row=daily_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.TEAL, end_color=Colors.TEAL, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    for i in range(14):
        row = daily_row + 2 + i

        # Date
        ws.cell(row=row, column=2, value=f'=TODAY()+{i}').number_format = 'MMM DD'

        # Day name
        ws.cell(row=row, column=3, value=f'=TEXT(B{row},"dddd")')

        # Inflows
        ws.cell(row=row, column=4, value=f'=SUMPRODUCT((CashTransactions!B:B=B{row})*(CashTransactions!D:D>0)*CashTransactions!D:D)+SUMPRODUCT((DepositsData!F:F=B{row})*(DepositsData!G:G="Active")*(DepositsData!D:D+DepositsData!H:H))').number_format = '"$"#,##0'

        # Outflows
        ws.cell(row=row, column=5, value=f'=SUMPRODUCT((PaymentsData!B:B=B{row})*(PaymentsData!F:F<>"Cancelled")*PaymentsData!D:D)+ABS(SUMPRODUCT((CashTransactions!B:B=B{row})*(CashTransactions!D:D<0)*CashTransactions!D:D))').number_format = '"$"#,##0'

        # Net
        ws.cell(row=row, column=6, value=f'=D{row}-E{row}').number_format = '"$"#,##0'

        # Running Balance
        if i == 0:
            ws.cell(row=row, column=7, value=f'=CashTransactions!J2+F{row}').number_format = '"$"#,##0'
        else:
            ws.cell(row=row, column=7, value=f'=G{row-1}+F{row}').number_format = '"$"#,##0'

        for col in range(2, 8):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font

        # Highlight today
        if i == 0:
            for col in range(2, 8):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=Colors.SOFT_BLUE, end_color=Colors.SOFT_BLUE, fill_type='solid')

        # Highlight weekends
        ws.conditional_formatting.add(f'B{row}:G{row}',
            FormulaRule(formula=[f'OR(WEEKDAY(B{row})=1,WEEKDAY(B{row})=7)'],
                       fill=PatternFill(start_color=Colors.LIGHT_GRAY, end_color=Colors.LIGHT_GRAY, fill_type='solid')))

    # ========== WEEKLY POSITIONS (Next 8 Weeks) ==========
    weekly_row = 32
    ws.cell(row=weekly_row, column=2, value="Weekly Positions (Next 8 Weeks)").font = Styles.header_font

    weekly_headers = ["Week Starting", "Week #", "Inflows", "Outflows", "Net", "End Balance"]
    for i, header in enumerate(weekly_headers, 2):
        cell = ws.cell(row=weekly_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.PURPLE, end_color=Colors.PURPLE, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    for i in range(8):
        row = weekly_row + 2 + i
        week_start = i * 7
        week_end = week_start + 6

        # Week start date
        ws.cell(row=row, column=2, value=f'=TODAY()-WEEKDAY(TODAY(),2)+1+{week_start}').number_format = 'MMM DD'

        # Week number
        ws.cell(row=row, column=3, value=f'=WEEKNUM(B{row})')

        # Inflows
        ws.cell(row=row, column=4, value=f'=SUMPRODUCT((CashTransactions!B:B>=B{row})*(CashTransactions!B:B<=B{row}+6)*(CashTransactions!D:D>0)*CashTransactions!D:D)+SUMPRODUCT((DepositsData!F:F>=B{row})*(DepositsData!F:F<=B{row}+6)*(DepositsData!G:G="Active")*(DepositsData!D:D+DepositsData!H:H))').number_format = '"$"#,##0'

        # Outflows
        ws.cell(row=row, column=5, value=f'=SUMPRODUCT((PaymentsData!B:B>=B{row})*(PaymentsData!B:B<=B{row}+6)*(PaymentsData!F:F<>"Cancelled")*PaymentsData!D:D)+ABS(SUMPRODUCT((CashTransactions!B:B>=B{row})*(CashTransactions!B:B<=B{row}+6)*(CashTransactions!D:D<0)*CashTransactions!D:D))').number_format = '"$"#,##0'

        # Net
        ws.cell(row=row, column=6, value=f'=D{row}-E{row}').number_format = '"$"#,##0'

        # End Balance
        if i == 0:
            ws.cell(row=row, column=7, value=f'=CashTransactions!J2+F{row}').number_format = '"$"#,##0'
        else:
            ws.cell(row=row, column=7, value=f'=G{row-1}+F{row}').number_format = '"$"#,##0'

        for col in range(2, 8):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font

    # ========== MONTHLY POSITIONS (Next 6 Months) ==========
    monthly_row = 44
    ws.cell(row=monthly_row, column=2, value="Monthly Positions (Next 6 Months)").font = Styles.header_font

    monthly_headers = ["Month", "Inflows", "Outflows", "Net Flow", "End Balance", "Deposits Maturing"]
    for i, header in enumerate(monthly_headers, 2):
        cell = ws.cell(row=monthly_row + 1, column=i, value=header)
        cell.font = Font(name='Calibri', size=10, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.ORANGE, end_color=Colors.ORANGE, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    for i in range(6):
        row = monthly_row + 2 + i

        # Month name
        ws.cell(row=row, column=2, value=f'=TEXT(EOMONTH(TODAY(),{i}),"MMMM YYYY")')

        # Inflows
        ws.cell(row=row, column=3, value=f'=SUMPRODUCT((MONTH(CashTransactions!B:B)=MONTH(EOMONTH(TODAY(),{i})))*(YEAR(CashTransactions!B:B)=YEAR(EOMONTH(TODAY(),{i})))*(CashTransactions!D:D>0)*CashTransactions!D:D)').number_format = '"$"#,##0'

        # Outflows
        ws.cell(row=row, column=4, value=f'=SUMPRODUCT((MONTH(PaymentsData!B:B)=MONTH(EOMONTH(TODAY(),{i})))*(YEAR(PaymentsData!B:B)=YEAR(EOMONTH(TODAY(),{i})))*(PaymentsData!F:F<>"Cancelled")*PaymentsData!D:D)').number_format = '"$"#,##0'

        # Net Flow
        ws.cell(row=row, column=5, value=f'=C{row}-D{row}').number_format = '"$"#,##0'

        # End Balance
        if i == 0:
            ws.cell(row=row, column=6, value=f'=CashTransactions!J2+E{row}').number_format = '"$"#,##0'
        else:
            ws.cell(row=row, column=6, value=f'=F{row-1}+E{row}').number_format = '"$"#,##0'

        # Deposits Maturing
        ws.cell(row=row, column=7, value=f'=SUMPRODUCT((MONTH(DepositsData!F:F)=MONTH(EOMONTH(TODAY(),{i})))*(YEAR(DepositsData!F:F)=YEAR(EOMONTH(TODAY(),{i})))*(DepositsData!G:G="Active")*(DepositsData!D:D+DepositsData!H:H))').number_format = '"$"#,##0'

        for col in range(2, 8):
            ws.cell(row=row, column=col).border = Styles.thin_border
            ws.cell(row=row, column=col).font = Styles.body_font


# ============================================================================
# DATA SHEETS
# ============================================================================
def create_payments_data_sheet(wb):
    """Create the payments data entry sheet"""
    ws = wb.create_sheet("PaymentsData")

    # Set column widths
    column_widths = [10, 15, 25, 15, 18, 12, 30]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Headers
    headers = ["Payment ID", "Due Date", "Payee", "Amount", "Category", "Status", "Notes"]
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font = Font(name='Calibri', size=11, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.BLUE, end_color=Colors.BLUE, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Status data validation
    status_dv = DataValidation(type="list", formula1='"Scheduled,Pending,Paid,Overdue,Cancelled"', allow_blank=True)
    ws.add_data_validation(status_dv)
    status_dv.add('F2:F500')

    # Category data validation
    category_dv = DataValidation(type="list", formula1='"Payroll,Vendor,Utilities,Rent,Insurance,Taxes,Supplies,Services,Other"', allow_blank=True)
    ws.add_data_validation(category_dv)
    category_dv.add('E2:E500')

    # Sample data
    sample_payments = [
        ["PAY001", datetime.now() + timedelta(days=2), "ABC Suppliers", 15000, "Vendor", "Scheduled", "Monthly supplies"],
        ["PAY002", datetime.now() + timedelta(days=5), "City Utilities", 3500, "Utilities", "Pending", "Electric & Water"],
        ["PAY003", datetime.now() + timedelta(days=7), "Office Rent LLC", 25000, "Rent", "Scheduled", "Monthly rent"],
        ["PAY004", datetime.now() + timedelta(days=10), "Insurance Co.", 5000, "Insurance", "Scheduled", "Quarterly premium"],
        ["PAY005", datetime.now() + timedelta(days=12), "Payroll", 85000, "Payroll", "Scheduled", "Bi-weekly payroll"],
        ["PAY006", datetime.now() + timedelta(days=15), "Tech Solutions", 8500, "Services", "Pending", "IT support"],
        ["PAY007", datetime.now() + timedelta(days=18), "Marketing Agency", 12000, "Services", "Scheduled", "Monthly retainer"],
        ["PAY008", datetime.now() + timedelta(days=20), "Office Supplies Inc", 2500, "Supplies", "Scheduled", "Office materials"],
        ["PAY009", datetime.now() - timedelta(days=3), "Contractor XYZ", 7500, "Services", "Overdue", "Project milestone"],
        ["PAY010", datetime.now() - timedelta(days=5), "Equipment Lease", 4000, "Other", "Paid", "Monthly lease"],
        ["PAY011", datetime.now() + timedelta(days=25), "Payroll", 85000, "Payroll", "Scheduled", "Bi-weekly payroll"],
        ["PAY012", datetime.now() + timedelta(days=28), "Tax Authority", 45000, "Taxes", "Scheduled", "Quarterly taxes"],
    ]

    for row_num, payment in enumerate(sample_payments, 2):
        for col_num, value in enumerate(payment, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = Styles.thin_border
            cell.font = Styles.body_font
            if col_num == 2:  # Date
                cell.number_format = 'YYYY-MM-DD'
            elif col_num == 4:  # Amount
                cell.number_format = '"$"#,##0.00'

    # Conditional formatting for status
    ws.conditional_formatting.add('F2:F500',
        FormulaRule(formula=['F2="Paid"'], fill=PatternFill(start_color=Colors.SOFT_GREEN, end_color=Colors.SOFT_GREEN, fill_type='solid')))
    ws.conditional_formatting.add('F2:F500',
        FormulaRule(formula=['F2="Overdue"'], fill=PatternFill(start_color=Colors.SOFT_RED, end_color=Colors.SOFT_RED, fill_type='solid')))
    ws.conditional_formatting.add('F2:F500',
        FormulaRule(formula=['F2="Pending"'], fill=PatternFill(start_color=Colors.SOFT_ORANGE, end_color=Colors.SOFT_ORANGE, fill_type='solid')))


def create_deposits_data_sheet(wb):
    """Create the time deposits data entry sheet"""
    ws = wb.create_sheet("DepositsData")

    # Set column widths
    column_widths = [12, 25, 15, 18, 10, 15, 12, 15, 30]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Headers
    headers = ["Deposit ID", "Bank/Institution", "Start Date", "Principal", "Rate", "Maturity Date", "Status", "Interest", "Notes"]
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font = Font(name='Calibri', size=11, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.GREEN, end_color=Colors.GREEN, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Status data validation
    status_dv = DataValidation(type="list", formula1='"Active,Matured,Withdrawn,Renewed"', allow_blank=True)
    ws.add_data_validation(status_dv)
    status_dv.add('G2:G100')

    # Sample data
    sample_deposits = [
        ["TD001", "First National Bank", datetime.now() - timedelta(days=90), 500000, 0.045, datetime.now() + timedelta(days=90), "Active", None, "6-month CD"],
        ["TD002", "City Credit Union", datetime.now() - timedelta(days=60), 250000, 0.0425, datetime.now() + timedelta(days=120), "Active", None, "6-month term"],
        ["TD003", "State Bank", datetime.now() - timedelta(days=30), 750000, 0.05, datetime.now() + timedelta(days=335), "Active", None, "12-month CD"],
        ["TD004", "Federal Savings", datetime.now() - timedelta(days=180), 300000, 0.04, datetime.now() + timedelta(days=5), "Active", None, "Maturing soon"],
        ["TD005", "Capital One", datetime.now() - timedelta(days=120), 400000, 0.0475, datetime.now() + timedelta(days=60), "Active", None, "6-month CD"],
        ["TD006", "Wells Fargo", datetime.now() - timedelta(days=365), 200000, 0.035, datetime.now() - timedelta(days=5), "Matured", None, "Awaiting reinvestment"],
        ["TD007", "Chase Bank", datetime.now() - timedelta(days=45), 600000, 0.0525, datetime.now() + timedelta(days=135), "Active", None, "6-month CD"],
        ["TD008", "Bank of America", datetime.now() - timedelta(days=15), 350000, 0.048, datetime.now() + timedelta(days=165), "Active", None, "6-month term"],
    ]

    for row_num, deposit in enumerate(sample_deposits, 2):
        for col_num, value in enumerate(deposit, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = Styles.thin_border
            cell.font = Styles.body_font
            if col_num in [3, 6]:  # Dates
                cell.number_format = 'YYYY-MM-DD'
            elif col_num == 4:  # Principal
                cell.number_format = '"$"#,##0.00'
            elif col_num == 5:  # Rate
                cell.number_format = '0.00%'

        # Interest calculation formula
        interest_cell = ws.cell(row=row_num, column=8, value=f'=D{row_num}*E{row_num}*(F{row_num}-C{row_num})/365')
        interest_cell.number_format = '"$"#,##0.00'
        interest_cell.border = Styles.thin_border


def create_cash_transactions_sheet(wb):
    """Create the cash transactions data sheet"""
    ws = wb.create_sheet("CashTransactions")

    # Set column widths
    column_widths = [12, 15, 12, 18, 25, 18, 15, 15, 30, 18]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Headers
    headers = ["Trans ID", "Date", "Type", "Amount", "Description", "Category", "Account", "Reference", "Notes", "Running Balance"]
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font = Font(name='Calibri', size=11, bold=True, color=Colors.WHITE)
        cell.fill = PatternFill(start_color=Colors.DARK_GRAY, end_color=Colors.DARK_GRAY, fill_type='solid')
        cell.alignment = Styles.center_align
        cell.border = Styles.thin_border

    # Type data validation
    type_dv = DataValidation(type="list", formula1='"Inflow,Outflow"', allow_blank=True)
    ws.add_data_validation(type_dv)
    type_dv.add('C2:C500')

    # Sample transactions - starting balance is assumed to be 2,500,000
    starting_balance = 2500000
    sample_transactions = [
        ["TRX001", datetime.now() - timedelta(days=10), "Inflow", 150000, "Customer Payment - ABC Corp", "Revenue", "Operating", "INV-2024-001", ""],
        ["TRX002", datetime.now() - timedelta(days=9), "Outflow", -45000, "Supplier Payment", "COGS", "Operating", "PO-2024-123", ""],
        ["TRX003", datetime.now() - timedelta(days=8), "Inflow", 200000, "Customer Payment - XYZ Inc", "Revenue", "Operating", "INV-2024-002", ""],
        ["TRX004", datetime.now() - timedelta(days=7), "Outflow", -85000, "Payroll", "Payroll", "Operating", "PR-2024-24", "Bi-weekly"],
        ["TRX005", datetime.now() - timedelta(days=5), "Inflow", 75000, "Customer Payment - DEF Ltd", "Revenue", "Operating", "INV-2024-003", ""],
        ["TRX006", datetime.now() - timedelta(days=4), "Outflow", -25000, "Office Rent", "Rent", "Operating", "RENT-DEC", "Monthly"],
        ["TRX007", datetime.now() - timedelta(days=3), "Inflow", 500000, "TD Maturity - Wells Fargo", "Interest", "Investments", "TD006", "Plus interest"],
        ["TRX008", datetime.now() - timedelta(days=2), "Outflow", -350000, "New Time Deposit", "Investment", "Investments", "TD008", "6-month term"],
        ["TRX009", datetime.now() - timedelta(days=1), "Inflow", 125000, "Customer Payment - GHI Co", "Revenue", "Operating", "INV-2024-004", ""],
        ["TRX010", datetime.now(), "Outflow", -15000, "Utility Bills", "Utilities", "Operating", "UTIL-DEC", "Electric, Water, Gas"],
    ]

    running_balance = starting_balance
    for row_num, trans in enumerate(sample_transactions, 2):
        for col_num, value in enumerate(trans, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = Styles.thin_border
            cell.font = Styles.body_font
            if col_num == 2:  # Date
                cell.number_format = 'YYYY-MM-DD'
            elif col_num == 4:  # Amount
                cell.number_format = '"$"#,##0.00'

        # Running balance
        running_balance += trans[3]
        balance_cell = ws.cell(row=row_num, column=10, value=running_balance)
        balance_cell.number_format = '"$"#,##0.00'
        balance_cell.border = Styles.thin_border

    # Conditional formatting for amount (green for inflow, red for outflow)
    ws.conditional_formatting.add('D2:D500',
        FormulaRule(formula=['D2>0'], font=Font(color=Colors.GREEN)))
    ws.conditional_formatting.add('D2:D500',
        FormulaRule(formula=['D2<0'], font=Font(color=Colors.RED)))


def create_settings_sheet(wb):
    """Create settings and configuration sheet"""
    ws = wb.create_sheet("Settings")
    ws.sheet_view.showGridLines = False

    # Set column widths
    for i in range(1, 6):
        ws.column_dimensions[get_column_letter(i)].width = 25

    # Title
    ws.merge_cells('B2:D2')
    ws['B2'].value = "Settings & Configuration"
    ws['B2'].font = Styles.title_font

    # ========== COMPANY INFO ==========
    ws.cell(row=4, column=2, value="Company Information").font = Styles.header_font

    settings = [
        ("Company Name:", "Your Company Name"),
        ("Currency:", "USD"),
        ("Fiscal Year Start:", "January"),
        ("Default Payment Terms:", "Net 30"),
    ]

    for i, (label, value) in enumerate(settings):
        ws.cell(row=5 + i, column=2, value=label).font = Styles.body_font
        ws.cell(row=5 + i, column=3, value=value).font = Styles.body_font
        ws.cell(row=5 + i, column=3).fill = Styles.light_gray_fill

    # ========== CATEGORIES ==========
    ws.cell(row=11, column=2, value="Payment Categories").font = Styles.header_font

    categories = ["Payroll", "Vendor", "Utilities", "Rent", "Insurance", "Taxes", "Supplies", "Services", "Other"]
    for i, cat in enumerate(categories):
        ws.cell(row=12 + i, column=2, value=cat).font = Styles.body_font

    # ========== STATUS DEFINITIONS ==========
    ws.cell(row=11, column=4, value="Status Definitions").font = Styles.header_font

    statuses = [
        ("Scheduled", "Future payment, not yet processed"),
        ("Pending", "Payment initiated, awaiting completion"),
        ("Paid", "Payment completed successfully"),
        ("Overdue", "Payment past due date"),
        ("Cancelled", "Payment cancelled"),
    ]

    for i, (status, desc) in enumerate(statuses):
        ws.cell(row=12 + i, column=4, value=status).font = Font(name='Calibri', size=11, bold=True)
        ws.cell(row=12 + i, column=5, value=desc).font = Styles.small_font

    # ========== THRESHOLDS ==========
    ws.cell(row=22, column=2, value="Alert Thresholds").font = Styles.header_font

    thresholds = [
        ("Minimum Cash Balance:", 100000, "Alert when balance falls below"),
        ("Days Before Maturity Alert:", 30, "Days before deposit maturity"),
        ("Days Before Payment Due:", 7, "Days before payment reminder"),
    ]

    for i, (label, value, desc) in enumerate(thresholds):
        ws.cell(row=23 + i, column=2, value=label).font = Styles.body_font
        ws.cell(row=23 + i, column=3, value=value).font = Styles.body_font
        ws.cell(row=23 + i, column=3).fill = Styles.light_gray_fill
        ws.cell(row=23 + i, column=4, value=desc).font = Styles.small_font


def main():
    """Main function to generate the workbook"""
    print("Creating Cash Management Controller Workbook...")
    print("=" * 50)

    wb = create_workbook()

    # Set active sheet to Dashboard
    wb.active = wb["Dashboard"]

    # Save the workbook
    output_file = "Cash_Management_Controller.xlsx"
    wb.save(output_file)

    print(f"\nWorkbook created successfully: {output_file}")
    print("\nSheets included:")
    for sheet in wb.sheetnames:
        print(f"  - {sheet}")

    print("\nFeatures:")
    print("  - Apple-inspired modern design")
    print("  - Payment Calendar with day boxes and status")
    print("  - Time Deposit Manager with maturity tracking")
    print("  - Cash Flow Positions (Daily/Weekly/Monthly)")
    print("  - All data from internal sheet sources")
    print("  - Conditional formatting for visual status")
    print("  - Data validation for consistent entry")

    return output_file


if __name__ == "__main__":
    main()
