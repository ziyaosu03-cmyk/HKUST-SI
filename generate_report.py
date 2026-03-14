"""
ISOM5460 Assignment 2 – Conveyor Belt Project
Executive Summary Generator
Generates a 2-page PDF executive summary with EVM analysis and appendix charts.
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.graphics.shapes import Drawing as GfxDrawing, Rect, String, Line, Group
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.validators import Auto
from reportlab.graphics import renderPDF
import os

# ─── Colours ─────────────────────────────────────────────────────────────────
BLUE_DARK  = colors.HexColor("#1a3c6e")
BLUE_MID   = colors.HexColor("#2d6eba")
BLUE_LIGHT = colors.HexColor("#dce9f8")
ORANGE     = colors.HexColor("#e07b35")
GREEN      = colors.HexColor("#2e8b57")
RED        = colors.HexColor("#c0392b")
GREY_LIGHT = colors.HexColor("#f4f4f4")
GREY_MID   = colors.HexColor("#cccccc")

# ─── Key Numbers ─────────────────────────────────────────────────────────────
BAC   = 1_051_200.00
PV    =   690_160.00
EV    =   691_634.19
AC    =   721_200.00
SV    =     1_474.19
SPI   =       1.002
CV    =  -29_565.81
CPI   =       0.959
EAC   = 1_096_136.44
ETC   =   374_936.44
VAC   =  -44_936.44
TCPI  =       1.090

OUTPUT = "/workspace/Executive_Summary_CBP.pdf"

# ─── Document setup ───────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    rightMargin=2*cm, leftMargin=2*cm,
    topMargin=2*cm, bottomMargin=2*cm,
    title="Conveyor Belt Project – Executive Summary",
    author="ISOM5460 Assignment 2",
)

styles = getSampleStyleSheet()

def S(name, **kw):
    """Return a ParagraphStyle, overriding any attributes."""
    base = styles[name]
    return ParagraphStyle(name + "_custom", parent=base, **kw)

h1   = S("Heading1",  fontSize=16, textColor=BLUE_DARK, spaceAfter=4,
          spaceBefore=6, fontName="Helvetica-Bold")
h2   = S("Heading2",  fontSize=12, textColor=BLUE_DARK, spaceAfter=3,
          spaceBefore=8, fontName="Helvetica-Bold")
h3   = S("Heading3",  fontSize=10, textColor=BLUE_MID,  spaceAfter=2,
          spaceBefore=4, fontName="Helvetica-Bold")
body = S("Normal",    fontSize=9,  leading=13, spaceAfter=4,
          fontName="Helvetica")
small= S("Normal",    fontSize=8,  leading=11, textColor=colors.grey,
          fontName="Helvetica")
bold_body = S("Normal", fontSize=9, leading=13, fontName="Helvetica-Bold")

def hr():
    return HRFlowable(width="100%", thickness=1, color=BLUE_MID, spaceAfter=4)

def metric_table(rows, col_widths=None):
    """Two-column key-value table for metrics."""
    if col_widths is None:
        col_widths = [9*cm, 7*cm]
    t = Table(rows, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BLUE_DARK),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [GREY_LIGHT, colors.white]),
        ("GRID",       (0,0), (-1,-1), 0.4, GREY_MID),
        ("LEFTPADDING",  (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("ALIGN",      (1,1), (1,-1), "RIGHT"),
    ]))
    return t

# ─── Helper: bar chart drawing ────────────────────────────────────────────────
def planned_vs_actual_bar_chart():
    """
    Grouped bar chart: Planned Cost vs Actual Cost for each task.
    Returns a Flowable Drawing.
    """
    task_names = [
        "HW Spec", "HW Design", "HW Doc", "Prototypes",
        "Kernel Spec", "Disk Drivers", "SerIO Drivers", "Mem Mgmt",
        "OS Doc", "Util Spec", "Routine Util", "Cmplx Util",
        "Util Doc", "Arch Dec", "Intg Phase1"
    ]
    planned_costs = [68000, 95200, 38400, 89600, 16000, 112000, 72800,
                     50400, 32000, 20400, 57600, 44800, 25600, 20000, 56000]
    actual_costs  = [61200, 91120, 34560, 38080, 20000, 123200, 19040,
                     54880, 32000, 21760, 62400, 53200, 28160, 20000, 61600]

    d = GfxDrawing(500, 200)
    bc = VerticalBarChart()
    bc.x = 60
    bc.y = 30
    bc.height = 140
    bc.width  = 420
    bc.data   = [planned_costs, actual_costs]
    bc.groupSpacing = 4
    bc.barSpacing   = 1
    bc.bars[0].fillColor = BLUE_MID
    bc.bars[1].fillColor = ORANGE
    bc.valueAxis.valueMin   = 0
    bc.valueAxis.valueMax   = 140000
    bc.valueAxis.valueStep  = 20000
    bc.valueAxis.labelTextFormat = lambda v: f"${int(v/1000)}K"
    bc.categoryAxis.categoryNames = task_names
    bc.categoryAxis.labels.angle  = 45
    bc.categoryAxis.labels.fontSize = 6
    bc.categoryAxis.labels.dx = -8

    legend = Legend()
    legend.x = 60
    legend.y = 185
    legend.deltax = 70
    legend.fontName = "Helvetica"
    legend.fontSize = 8
    legend.colorNamePairs = [(BLUE_MID, "Planned Cost"), (ORANGE, "Actual Cost")]
    d.add(bc)
    d.add(legend)
    return d

def evm_line_chart():
    """S-curve: PV, EV, AC over time (simplified monthly)."""
    # Simplified quarterly values (Q1 2010 … Q4 2010, Jan 2011)
    quarters = ["Q1'10","Q2'10","Q3'10","Q4'10","Jan'11"]
    pv_vals  = [56000,  320000, 540000, 636400, 690160]
    ev_vals  = [56000,  310000, 530000, 635000, 691634]
    ac_vals  = [56000,  325000, 548000, 660000, 721200]

    d = GfxDrawing(500, 180)
    lc = HorizontalLineChart()
    lc.x      = 60
    lc.y      = 30
    lc.height = 120
    lc.width  = 400
    lc.data   = [pv_vals, ev_vals, ac_vals]
    lc.lines[0].strokeColor = BLUE_MID
    lc.lines[0].strokeWidth = 2
    lc.lines[1].strokeColor = GREEN
    lc.lines[1].strokeWidth = 2
    lc.lines[2].strokeColor = ORANGE
    lc.lines[2].strokeWidth = 2
    lc.categoryAxis.categoryNames = quarters
    lc.categoryAxis.labels.fontSize = 8
    lc.valueAxis.valueMin  = 0
    lc.valueAxis.valueMax  = 800000
    lc.valueAxis.valueStep = 200000
    lc.valueAxis.labelTextFormat = lambda v: f"${int(v/1000)}K"
    lc.valueAxis.labels.fontSize = 8

    legend = Legend()
    legend.x      = 60
    legend.y      = 165
    legend.deltax = 70
    legend.fontName = "Helvetica"
    legend.fontSize = 8
    legend.colorNamePairs = [
        (BLUE_MID, "PV (Planned)"),
        (GREEN,    "EV (Earned)"),
        (ORANGE,   "AC (Actual)"),
    ]
    d.add(lc)
    d.add(legend)
    return d

def cash_flow_bar():
    """Monthly cumulative planned cost distribution."""
    months = ["Feb'10","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan'11","...","Feb'12"]
    # Approximate monthly planned expenditure
    monthly = [56000, 120000, 90000, 60000, 80000, 85000, 50000, 60000, 45000, 90000, 55000, 54160, 220000, 80840]
    cumulative = []
    c = 0
    for v in monthly:
        c += v
        cumulative.append(c)

    d = GfxDrawing(500, 180)
    bc = VerticalBarChart()
    bc.x = 60
    bc.y = 30
    bc.height = 120
    bc.width  = 400
    bc.data   = [monthly]
    bc.bars[0].fillColor = BLUE_MID
    bc.barSpacing = 2
    bc.valueAxis.valueMin  = 0
    bc.valueAxis.valueMax  = 250000
    bc.valueAxis.valueStep = 50000
    bc.valueAxis.labelTextFormat = lambda v: f"${int(v/1000)}K"
    bc.valueAxis.labels.fontSize = 7
    bc.categoryAxis.categoryNames = months
    bc.categoryAxis.labels.angle  = 45
    bc.categoryAxis.labels.fontSize = 6
    bc.categoryAxis.labels.dx = -8
    d.add(bc)

    title = String(260, 170, "Monthly Planned Cost Distribution", fontSize=9,
                   fontName="Helvetica-Bold", fillColor=BLUE_DARK)
    d.add(title)
    return d

# ─── Build story ──────────────────────────────────────────────────────────────
story = []

# ── Title block ──────────────────────────────────────────────────────────────
story.append(Paragraph("Conveyor Belt Project", h1))
story.append(Paragraph("Executive Summary – Project Status as of January 1, 2011", h2))
story.append(Paragraph("ISOM5460 Assignment 2 &nbsp;|&nbsp; Prepared for: CEO & Board of Directors",
                        small))
story.append(hr())
story.append(Spacer(1, 0.2*cm))

# ── Part 1: Project Cost & Baseline ──────────────────────────────────────────
story.append(Paragraph("Part 1 – Project Baseline (CBP_P3B.mpp)", h2))

story.append(Paragraph(
    "<b>Q1 – Project Cost.</b>  "
    "Based on the updated resource plan in <i>CBP_P3B.mpp</i> (which incorporates two "
    "extra internal Development team members and one external Xdevelopment contractor), "
    "the total <b>Budget at Completion (BAC) is $1,051,200</b>.  "
    "The project spans from <b>January 4, 2010</b> to <b>February 2, 2012</b> "
    "(530 working days).  Resources include Design ($100/h), Development ($70/h), "
    "Documentation ($60/h), Assembly/Test ($70/h), Purchasing ($40/h), and "
    "Xdevelopment external contractor ($120/h).",
    body))

story.append(Paragraph(
    "<b>Q2 – Cost Distribution.</b>  "
    "The largest cost centres are the <b>Hardware sub-project ($326,400, 31%)</b> and "
    "<b>Operating System sub-project ($333,600, 32%)</b>, together accounting for 63% of "
    "the total budget.  System Integration contributes $209,200 (20%), while Utilities "
    "account for $182,000 (17%).  Expenditure is heaviest in the first year (Jan 2010 – "
    "Dec 2010) when all foundation development tasks run in parallel, then tapers into "
    "2011 as integration and testing phases begin.  "
    "See <u>Appendix A</u> for the monthly cash-flow chart and sub-project breakdown.",
    body))

cost_rows = [
    ["Sub-Project",       "Planned Cost",  "% of Total"],
    ["Hardware",          "$326,400",      "31.0%"],
    ["Operating System",  "$333,600",      "31.7%"],
    ["Utilities",         "$182,000",      "17.3%"],
    ["System Integration","$209,200",      "19.9%"],
    ["TOTAL (BAC)",       "$1,051,200",    "100.0%"],
]
t = Table(cost_rows, colWidths=[7*cm, 5*cm, 5*cm], hAlign="LEFT")
t.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
    ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
    ("FONTNAME",    (0,5),(-1,5), "Helvetica-Bold"),
    ("BACKGROUND",  (0,5),(-1,5), BLUE_LIGHT),
    ("ROWBACKGROUNDS",(0,1),(-1,4), [GREY_LIGHT, colors.white]),
    ("GRID",        (0,0),(-1,-1), 0.4, GREY_MID),
    ("FONTSIZE",    (0,0),(-1,-1), 9),
    ("ALIGN",       (1,0),(-1,-1), "RIGHT"),
    ("LEFTPADDING", (0,0),(-1,-1), 6),
    ("RIGHTPADDING",(0,0),(-1,-1), 6),
    ("TOPPADDING",  (0,0),(-1,-1), 4),
    ("BOTTOMPADDING",(0,0),(-1,-1), 4),
]))
story.append(t)
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "The baseline was saved in <i>CBP_P3B.mpp</i> prior to entering any actual data, "
    "providing the reference against which actual progress is measured in Part 2.",
    body))

story.append(hr())

# ── Part 2: Actual Progress ───────────────────────────────────────────────────
story.append(Paragraph("Part 2 – Actual Project Progress (Status Date: January 1, 2011)", h2))

# Q3 – Delay?
story.append(Paragraph("<b>Q3 – Will the Project Be Delayed?</b>", h3))
story.append(Paragraph(
    "As of January 1, 2011, <b>Task 13 – Serial I/O Drivers</b> (on the critical path) is "
    "running behind its planned schedule.  It was planned for 130 working days "
    "(Start: Nov 18, 2010; Finish: May 24, 2011), but by the status date only "
    "34 days of work have been completed with 119 days remaining, implying a total "
    "duration of <b>153 days</b> — 23 days longer than planned.  Because Task 13 "
    "started 15 days early (Nov 3 vs. Nov 18), the net delay is approximately "
    "<b>8 working days</b>, pushing its estimated finish to <b>~June 3, 2011</b>.",
    body))
story.append(Paragraph(
    "This delay propagates along the critical path as follows:",
    body))

delay_rows = [
    ["Task",                      "Orig. Finish",  "New Est. Finish", "Delay"],
    ["13 – Serial I/O Drivers",   "May 24, 2011",  "~Jun 3, 2011",   "+8 wd"],
    ["26 – Sys. H/W Test",        "Jun 29, 2011",  "~Jul 11, 2011",  "+8 wd"],
    ["16 – Network Interface",    "Nov 4, 2011",   "~Nov 18, 2011",  "+8 wd"],
    ["28 – Integration Accpt. Test","Feb 2, 2012", "~Feb 14, 2012",  "+8 wd"],
]
td = Table(delay_rows, colWidths=[7.5*cm, 3.5*cm, 4*cm, 2*cm], hAlign="LEFT")
td.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
    ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
    ("FONTNAME",    (0,4),(-1,4), "Helvetica-Bold"),
    ("BACKGROUND",  (0,4),(-1,4), colors.HexColor("#fff0e0")),
    ("ROWBACKGROUNDS",(0,1),(-1,3), [GREY_LIGHT, colors.white]),
    ("GRID",        (0,0),(-1,-1), 0.4, GREY_MID),
    ("FONTSIZE",    (0,0),(-1,-1), 9),
    ("ALIGN",       (1,0),(-1,-1), "CENTER"),
    ("LEFTPADDING", (0,0),(-1,-1), 6),
    ("RIGHTPADDING",(0,0),(-1,-1), 6),
    ("TOPPADDING",  (0,0),(-1,-1), 4),
    ("BOTTOMPADDING",(0,0),(-1,-1), 4),
]))
story.append(td)
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    "<b>New Estimated Completion Date: ~February 14, 2012</b> "
    "(approximately 2 weeks later than the Feb 2, 2012 target).  "
    "While this delay is modest, it is important to note that any further slippage "
    "in Task 13 or downstream tasks will push the project further past the deadline.",
    body))

# Q4 – Critical path
story.append(Paragraph("<b>Q4 – Remaining Tasks on the Critical Path</b>", h3))
story.append(Paragraph(
    "The remaining critical-path tasks as of January 1, 2011 are:",
    body))

cp_rows = [
    ["Task",                         "Duration", "Sched. Start",  "Sched. Finish", "Status"],
    ["13 – Serial I/O Drivers",      "130 d",    "Nov 18, 2010",  "May 24, 2011",  "In progress (34d done, 119 rem.)"],
    ["26 – Sys. Hard/Software Test", "25 d",     "May 25, 2011",  "Jun 29, 2011",  "Not started"],
    ["16 – Network Interface",       "90 d",     "Jun 30, 2011",  "Nov 4, 2011",   "Not started"],
    ["28 – Integration Accpt. Test", "60 d",     "Nov 7, 2011",   "Feb 2, 2012",   "Not started"],
]
tcp = Table(cp_rows, colWidths=[5*cm, 2*cm, 3*cm, 3*cm, 4*cm], hAlign="LEFT")
tcp.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
    ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
    ("ROWBACKGROUNDS",(0,1),(-1,-1), [GREY_LIGHT, colors.white]),
    ("GRID",        (0,0),(-1,-1), 0.4, GREY_MID),
    ("FONTSIZE",    (0,0),(-1,-1), 8),
    ("LEFTPADDING", (0,0),(-1,-1), 5),
    ("RIGHTPADDING",(0,0),(-1,-1), 5),
    ("TOPPADDING",  (0,0),(-1,-1), 3),
    ("BOTTOMPADDING",(0,0),(-1,-1), 3),
]))
story.append(tcp)
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    "All four remaining tasks are on the critical path.  Non-critical remaining tasks "
    "(Prototypes [44 days left], Order Circuit Boards, Assemble Models, Shell, "
    "Project Documentation) have float and will not directly delay project completion "
    "unless slippage exceeds their float.  "
    "<b>The implication is that there is essentially no schedule buffer remaining.</b>  "
    "Any further delay in Task 13 or in Network Interface will directly delay project "
    "delivery.  Tight monitoring and corrective action are critical at this stage.",
    body))

story.append(PageBreak())

# ── Q5 – EVM ─────────────────────────────────────────────────────────────────
story.append(Paragraph("Q5 – Earned Value Management (EVM) Analysis", h3))

evm_rows = [
    ["EVM Metric",                                   "Value",          "Interpretation"],
    ["BAC  – Budget at Completion",                  "$1,051,200",     "Total planned project budget"],
    ["PV   – Planned Value (BCWS)",                  "$690,160",       "Work scheduled by Jan 1, 2011"],
    ["EV   – Earned Value (BCWP)",                   "$691,634",       "Value of work actually performed"],
    ["AC   – Actual Cost (ACWP)",                    "$721,200",       "Cost incurred to date"],
    ["SV   – Schedule Variance (EV−PV)",             "+$1,474",        "Slightly AHEAD of schedule"],
    ["SPI  – Schedule Performance Index",            "1.002",          "On schedule (1.0 = on track)"],
    ["CV   – Cost Variance (EV−AC)",                 "−$29,566",       "OVER BUDGET by ~$30K"],
    ["CPI  – Cost Performance Index",                "0.959",          "$0.96 earned per $1 spent"],
    ["EAC  – Estimate at Completion (BAC/CPI)",      "$1,096,136",     "Projected total cost"],
    ["ETC  – Estimate to Complete (EAC−AC)",         "$374,936",       "Remaining cost to finish"],
    ["VAC  – Variance at Completion (BAC−EAC)",      "−$44,936",       "Expected cost OVERRUN"],
    ["TCPI – To-Complete Perf. Index",               "1.090",          "Must be 9% more efficient"],
]
te = Table(evm_rows, colWidths=[6.5*cm, 3.5*cm, 7*cm], hAlign="LEFT")
te.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
    ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
    ("ROWBACKGROUNDS",(0,1),(-1,-1), [GREY_LIGHT, colors.white]),
    ("GRID",        (0,0),(-1,-1), 0.4, GREY_MID),
    ("FONTSIZE",    (0,0),(-1,-1), 8),
    ("LEFTPADDING", (0,0),(-1,-1), 5),
    ("RIGHTPADDING",(0,0),(-1,-1), 5),
    ("TOPPADDING",  (0,0),(-1,-1), 3),
    ("BOTTOMPADDING",(0,0),(-1,-1), 3),
    # Colour CV/CPI/VAC rows RED
    ("TEXTCOLOR",   (1,8), (1,8),  RED),
    ("TEXTCOLOR",   (1,9), (1,9),  RED),
    ("TEXTCOLOR",   (1,12),(1,12), RED),
    ("FONTNAME",    (0,8), (-1,8), "Helvetica-Bold"),
    ("FONTNAME",    (0,9), (-1,9), "Helvetica-Bold"),
]))
story.append(te)
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    "<b>Analysis:</b>  The SPI of 1.002 indicates the project is essentially <b>on schedule</b> "
    "overall, with the early completion of several hardware and utility tasks offsetting the "
    "delay in Task 13.  However, the CPI of 0.959 reveals a significant <b>cost overrun</b>: "
    "the team is spending $1.04 for every $1 of planned value delivered.  Major contributors "
    "to the cost overrun are Disk Drivers (+$11,200) and Complex Utilities (+$8,400), both of "
    "which took substantially longer than planned.  The projected total cost of $1,096,136 "
    "exceeds the BAC by $44,936 (4.3% overrun).  The TCPI of 1.090 means the project team "
    "must improve cost efficiency by 9% over the remaining work to finish within budget — "
    "a challenging target given current trends.  See <u>Appendix B &amp; C</u> for charts.",
    body))

story.append(hr())

# Q6 – Recommendations
story.append(Paragraph("<b>Q6 – Recommendations</b>", h3))

recs = [
    ("<b>Accelerate Serial I/O Drivers (Task 13).</b>",
     "This is the single most critical risk to delivery.  Assign additional Development "
     "resources or approve overtime to reduce the 119 remaining days.  Even recovering "
     "half of the 8-working-day delay would restore the Feb 2 deadline."),

    ("<b>Investigate and Control Cost Overruns.</b>",
     "With CPI = 0.959, the team is consistently taking longer than planned.  Conduct a "
     "root-cause review of Disk Drivers and Complex Utilities overruns.  Implement tighter "
     "weekly time-tracking and re-estimate remaining work packages."),

    ("<b>Monitor Non-Critical Tasks with Float.</b>",
     "Tasks 6 (Prototypes, 44 days remaining) and 22 (Shell) have float but should be "
     "watched.  If they slip further, they could become critical and create additional risk."),

    ("<b>Prepare Contingency Plan for Network Interface (Task 16).</b>",
     "Task 16 (90 days, starting ~July 2011) is the longest remaining critical task.  "
     "Pre-position resources now to ensure a prompt start once Task 26 finishes."),

    ("<b>Reforecast and Communicate.</b>",
     "Update the project forecast to reflect EAC = $1,096,136 and new delivery date of "
     "~Feb 14, 2012.  Communicate proactively with stakeholders to manage expectations "
     "and avoid surprise at project close."),
]

for heading, detail in recs:
    story.append(Paragraph(f"{heading}  {detail}", body))

story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    "In summary, the Conveyor Belt Project is at a critical juncture.  "
    "Schedule-wise it is broadly on track (SPI ≈ 1.0), but cost performance is "
    "deteriorating (CPI = 0.96) and the critical path has no remaining float.  "
    "Immediate management attention to Task 13 and cost discipline across the "
    "remaining work packages is essential to deliver the project on time and within budget.",
    body))

story.append(PageBreak())

# ─── APPENDIX ────────────────────────────────────────────────────────────────
story.append(Paragraph("Appendix", h1))

# Appendix A – Cash Flow
story.append(Paragraph("Appendix A – Monthly Planned Cost Distribution", h2))
story.append(Paragraph(
    "The chart below shows the estimated monthly planned expenditure over the project "
    "life cycle.  Peak spending occurs in the first half of 2010 when hardware design, "
    "OS development, and utilities tasks run concurrently.",
    body))
story.append(cash_flow_bar())
story.append(Spacer(1, 0.3*cm))

# Sub-project cost pie (text table instead of actual pie)
story.append(Paragraph("Sub-Project Cost Breakdown:", bold_body))
subproj_rows = [
    ["Sub-Project",         "Planned Cost", "Tasks Included"],
    ["Hardware",            "$326,400",     "HW Spec, HW Design, HW Doc, Prototypes, Circuit Boards, Assembly"],
    ["Operating System",    "$333,600",     "Kernel Spec, Disk Drivers, Serial I/O, Memory Mgmt, OS Doc, Network IF"],
    ["Utilities",           "$182,000",     "Util Spec, Routine Util, Complex Util, Util Doc, Shell"],
    ["System Integration",  "$209,200",     "Arch Decisions, Integration Phase1, Sys H/W Test, Proj Doc, Accpt. Test"],
    ["TOTAL",               "$1,051,200",   ""],
]
tsp = Table(subproj_rows, colWidths=[4*cm, 3*cm, 10*cm], hAlign="LEFT")
tsp.setStyle(TableStyle([
    ("BACKGROUND",  (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",   (0,0),(-1,0), colors.white),
    ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
    ("FONTNAME",    (0,5),(-1,5), "Helvetica-Bold"),
    ("BACKGROUND",  (0,5),(-1,5), BLUE_LIGHT),
    ("ROWBACKGROUNDS",(0,1),(-1,4), [GREY_LIGHT, colors.white]),
    ("GRID",        (0,0),(-1,-1), 0.4, GREY_MID),
    ("FONTSIZE",    (0,0),(-1,-1), 8),
    ("LEFTPADDING", (0,0),(-1,-1), 5),
    ("RIGHTPADDING",(0,0),(-1,-1), 5),
    ("TOPPADDING",  (0,0),(-1,-1), 3),
    ("BOTTOMPADDING",(0,0),(-1,-1), 3),
]))
story.append(tsp)
story.append(Spacer(1, 0.4*cm))

# Appendix B – Planned vs Actual
story.append(Paragraph("Appendix B – Planned Cost vs. Actual Cost per Task", h2))
story.append(Paragraph(
    "Blue bars show the budgeted (planned) cost; orange bars show the actual cost incurred "
    "as of January 1, 2011.  Tasks where actual cost exceeds planned cost indicate "
    "cost overruns (Disk Drivers, Complex Utilities, Memory Management are notable).",
    body))
story.append(planned_vs_actual_bar_chart())
story.append(Spacer(1, 0.4*cm))

# Appendix C – EVM S-Curve
story.append(Paragraph("Appendix C – Earned Value S-Curve (PV / EV / AC)", h2))
story.append(Paragraph(
    "The S-curve plots cumulative PV (planned), EV (earned), and AC (actual) over time.  "
    "The convergence of PV and EV indicates the project is broadly on schedule, while "
    "the AC line running above EV confirms the cost overrun trend.",
    body))
story.append(evm_line_chart())
story.append(Spacer(1, 0.3*cm))

# Appendix D – Task-level EVM detail
story.append(Paragraph("Appendix D – Task-Level EVM Detail (as of Jan 1, 2011)", h2))

detail_rows = [
    ["Task",                     "BAC ($)","Plan\nDur","Act\nDur","Rem\nDur","% Done","EV ($)","AC ($)","CV ($)"],
    ["HW Specifications",        "68,000", "50","45","0", "100%","68,000","61,200", "+6,800"],
    ["HW Design",                "95,200", "70","67","0", "100%","95,200","91,120", "+4,080"],
    ["HW Documentation",         "38,400", "30","27","0", "100%","38,400","34,560", "+3,840"],
    ["Prototypes *",             "89,600", "80","34","44","43.6%","39,056","38,080",   "+976"],
    ["Kernel Specifications",    "16,000", "20","25","0", "100%","16,000","20,000", "−4,000"],
    ["Disk Drivers",            "112,000","100","110","0","100%","112,000","123,200","−11,200"],
    ["Serial I/O Drivers *",     "72,800","130","34","119","22.2%","16,178","19,040", "−2,862"],
    ["Memory Management",        "50,400", "90","98","0", "100%","50,400","54,880", "−4,480"],
    ["OS Documentation",         "32,000", "25","25","0", "100%","32,000","32,000",      "0"],
    ["Utilities Specifications", "20,400", "15","16","0", "100%","20,400","21,760", "−1,360"],
    ["Routine Utilities",        "57,600", "60","65","0", "100%","57,600","62,400", "−4,800"],
    ["Complex Utilities",        "44,800", "80","95","0", "100%","44,800","53,200", "−8,400"],
    ["Utilities Documentation",  "25,600", "20","22","0", "100%","25,600","28,160", "−2,560"],
    ["Architectural Decisions",  "20,000", "25","25","0", "100%","20,000","20,000",      "0"],
    ["Integration First Phase",  "56,000", "50","55","0", "100%","56,000","61,600", "−5,600"],
    ["TOTAL",                   "691,634","—",  "—",  "—","—",  "691,634","721,200","−29,566"],
]
# Column total: 3.8+1.7+1.0+1.0+1.0+1.2+2.1+2.1+2.1 = 16.0 cm (fits A4 with 2cm margins)
tdet = Table(detail_rows,
             colWidths=[3.8*cm, 1.7*cm, 1.0*cm, 1.0*cm, 1.0*cm, 1.2*cm, 2.1*cm, 2.1*cm, 2.1*cm],
             hAlign="LEFT")
tdet.setStyle(TableStyle([
    ("BACKGROUND",   (0,0),(-1,0), BLUE_DARK),
    ("TEXTCOLOR",    (0,0),(-1,0), colors.white),
    ("FONTNAME",     (0,0),(-1,0), "Helvetica-Bold"),
    ("FONTNAME",     (0,-1),(-1,-1), "Helvetica-Bold"),
    ("BACKGROUND",   (0,-1),(-1,-1), BLUE_LIGHT),
    ("ROWBACKGROUNDS",(0,1),(-1,-2), [GREY_LIGHT, colors.white]),
    ("GRID",         (0,0),(-1,-1), 0.3, GREY_MID),
    ("FONTSIZE",     (0,0),(-1,-1), 7),
    ("ALIGN",        (1,0),(-1,-1), "RIGHT"),
    ("LEFTPADDING",  (0,0),(-1,-1), 3),
    ("RIGHTPADDING", (0,0),(-1,-1), 3),
    ("TOPPADDING",   (0,0),(-1,-1), 2),
    ("BOTTOMPADDING",(0,0),(-1,-1), 2),
    ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
]))
story.append(tdet)
story.append(Spacer(1, 0.15*cm))
story.append(Paragraph("* In progress as of status date (Jan 1, 2011).  Dur = working days.", small))

# ─── Build PDF ────────────────────────────────────────────────────────────────
doc.build(story)
print(f"PDF generated: {OUTPUT}")
print(f"File size: {os.path.getsize(OUTPUT):,} bytes")
