"""
generate_documents.py
Generates official real document deliverables:
1. BankNifty_ShortStrangle_Backtest_Report.pdf (Professional PDF via ReportLab)
2. BankNifty_ShortStrangle_Backtest_Report.docx (Formatted Word Document via python-docx)
"""

import os
from pathlib import Path
import pandas as pd
from datetime import datetime

# ReportLab imports
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas

# python-docx imports
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = Path("d:/Qodo")
CHARTS_DIR = BASE_DIR / "charts"
PDF_PATH = BASE_DIR / "BankNifty_ShortStrangle_Backtest_Report.pdf"
DOCX_PATH = BASE_DIR / "BankNifty_ShortStrangle_Backtest_Report.docx"


# =============================================================================
# DATA DEFINITIONS
# =============================================================================

TIMING_DATA = [
    ["Stage #", "Pipeline Stage", "Time (s)", "% Total", "Technical Operations"],
    ["1", "Data Loading & Preprocessing", "44.94s", "85.7%", "Parses 10.26M rows; float32 dtypes; string slice strike extraction (str[9:-2]); drops 14.5k duplicates."],
    ["2", "Week-1 Day Selection", "0.27s", "0.5%", "Calendar-month grouping to identify 1st Wednesday; filters universe to 30 Week-1 trading days."],
    ["3", "Backtest Simulation Engine", "4.52s", "8.6%", "Vectorized strike selection at 09:20; High-based 50% SL scan using boolean masks & argmax."],
    ["4", "Statistical Computations", "0.06s", "0.1%", "CAGR, trade-wise peak-to-trough max drawdown, 9-segment win/loss & avg % P&L, monthly NAV."],
    ["5", "Chart Rendering", "1.54s", "2.9%", "Headless Matplotlib rendering of high-res dark-themed Equity & Drawdown curves."],
    ["6", "Excel Workbook Export", "1.09s", "2.1%", "Generates 3 styled worksheets (Guide, Statistics, Tradesheet) with embedded charts & formulas."],
    ["TOTAL", "Full Pipeline Execution", "52.41s", "100.0%", "Target: < 60.0s (PASSED with a 7.6s margin)."]
]

SUMMARY_STATS = [
    ["Metric", "Value", "Notes / Specification"],
    ["Strategy", "Bank Nifty 09:20 Short Strangle", "Sell 1 CE + 1 PE closest to Rs 50"],
    ["Trade Universe", "Week 1 of Each Calendar Month", "30 trading days (Mon-Wed of Week 1)"],
    ["Execution Time Window", "09:20:59 Entry -> 15:20:59 Exit", "Intraday (no overnight carry)"],
    ["Stop Loss Rule", "50% Hard Stop on Bar High", "Exit level = Entry * 1.5; checked 09:21-15:20"],
    ["Position Sizing", "1 Lot (15 units) Fixed", "No compounding, unscaled capital"],
    ["Base Starting Capital", "Rs 1,00,0000.00", "Rs 10 Lakhs base capital"],
    ["Ending Capital", "Rs 10,05,245.12", "Rs 10.05 Lakhs"],
    ["Total Net P&L", "+Rs 5,245.12", "Net profitable across 60 trades"],
    ["CAGR (%)", "0.57%", "Annualized NAV growth"],
    ["Max Drawdown (%)", "-0.19%", "Trade-wise running peak-to-trough drop"],
    ["Total Trades Executed", "60 trades", "30 trading days x 2 legs"],
    ["Win Rate", "58.33%", "35 Wins / 25 Losses"]
]

WIN_LOSS_DATA = [
    ["Segment", "Total Trades", "Wins", "Losses", "Win %", "Loss %"],
    ["Call Options (CE)", "30", "18", "12", "60.00%", "40.00%"],
    ["Put Options (PE)", "30", "17", "13", "56.67%", "43.33%"],
    ["Combined Strategy", "60", "35", "25", "58.33%", "41.67%"]
]

AVG_PNL_DATA = [
    ["Segment", "Trade Count", "Average % P&L"],
    ["CE -- Non-Expiry Days", "18", "+17.06%"],
    ["CE -- Expiry Days (Wed)", "12", "+19.82%"],
    ["CE -- Combined", "30", "+18.16%"],
    ["PE -- Non-Expiry Days", "18", "+18.15%"],
    ["PE -- Expiry Days (Wed)", "12", "+14.65%"],
    ["PE -- Combined", "30", "+16.75%"],
    ["All -- Non-Expiry Days", "36", "+17.60%"],
    ["All -- Expiry Days (Wed)", "24", "+17.23%"],
    ["All -- Combined", "60", "+17.45%"]
]

MONTHLY_DATA = [
    ["Month", "Start NAV", "End NAV", "Monthly % P&L", "Status"],
    ["2023-01", "100.0000", "100.0491", "+0.0491%", "Profit"],
    ["2023-02", "100.0491", "99.9740", "-0.0751%", "Loss"],
    ["2023-03", "99.9740", "99.9906", "+0.0166%", "Profit"],
    ["2023-04", "99.9906", "100.0851", "+0.0945%", "Profit"],
    ["2023-05", "100.0851", "100.1351", "+0.0499%", "Profit"],
    ["2023-06", "100.1351", "100.2786", "+0.1434%", "Profit"],
    ["2023-07", "100.2786", "100.1814", "-0.0970%", "Loss"],
    ["2023-08", "100.1814", "100.1992", "+0.0177%", "Profit"],
    ["2023-09", "100.1992", "100.2377", "+0.0384%", "Profit"],
    ["2023-10", "100.2377", "100.4596", "+0.2214%", "Profit"],
    ["2023-11", "100.4596", "100.4848", "+0.0251%", "Profit"],
    ["2023-12", "100.4848", "100.5245", "+0.0395%", "Profit"]
]

ASSUMPTIONS = [
    ("Assumption #1: Week-1 Definition", "Week 1 is defined as the trading days from the 1st of each calendar month up to and including the month's first Wednesday (expiry day). Thursdays onward (weeks 2-4) are excluded. This avoids the circular weekly no-op trap and restricts trading to ~1/4 of trading days."),
    ("Assumption #2: Strike Tie-Break", "If two option strikes are exactly equidistant from Rs 50, the strike with the lower premium (cheaper, further OTM) is selected. This is the more conservative risk choice for short positions."),
    ("Assumption #3: Stop-Loss Exit Fill", "When a 1-minute bar's High >= Stop Level (Entry * 1.5), the exit price is filled exactly at the Stop Level (the trigger price), reflecting an automated stop-market order fill rather than the bar close."),
    ("Assumption #4: Capital Base & Position Sizing", "Rs 10,00,000 base starting capital. Sizing is strictly fixed at 1 lot (15 units) without compounding or scaling. Available capital at any row = Starting Capital + Cumulative Realized P&L."),
    ("Assumption #5: 09:20 Bar Timestamp", "The options dataset timestamps denote bar completion at :59. The 09:20 entry corresponds to Time == '09:20:59'. Entry price is the close of this bar."),
    ("Assumption #6: Stop-Loss Monitoring Window", "Monitoring begins at 09:21:59 (the bar after entry) up to 15:20:59. The entry bar itself is excluded to eliminate lookahead bias."),
    ("Assumption #7: Underlying Spot Price Logging", "Spot close price is recorded at entry (09:20:00). When real spot data is present (Nov/Dec 2023), it is read directly; for dates prior to 2023-10-20, synthetic underlying spot is derived via options Put-Call Parity forward (ATM Strike + CE - PE).")
]


# =============================================================================
# 1. GENERATE PDF REPORT (ReportLab)
# =============================================================================

class NumberedCanvas(canvas.Canvas):
    """Adds running headers and footers with page numbers."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_header_footer(num_pages)
            super().showPage()
        super().save()

    def draw_header_footer(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#666666"))

        # Header (pages > 1)
        if self._pageNumber > 1:
            self.drawString(54, 750, "Bank Nifty 09:20 Short Strangle Backtest -- Quantitative Research Report")
            self.setStrokeColor(colors.HexColor("#CCCCCC"))
            self.setLineWidth(0.5)
            self.line(54, 744, 558, 744)

        # Footer
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.setLineWidth(0.5)
        self.line(54, 45, 558, 45)
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(558, 32, page_str)
        self.drawString(54, 32, "Confidential -- Quantitative Hiring Assignment Submission")
        self.restoreState()


def generate_pdf():
    doc = SimpleDocTemplate(
        str(PDF_PATH),
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=22,
        leading=26,
        textColor=colors.HexColor('#1F4E79'),
        spaceAfter=6
    )

    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor('#555555'),
        spaceAfter=14
    )

    h1_style = ParagraphStyle(
        'H1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#1F4E79'),
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    )

    h2_style = ParagraphStyle(
        'H2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=13,
        textColor=colors.HexColor('#2E75B6'),
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12.5,
        textColor=colors.HexColor('#222222'),
        spaceAfter=6
    )

    cell_style = ParagraphStyle(
        'Cell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#222222')
    )

    cell_bold = ParagraphStyle(
        'CellBold',
        parent=cell_style,
        fontName='Helvetica-Bold'
    )

    cell_hdr = ParagraphStyle(
        'CellHdr',
        parent=cell_style,
        fontName='Helvetica-Bold',
        textColor=colors.white
    )

    story = []

    # Title Block
    story.append(Paragraph("Bank Nifty 09:20 Short Strangle Backtest", title_style))
    story.append(Paragraph("<b>Quantitative Strategy Research Report</b> | Full 1-Year Backtest Deliverable", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1F4E79'), spaceAfter=14))

    # Executive Summary
    story.append(Paragraph("1. Executive Summary & Strategy Overview", h1_style))
    story.append(Paragraph(
        "This report documents the research methodology, execution speed benchmarks, and performance analysis of the "
        "<b>Bank Nifty 09:20 Short Strangle</b> systematic strategy. The backtest processes the complete 1-minute OHLC "
        "options dataset (<b>10,266,681 rows, 720 MB</b>) and spot dataset for the 2023 cycle under strict lookahead-bias avoidance.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Strategy Rules:</b> Every day in Week 1 of each calendar month, at 09:20:59, the algorithm sells one CE and one PE "
        "whose 1-minute close price is closest to Rs 50. Positions are managed with a <b>50% hard stop-loss</b> on bar High "
        "(fill price = Entry * 1.5) or closed at <b>15:20:59 close</b>. Sizing is fixed at 1 lot (15 units) without compounding.",
        body_style
    ))

    # Summary Statistics Table
    story.append(Spacer(1, 4))
    story.append(Paragraph("Key Strategy Results Summary", h2_style))
    t_data = []
    for r_idx, row in enumerate(SUMMARY_STATS):
        if r_idx == 0:
            t_data.append([Paragraph(c, cell_hdr) for c in row])
        else:
            t_data.append([Paragraph(row[0], cell_bold), Paragraph(row[1], cell_bold), Paragraph(row[2], cell_style)])
    
    table1 = Table(t_data, colWidths=[130, 110, 264])
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
    ]))
    story.append(table1)

    # Step-by-Step Runtime Breakdown
    story.append(Spacer(1, 10))
    story.append(Paragraph("2. Step-by-Step Runtime Breakdown (< 60 Seconds Benchmark)", h1_style))
    story.append(Paragraph(
        "The hiring assignment requires the backtest to run end-to-end (including raw data loading, deduplication, "
        "filtering, signal simulation, analytics, and Excel generation) in under 1 minute. "
        "The pipeline completed in <b>52.41 seconds</b>, surpassing this benchmark with a 7.6-second safety buffer.",
        body_style
    ))

    t_time = []
    for r_idx, row in enumerate(TIMING_DATA):
        if r_idx == 0:
            t_time.append([Paragraph(c, cell_hdr) for c in row])
        elif r_idx == len(TIMING_DATA) - 1:
            t_time.append([Paragraph(c, cell_bold) for c in row])
        else:
            t_time.append([Paragraph(row[0], cell_style), Paragraph(row[1], cell_bold),
                           Paragraph(row[2], cell_style), Paragraph(row[3], cell_style), Paragraph(row[4], cell_style)])

    table_time = Table(t_time, colWidths=[40, 115, 55, 45, 249])
    table_time.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E75B6')),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EBF3FF')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
    ]))
    story.append(table_time)

    # Performance Charts
    story.append(PageBreak())
    story.append(Paragraph("3. Visual Performance Curves", h1_style))
    story.append(Paragraph(
        "The equity curve is computed trade-wise (base NAV = 100), updating immediately upon realization of each individual leg. "
        "The drawdown curve tracks the percentage drop from running historical peaks.",
        body_style
    ))

    eq_img_path = CHARTS_DIR / "equity_curve.png"
    dd_img_path = CHARTS_DIR / "drawdown_curve.png"

    if eq_img_path.exists():
        story.append(RLImage(str(eq_img_path), width=500, height=210))
        story.append(Spacer(1, 8))

    if dd_img_path.exists():
        story.append(RLImage(str(dd_img_path), width=500, height=140))
        story.append(Spacer(1, 10))

    # Win/Loss and Segment Analytics
    story.append(Paragraph("4. Segment Performance & Win/Loss Breakdown", h1_style))
    
    col_w = [110, 75, 75, 75, 80, 89]
    t_wl = []
    for r_idx, row in enumerate(WIN_LOSS_DATA):
        if r_idx == 0:
            t_wl.append([Paragraph(c, cell_hdr) for c in row])
        else:
            t_wl.append([Paragraph(c, cell_bold if r_idx==3 else cell_style) for c in row])

    table_wl = Table(t_wl, colWidths=col_w)
    table_wl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#375623')),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EBF6EB')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
        ('TOPPADDING', (0, 0), (-1, -1), 3.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
    ]))
    story.append(table_wl)

    # Average % P&L Table
    story.append(Spacer(1, 8))
    story.append(Paragraph("Average % P&L per Trade (Expiry vs Non-Expiry Breakdown)", h2_style))
    t_avg = []
    for r_idx, row in enumerate(AVG_PNL_DATA):
        if r_idx == 0:
            t_avg.append([Paragraph(c, cell_hdr) for c in row])
        else:
            t_avg.append([Paragraph(row[0], cell_bold if "Combined" in row[0] else cell_style),
                          Paragraph(row[1], cell_style),
                          Paragraph(row[2], cell_bold)])

    table_avg = Table(t_avg, colWidths=[200, 100, 204])
    table_avg.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6B3FA0')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
    ]))
    story.append(table_avg)

    # Monthly Returns Table
    story.append(PageBreak())
    story.append(Paragraph("5. Monthly Performance & NAV Progression", h1_style))
    story.append(Paragraph(
        "Monthly returns are calculated from month-start NAV to month-end NAV. "
        "Ten out of twelve calendar months produced positive returns, yielding strong consistency.",
        body_style
    ))

    t_m = []
    for r_idx, row in enumerate(MONTHLY_DATA):
        if r_idx == 0:
            t_m.append([Paragraph(c, cell_hdr) for c in row])
        else:
            p_style = cell_bold if row[4] == "Profit" else ParagraphStyle('Loss', parent=cell_bold, textColor=colors.HexColor('#C00000'))
            t_m.append([Paragraph(row[0], cell_style), Paragraph(row[1], cell_style),
                        Paragraph(row[2], cell_style), Paragraph(row[3], p_style), Paragraph(row[4], p_style)])

    table_m = Table(t_m, colWidths=[90, 100, 100, 110, 104])
    table_m.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#9E3B00')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#FFFFFF'), colors.HexColor('#F8FAFC')]),
    ]))
    story.append(table_m)

    # Methodological Assumptions
    story.append(Spacer(1, 10))
    story.append(Paragraph("6. Methodological Assumptions & Research Design", h1_style))
    story.append(Paragraph(
        "Per the hiring assignment brief, all critical design ambiguities are explicitly stated and justified below:",
        body_style
    ))

    for title, desc in ASSUMPTIONS:
        story.append(Paragraph(f"<b>{title}:</b> {desc}", body_style))
        story.append(Spacer(1, 2))

    # Deliverables & Excel Structure
    story.append(Spacer(1, 8))
    story.append(Paragraph("7. Structure of the Excel Deliverable", h1_style))
    story.append(Paragraph(
        "The Excel workbook <b>BankNifty_ShortStrangle_Backtest.xlsx</b> contains three structured worksheets:",
        body_style
    ))
    story.append(Paragraph("• <b>Worksheet 1 (Guide):</b> Complete strategy rules, runtime benchmarks, output definitions, and assumption documentation.", body_style))
    story.append(Paragraph("• <b>Worksheet 2 (Statistics):</b> Key performance indicators, win/loss tables, average % P&L matrix, monthly returns, trade-wise NAV index series, and embedded graphical curves.", body_style))
    story.append(Paragraph("• <b>Worksheet 3 (Tradesheet):</b> Granular trade log of all 60 trades featuring entry/exit timestamps, option tickers, strikes, execution prices, quantities, gross and cumulative P&L, available capital, and underlying spot close.", body_style))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Generated PDF: {PDF_PATH}")


# =============================================================================
# 2. GENERATE WORD DOCUMENT (.DOCX)
# =============================================================================

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)


def generate_docx():
    doc = Document()

    # Page Margins
    for sec in doc.sections:
        sec.top_margin = Inches(0.75)
        sec.bottom_margin = Inches(0.75)
        sec.left_margin = Inches(0.75)
        sec.right_margin = Inches(0.75)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("Bank Nifty 09:20 Short Strangle Backtest")
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Quantitative Strategy Research Report | Hiring Assignment Submission")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # Section 1: Executive Summary
    h1 = doc.add_heading("1. Executive Summary & Strategy Overview", level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    h1.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph(
        "This report documents the research methodology, execution speed benchmarks, and performance analysis of the "
        "Bank Nifty 09:20 Short Strangle systematic strategy. The backtest processes the complete 1-minute OHLC "
        "options dataset (10,266,681 rows, 720 MB) and spot dataset for the 2023 cycle under strict lookahead-bias avoidance."
    )
    doc.add_paragraph(
        "Strategy Rules: Every day in Week 1 of each calendar month, at 09:20:59, the algorithm sells one CE and one PE "
        "whose 1-minute close price is closest to Rs 50. Positions are managed with a 50% hard stop-loss on bar High "
        "(fill price = Entry * 1.5) or closed at 15:20:59 close. Sizing is fixed at 1 lot (15 units) without compounding."
    )

    # Summary Table
    t1 = doc.add_table(rows=len(SUMMARY_STATS), cols=3)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(SUMMARY_STATS):
        for c_idx, val in enumerate(row):
            cell = t1.cell(r_idx, c_idx)
            cell.text = val
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            if r_idx == 0:
                set_cell_background(cell, "1F4E79")
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif c_idx == 0 or c_idx == 1:
                cell.paragraphs[0].runs[0].font.bold = True
                if r_idx % 2 == 1:
                    set_cell_background(cell, "F2F2F2")
            else:
                if r_idx % 2 == 1:
                    set_cell_background(cell, "F2F2F2")

    doc.add_paragraph()

    # Section 2: Runtime Benchmark
    h2 = doc.add_heading("2. Step-by-Step Runtime Breakdown (< 60s Target)", level=1)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    h2.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph(
        "The assignment requires the entire pipeline to run start-to-end in under 1 minute. "
        "The code executed in 52.41 seconds on the complete ~10.26M row dataset, beating the target with a 7.6-second margin."
    )

    t2 = doc.add_table(rows=len(TIMING_DATA), cols=5)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(TIMING_DATA):
        for c_idx, val in enumerate(row):
            cell = t2.cell(r_idx, c_idx)
            cell.text = val
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if r_idx == 0:
                set_cell_background(cell, "2E75B6")
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif r_idx == len(TIMING_DATA) - 1:
                set_cell_background(cell, "EBF3FF")
                cell.paragraphs[0].runs[0].font.bold = True
            else:
                if r_idx % 2 == 1:
                    set_cell_background(cell, "F9FBFD")

    doc.add_page_break()

    # Section 3: Performance Charts
    h3 = doc.add_heading("3. Visual Performance Curves", level=1)
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)
    h3.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph(
        "The equity curve is computed trade-wise (base NAV = 100), updating immediately upon realization of each individual leg. "
        "The drawdown curve tracks peak-to-trough drop across the full 60 trades."
    )

    eq_img_path = CHARTS_DIR / "equity_curve.png"
    dd_img_path = CHARTS_DIR / "drawdown_curve.png"

    if eq_img_path.exists():
        doc.add_picture(str(eq_img_path), width=Inches(6.5))
        doc.add_paragraph()

    if dd_img_path.exists():
        doc.add_picture(str(dd_img_path), width=Inches(6.5))
        doc.add_paragraph()

    # Section 4: Win/Loss Breakdown
    h4 = doc.add_heading("4. Win/Loss and Segment Performance", level=1)
    h4.paragraph_format.space_before = Pt(12)
    h4.paragraph_format.space_after = Pt(4)
    h4.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    t3 = doc.add_table(rows=len(WIN_LOSS_DATA), cols=6)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(WIN_LOSS_DATA):
        for c_idx, val in enumerate(row):
            cell = t3.cell(r_idx, c_idx)
            cell.text = val
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            if r_idx == 0:
                set_cell_background(cell, "375623")
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif r_idx == len(WIN_LOSS_DATA) - 1:
                set_cell_background(cell, "EBF6EB")
                cell.paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # Section 5: Monthly Table
    h5 = doc.add_heading("5. Monthly Performance Table (NAV-Indexed)", level=1)
    h5.paragraph_format.space_before = Pt(12)
    h5.paragraph_format.space_after = Pt(4)
    h5.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    t4 = doc.add_table(rows=len(MONTHLY_DATA), cols=5)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(MONTHLY_DATA):
        for c_idx, val in enumerate(row):
            cell = t4.cell(r_idx, c_idx)
            cell.text = val
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            if r_idx == 0:
                set_cell_background(cell, "9E3B00")
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif r_idx % 2 == 1:
                set_cell_background(cell, "FDF8F5")

    doc.add_page_break()

    # Section 6: Assumptions
    h6 = doc.add_heading("6. Methodological Assumptions & Research Design", level=1)
    h6.paragraph_format.space_before = Pt(12)
    h6.paragraph_format.space_after = Pt(4)
    h6.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    for title, desc in ASSUMPTIONS:
        p = doc.add_paragraph()
        run_t = p.add_run(title + ": ")
        run_t.bold = True
        run_t.font.color.rgb = RGBColor(0x37, 0x56, 0x23)
        p.add_run(desc)

    # Section 7: Deliverables
    h7 = doc.add_heading("7. Submission Deliverables & Excel Layout", level=1)
    h7.paragraph_format.space_before = Pt(12)
    h7.paragraph_format.space_after = Pt(4)
    h7.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph(
        "The Excel workbook BankNifty_ShortStrangle_Backtest.xlsx is structured into three sheets:"
    )
    doc.add_paragraph("• Sheet 1: Guide -- Complete documentation, timing benchmarks, output guide, and assumptions.")
    doc.add_paragraph("• Sheet 2: Statistics -- High-level performance metrics, segment win/loss, monthly NAV returns, and embedded charts.")
    doc.add_paragraph("• Sheet 3: Tradesheet -- Granular 60-trade log with all 12 requested fields and spot prices at entry.")

    doc.save(str(DOCX_PATH))
    print(f"Generated DOCX: {DOCX_PATH}")


if __name__ == "__main__":
    generate_pdf()
    generate_docx()
