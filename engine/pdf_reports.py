from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


class NumberedCanvas(Canvas):
    def __init__(self, *args, **kwargs):
        Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(page_count)
            Canvas.showPage(self)
        Canvas.save(self)

    def _draw_page_number(self, page_count: int):
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.grey)
        page_width = float(self._pagesize[0]) if self._pagesize else 11.0 * inch
        self.drawCentredString(page_width / 2.0, 0.35 * inch, f"-- {self._pageNumber} of {page_count} --")


def _kind_from_title(title: str) -> str:
    title_cf = str(title or "").casefold()
    if "staff" in title_cf:
        return "staff"
    if "board" in title_cf:
        return "board"
    if "vendor" in title_cf:
        return "vendors"
    return "generic"


def _format_exclusion_date(status, date):
    if status in {"CONFIRMED", "POSSIBLE"} and date:
        return str(date)
    return ""


def _split_name(value: str):
    name = str(value or "").strip()
    if not name:
        return "", ""
    if "," in name:
        left, right = name.split(",", 1)
        return left.strip(), right.strip()
    parts = name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[-1], " ".join(parts[:-1])


def _format_display_dob(value):
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.strptime(text, "%Y-%m-%d")
        return f"{parsed.month}/{parsed.day}/{parsed.year}"
    except ValueError:
        return text


def _build_table_data(kind: str, rows: list[dict], body_style: ParagraphStyle):
    if kind == "staff":
        headers = [
            "Last Name",
            "First Name",
            "Job Title",
            "Employment Status",
        ]
        col_widths = [2.0 * inch, 1.9 * inch, 4.9 * inch, 1.7 * inch]
        data = [headers]
        for row in rows:
            last_name = str(row.get("last_name_display") or "").strip()
            first_name = str(row.get("first_name_display") or "").strip()
            if not last_name and not first_name:
                name_text = row.get("name_display") or row.get("name", "")
                last_name, first_name = _split_name(name_text)
            data.append(
                [
                    Paragraph(last_name, body_style),
                    Paragraph(first_name, body_style),
                    Paragraph(str(row.get("role", "")), body_style),
                    Paragraph(str(row.get("status", "")), body_style),
                ]
            )
        return data, col_widths

    if kind == "board":
        headers = ["Name", "DOB", "SAM.gov Exclusion Date", "HHS/OIG Exclusion Date"]
        col_widths = [4.2 * inch, 1.3 * inch, 2.2 * inch, 2.3 * inch]
        data = [headers]
        for row in rows:
            data.append(
                [
                    Paragraph(str(row.get("name_display") or row.get("name", "")), body_style),
                    Paragraph(_format_display_dob(row.get("dob", "")), body_style),
                    Paragraph(_format_exclusion_date(row.get("sam_status"), row.get("sam_date")), body_style),
                    Paragraph(_format_exclusion_date(row.get("oig_status"), row.get("oig_date")), body_style),
                ]
            )
        return data, col_widths

    if kind == "vendors":
        headers = ["Name", "City", "State", "SAM.gov Exclusion Date", "HHS/OIG Exclusion Date"]
        col_widths = [5.5 * inch, 1.5 * inch, 0.8 * inch, 1.1 * inch, 1.1 * inch]
        data = [headers]
        for row in rows:
            data.append(
                [
                    Paragraph(str(row.get("name_display") or row.get("name", "")), body_style),
                    Paragraph(str(row.get("city", "")), body_style),
                    Paragraph(str(row.get("state", "")), body_style),
                    Paragraph(_format_exclusion_date(row.get("sam_status"), row.get("sam_date")), body_style),
                    Paragraph(_format_exclusion_date(row.get("oig_status"), row.get("oig_date")), body_style),
                ]
            )
        return data, col_widths

    headers = ["Name", "Role", "SAM Status", "OIG Status"]
    col_widths = [2.8 * inch, 2.0 * inch, 1.35 * inch, 1.35 * inch]
    data = [headers]
    for row in rows:
        data.append(
            [
                Paragraph(str(row.get("name_display") or row.get("name", "")), body_style),
                Paragraph(str(row.get("role", "")), body_style),
                Paragraph(str(row.get("sam_status", "")), body_style),
                Paragraph(str(row.get("oig_status", "")), body_style),
            ]
        )
    return data, col_widths


def _intro_text(kind: str) -> tuple[str, str]:
    if kind == "staff":
        return (
            "NO EXCLUSIONS FOUND",
            "All staff members have been screened against the OIG and SAM exclusion databases. "
            "No matches were found. All staff are clear to continue employment.",
        )
    if kind == "board":
        return (
            "NO EXCLUSIONS FOUND",
            "All board members have been screened against the OIG and SAM exclusion databases. "
            "No matches were found.",
        )
    if kind == "vendors":
        return (
            "NO EXCLUSIONS FOUND",
            "All vendors have been screened against the OIG and SAM exclusion databases. "
            "No matches were found.",
        )
    return (
        "NO EXCLUSIONS FOUND",
        "All records have been screened against the OIG and SAM exclusion databases.",
    )


def generate_pdf_report(output_path, client_name, month, title, rows):
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(letter),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
    )
    styles = getSampleStyleSheet()
    elements = []

    heading_client = ParagraphStyle(
        "heading_client",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=16,
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    heading_title = ParagraphStyle(
        "heading_title",
        parent=styles["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=14,
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    heading_month = ParagraphStyle(
        "heading_month",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=12,
        alignment=TA_CENTER,
        spaceAfter=12,
    )
    summary_heading = ParagraphStyle(
        "summary_heading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        alignment=TA_LEFT,
        spaceAfter=3,
    )
    summary_body = ParagraphStyle(
        "summary_body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
        alignment=TA_LEFT,
        spaceAfter=2,
    )
    notice_style = ParagraphStyle(
        "notice_style",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#2E7D32"),
        alignment=TA_LEFT,
        spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "body_style",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=TA_LEFT,
    )
    footer_style = ParagraphStyle(
        "footer_style",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.grey,
        alignment=TA_LEFT,
    )

    kind = _kind_from_title(title)
    total = len(rows or [])
    oig_found = sum(1 for row in rows if row.get("oig_status") == "CONFIRMED")
    sam_found = sum(1 for row in rows if row.get("sam_status") == "CONFIRMED")
    total_matches = sum(
        1
        for row in rows
        if row.get("oig_status") in {"CONFIRMED", "POSSIBLE"} or row.get("sam_status") in {"CONFIRMED", "POSSIBLE"}
    )

    title_display = str(title or "").replace(" Exclusion Report", " Exclusion Screening Report")
    elements.append(Paragraph(str(client_name or ""), heading_client))
    elements.append(Paragraph(title_display, heading_title))
    elements.append(Paragraph(str(month or ""), heading_month))

    section_label = {
        "staff": "Staff",
        "board": "Board Members",
        "vendors": "Vendors",
    }.get(kind, "Records")

    elements.append(Paragraph("Screening Summary", summary_heading))
    elements.append(Paragraph(f"Total {section_label} Screened: {total}", summary_body))
    elements.append(Paragraph(f"OIG Exclusions Found: {oig_found}", summary_body))
    elements.append(Paragraph(f"SAM Exclusions Found: {sam_found}", summary_body))
    elements.append(Paragraph(f"Total Matches: {total_matches}", summary_body))
    elements.append(Paragraph(f"Report Date: {datetime.now().strftime('%B %d, %Y')}", summary_body))
    elements.append(Spacer(1, 0.1 * inch))

    notice_title, notice_text = _intro_text(kind)
    elements.append(Paragraph(notice_title, notice_style))
    elements.append(Paragraph(notice_text, summary_body))
    elements.append(Spacer(1, 0.1 * inch))

    table_title = {
        "staff": "Screened Staff List",
        "board": "Screened Board Members",
        "vendors": "Screened Vendors",
    }.get(kind, "Screened Records")
    elements.append(Paragraph(table_title, summary_heading))

    data, col_widths = _build_table_data(kind, rows or [], body_style)
    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#224B7A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B0BEC5")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elements.append(table)
    elements.append(Spacer(1, 0.16 * inch))
    elements.append(
        Paragraph(
            "This report was generated using automated screening against the U.S. Department of Health and Human "
            "Services Office of Inspector General (OIG) List of Excluded Individuals and Entities and the System "
            "for Award Management (SAM) Exclusions database.",
            footer_style,
        )
    )

    doc.build(elements, canvasmaker=NumberedCanvas)