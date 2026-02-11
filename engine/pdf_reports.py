from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from datetime import datetime


def generate_pdf_report(
    output_path,
    client_name,
    month,
    title,
    rows
):
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=letter
    )

    elements = []
    styles = getSampleStyleSheet()

    header_text = f"{client_name}\n{title}\nMonth: {month}\nGenerated: {datetime.utcnow().isoformat()}"
    elements.append(Paragraph(header_text, styles["Heading2"]))
    elements.append(Spacer(1, 0.5 * inch))

    if not rows:
        elements.append(Paragraph("No records.", styles["Normal"]))
        doc.build(elements)
        return

    # Determine columns dynamically
    headers = ["Name", "Role", "Vendor Type", "OIG Status", "SAM Status"]

    data = [headers]

    for row in rows:
        name = row.get("name", "")
        role = row.get("role", "")
        vendor_type = row.get("classification", "")

        oig_status = format_status(
            row.get("oig_status"),
            row.get("oig_date")
        )

        sam_status = format_status(
            row.get("sam_status"),
            row.get("sam_date")
        )

        data.append([
            name,
            role,
            vendor_type,
            oig_status,
            sam_status
        ])

    table = Table(data, repeatRows=1)

    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (3, 1), (-1, -1), "LEFT"),
    ])

    table.setStyle(style)
    elements.append(table)

    doc.build(elements)


def format_status(status, date):
    if status == "CONFIRMED":
        return f"Confirmed Match – {date}"
    elif status == "POSSIBLE":
        return f"Possible Match – Review Required"
    else:
        return "Not Found"