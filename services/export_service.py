"""Export měsíčních reportů do CSV a PDF."""

import csv
import io
from datetime import datetime


def build_csv(report_data, month, year, project_name=""):
    """Vytvoří CSV kompatibilní s Excel Windows."""
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";", lineterminator="\n")
    writer.writerow(["Projekt", project_name])
    writer.writerow(["Období", f"{month:02d}/{year}"])
    writer.writerow([])
    writer.writerow(["Zaměstnanec", "Celkem hodin", "Volné dny"])
    total_hours = 0.0
    total_free_days = 0
    for employee, data in sorted(report_data.items()):
        hours = float(data.get("total_hours", 0))
        free_days = int(data.get("free_days", 0))
        writer.writerow([employee, f"{hours:.2f}".replace(".", ","), free_days])
        total_hours += hours
        total_free_days += free_days
    writer.writerow([])
    writer.writerow(["Celkem", f"{total_hours:.2f}".replace(".", ","), total_free_days])
    return output.getvalue().encode("utf-8-sig")


def build_pdf(report_data, month, year, project_name=""):
    """Vytvoří jednoduchý PDF report; reportlab je jediná exportní závislost."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle, Paragraph

    output = io.BytesIO()
    document = SimpleDocTemplate(output, pagesize=A4, rightMargin=18 * mm, leftMargin=18 * mm)
    styles = getSampleStyleSheet()
    font_name = "Helvetica"
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    try:
        pdfmetrics.registerFont(TTFont("HodinyUnicode", font_path))
        font_name = "HodinyUnicode"
    except Exception:
        pass
    styles["Title"].fontName = font_name
    styles["Normal"].fontName = font_name
    elements = [
        Paragraph("Evidence pracovní doby", styles["Title"]),
        Paragraph(f"Projekt: {project_name or 'Nepojmenovaný projekt'}", styles["Normal"]),
        Paragraph(f"Období: {month:02d}/{year} | Vytvořeno: {datetime.now():%d.%m.%Y %H:%M}", styles["Normal"]),
        Spacer(1, 8 * mm),
    ]
    rows = [["Zaměstnanec", "Celkem hodin", "Volné dny"]]
    total_hours = 0.0
    total_free_days = 0
    for employee, data in sorted(report_data.items()):
        hours = float(data.get("total_hours", 0))
        free_days = int(data.get("free_days", 0))
        rows.append([employee, f"{hours:.2f} h", str(free_days)])
        total_hours += hours
        total_free_days += free_days
    rows.append(["Celkem", f"{total_hours:.2f} h", str(total_free_days)])
    table = Table(rows, colWidths=[95 * mm, 45 * mm, 35 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#004EA3")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#eaf2f8")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(table)
    document.build(elements)
    return output.getvalue()
