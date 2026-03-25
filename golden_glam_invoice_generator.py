from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from pathlib import Path

styles = getSampleStyleSheet()


def draw_invoice(inv, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []

    # ── HEADER ──
    elements.append(Paragraph(f"<b>Golden Glam Interiors</b>", styles["Title"]))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph(f"<b>Invoice:</b> {inv.get('number')}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Date:</b> {inv.get('date')}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Client:</b> {inv.get('client_name')}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # ── ITEMS TABLE ──
    data = [["Item", "Description", "Qty", "Unit Price", "Total"]]

    subtotal = 0

    for item in inv["items"]:
        total = item["qty"] * item["unit_price"] * (1 - item.get("discount", 0))
        subtotal += total

        data.append([
            item.get("no"),
            item.get("description"),
            item.get("qty"),
            f"${item.get('unit_price'):.2f}",
            f"${total:.2f}",
        ])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 20))

    # ── TOTALS ──
    delivery = inv.get("delivery_charge", 0)
    tax = subtotal * inv.get("tax_rate", 0)
    total = subtotal + delivery + tax

    elements.append(Paragraph(f"Subtotal: ${subtotal:.2f}", styles["Normal"]))
    elements.append(Paragraph(f"Delivery: ${delivery:.2f}", styles["Normal"]))
    elements.append(Paragraph(f"Tax: ${tax:.2f}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Total: ${total:.2f}</b>", styles["Normal"]))

    doc.build(elements)

    # ── EXCEL ──
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws.append(["Item", "Description", "Qty", "Unit Price", "Total"])

    for item in inv["items"]:
        total = item["qty"] * item["unit_price"]
        ws.append([
            item.get("no"),
            item.get("description"),
            item.get("qty"),
            item.get("unit_price"),
            total,
        ])

    xlsx_path = Path(output_path).with_name(Path(output_path).stem + "_INTERNAL.xlsx")
    wb.save(xlsx_path)
