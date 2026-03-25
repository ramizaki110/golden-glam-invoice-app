import os
import re
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from openpyxl import Workbook

def generate_from_summary(summary_path, output_dir):
    with open(summary_path, "r", encoding="utf-8") as f:
        txt = f.read()

    # Invoice number
    number, date = re.search(r"GG-INV\|([^|]+)\|([^|]+)", txt).groups()

    # Totals
    totals = re.search(
        r"(Standard Delivery)\|sub:\$([0-9.,]+)\|del:\$([0-9.,]+)\|tax\(8%\):\$([0-9.,]+)\|total:\$([0-9.,]+)",
        txt
    )

    subtotal = float(totals.group(2).replace(",", ""))
    delivery = float(totals.group(3).replace(",", ""))
    tax = float(totals.group(4).replace(",", ""))
    total = float(totals.group(5).replace(",", ""))

    # Items
    item_pattern = re.compile(
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$([0-9.,]+)\|tot:\$([0-9.,]+)",
        re.S
    )

    items = []
    for m in item_pattern.finditer(txt):
        no, desc, unit, qty, price, tot = m.groups()
        items.append({
            "desc": desc.strip(),
            "qty": int(qty),
            "price": float(price.replace(",", "")),
            "total": float(tot.replace(",", ""))
        })

    # PDF
    pdf_path = os.path.join(output_dir, f"GG__{number}.pdf")
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)

    table_data = [["Item", "Qty", "Price", "Total"]]

    for i in items:
        table_data.append([
            i["desc"],
            i["qty"],
            int(i["price"]),
            int(i["total"])
        ])

    table_data += [
        ["", "", "Subtotal", int(subtotal)],
        ["", "", "Delivery", int(delivery)],
        ["", "", "Tax", int(tax)],
        ["", "", "TOTAL", int(total)],
    ]

    table = Table(table_data)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black)
    ]))

    doc.build([table])

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.append(["Item", "Qty", "Price", "Total"])

    for i in items:
        ws.append([
            i["desc"],
            i["qty"],
            int(i["price"]),
            int(i["total"])
        ])

    ws.append(["", "", "Subtotal", int(subtotal)])
    ws.append(["", "", "Delivery", int(delivery)])
    ws.append(["", "", "Tax", int(tax)])
    ws.append(["", "", "TOTAL", int(total)])

    excel_path = os.path.join(output_dir, f"GG__{number}_INTERNAL.xlsx")
    wb.save(excel_path)

    return pdf_path, excel_path