import os
import re
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from openpyxl import Workbook

def generate_from_summary(summary_path, output_dir):
    with open(summary_path, "r", encoding="utf-8") as f:
        txt = f.read()

    header = re.search(r"GG-INV\|([^|]+)\|([^|]+)\|ref:(.*)", txt)
    if not header:
        raise ValueError("Could not parse invoice header line.")
    number = header.group(1).strip()
    date = header.group(2).strip()
    reference = header.group(3).strip()

    client = re.search(r"client:([^|]+)\|no:([^|]+)\|ph:(.+)", txt)
    if not client:
        raise ValueError("Could not parse client line.")
    client_name = client.group(1).strip()
    client_no = client.group(2).strip()
    client_phone = client.group(3).strip()

    addr_match = re.search(r"addr:(.+)", txt)
    addr = addr_match.group(1).strip() if addr_match else ""

    totals = re.search(
        r"([^\n|]+ Delivery(?: \([^)]+\))?)\|sub:\$([0-9.,]+)\|del:\$([0-9.,]+)\|tax\((\d+)%\):\$([0-9.,]+)\|total:\$([0-9.,]+)",
        txt
    )
    if not totals:
        raise ValueError(
            "Could not parse delivery/totals line from summary. "
            "Expected something like 'White Glove Delivery|sub:$500.00|del:$0.00|tax(8%):$40.00|total:$540.00'"
        )

    delivery_type = totals.group(1).strip()
    subtotal = float(totals.group(2).replace(",", ""))
    delivery = float(totals.group(3).replace(",", ""))
    tax_rate = int(totals.group(4))
    tax = float(totals.group(5).replace(",", ""))
    total = float(totals.group(6).replace(",", ""))

    item_pattern = re.compile(
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$([0-9.,]+)\|tot:\$([0-9.,]+)\|del:(.*?)(?=\n\[|\n[^\n|]+ Delivery|\npay:|\nINTERNAL|\Z)",
        re.S
    )

    items = []
    for m in item_pattern.finditer(txt):
        item_no, desc, unit, qty, price, tot, delivery_txt = m.groups()
        items.append({
            "item_no": item_no.strip(),
            "desc": desc.strip(),
            "unit": unit.strip(),
            "qty": int(qty),
            "price": float(price.replace(",", "")),
            "total": float(tot.replace(",", "")),
            "delivery_txt": delivery_txt.strip()
        })

    if not items:
        raise ValueError("Could not parse any invoice items.")

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
        ["", "", f"Tax ({tax_rate}%)", int(tax)],
        ["", "", "TOTAL", int(total)],
    ]

    table = Table(table_data)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("BACKGROUND", (-2, -1), (-1, -1), colors.lightgrey),
    ]))

    doc.build([table])

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Internal"

    ws.append(["Item No", "Item", "Unit", "Qty", "Price", "Total", "Delivery"])
    for i in items:
        ws.append([
            i["item_no"],
            i["desc"],
            i["unit"],
            i["qty"],
            int(i["price"]),
            int(i["total"]),
            i["delivery_txt"]
        ])

    ws.append(["", "", "", "", "Subtotal", int(subtotal), ""])
    ws.append(["", "", "", "", "Delivery", int(delivery), delivery_type])
    ws.append(["", "", "", "", f"Tax ({tax_rate}%)", int(tax), ""])
    ws.append(["", "", "", "", "TOTAL", int(total), ""])

    excel_path = os.path.join(output_dir, f"GG__{number}_INTERNAL.xlsx")
    wb.save(excel_path)

    return pdf_path, excel_path
