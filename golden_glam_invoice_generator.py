import os
import base64
import tempfile
from pathlib import Path
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "golden_glam_logo_final.png"

styles = getSampleStyleSheet()

BLACK = colors.HexColor("#231f1e")
DARK = colors.HexColor("#4a4745")
MID = colors.HexColor("#8c8a87")
BORDER = colors.HexColor("#d8d5d2")


def usd(v: float) -> str:
    return f"${v:,.0f}"


def fmt_date_for_footer(date_str: str) -> str:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return date_str or ""


def _decode_image(image_value: str) -> str | None:
    if not image_value:
        return None

    if not image_value.startswith("data:image"):
        return image_value if os.path.exists(image_value) else None

    try:
        header, b64 = image_value.split(",", 1)
        ext = "png" if "png" in header.lower() else "jpg"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
        tmp.write(base64.b64decode(b64))
        tmp.close()
        return tmp.name
    except Exception:
        return None


def _footer(canvas, doc):
    inv = getattr(doc, "_gg_invoice", {})
    print_date = fmt_date_for_footer(inv.get("date", ""))
    page_num = canvas.getPageNumber()

    canvas.saveState()
    y_line = 0.70 * inch

    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.6)
    canvas.line(doc.leftMargin, y_line, letter[0] - doc.rightMargin, y_line)

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(MID)
    canvas.drawString(doc.leftMargin, y_line - 0.16 * inch, f"Print date: {print_date}")
    canvas.drawRightString(letter[0] - doc.rightMargin, y_line - 0.16 * inch, f"Page {page_num} of 1")

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(BLACK)
    canvas.drawCentredString(letter[0] / 2, 0.46 * inch, "GOLDEN GLAM INTERIORS LLC")

    canvas.setFont("Helvetica", 6)
    canvas.setFillColor(DARK)
    canvas.drawCentredString(
        letter[0] / 2,
        0.32 * inch,
        "Address: 828 Highland Ln Ne, Apt. 2204, Atlanta, GA 30306  |  Phone: 770-375-7343",
    )
    canvas.drawCentredString(
        letter[0] / 2,
        0.20 * inch,
        "Bank #: 930283558  |  Routing: 061092387  |  Zelle: rana_salah@goldenglam.nl  |  E-mail: sales@goldenglam.nl",
    )
    canvas.restoreState()


def _autosize(ws, widths=None):
    if widths:
        for col_idx, width in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        return

    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 2, 40)


def _write_internal_excel(inv: dict, output_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Internal"

    thin = Side(style="thin", color="D9D9D9")
    hdr_fill = PatternFill("solid", fgColor="231F1E")
    hdr_font = Font(color="FFFFFF", bold=True)
    section_fill = PatternFill("solid", fgColor="F5F2EF")

    headers = [
        "Item No",
        "Vendor No",
        "Description",
        "Qty",
        "Unit",
        "Unit Price",
        "Discount %",
        "Line Total",
        "Raw Cost",
        "Cost Disc %",
        "Unit Cost",
        "Extended Cost",
        "Profit",
        "GM %",
        "Delivery",
        "Image Included",
    ]
    ws.append(headers)

    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    subtotal = 0

    for item in inv["items"]:
        qty = item.get("qty", 0)
        unit_price = item.get("unit_price", 0)
        disc = item.get("discount", 0)
        line_total = qty * unit_price * (1 - disc)
        raw_cost = item.get("raw_cost", 0) or 0
        cost_disc = item.get("cost_disc", 0) or 0
        unit_cost = item.get("cost", 0) or 0
        ext_cost = unit_cost * qty
        profit = line_total - ext_cost
        gm = (profit / line_total) if line_total else 0

        subtotal += line_total

        ws.append([
            item.get("no", ""),
            item.get("vendor_no", ""),
            item.get("description", ""),
            qty,
            item.get("unit", ""),
            unit_price,
            disc,
            line_total,
            raw_cost,
            cost_disc / 100 if cost_disc else 0,
            unit_cost,
            ext_cost,
            profit,
            gm,
            item.get("delivery", ""),
            "Yes" if item.get("image") else "",
        ])

    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_amt = (subtotal + delivery_charge) * (inv.get("tax_rate", 0) or 0)
    total = subtotal + delivery_charge + tax_amt

    row = ws.max_row + 2
    summary_rows = [
        ("Invoice", inv.get("number", "")),
        ("Date", inv.get("date", "")),
        ("Client", inv.get("client_name", "")),
        ("Client No", inv.get("client_no", "")),
        ("Delivery Type", inv.get("delivery_type", "")),
        ("Payment Terms", inv.get("payment_terms", "")),
        ("Subtotal", subtotal),
        ("Delivery Charge", delivery_charge),
        ("Sales Tax", tax_amt),
        ("Total", total),
    ]

    for label, value in summary_rows:
        ws.cell(row, 1).value = label
        ws.cell(row, 1).fill = section_fill
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 2).value = value
        row += 1

    for r in ws.iter_rows():
        for cell in r:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in [6, 8, 9, 11, 12, 13]:
            ws.cell(row_idx, col_idx).number_format = "$#,##0"
        for col_idx in [7, 10, 14]:
            ws.cell(row_idx, col_idx).number_format = "0.0%"

    _autosize(ws, {
        1: 14, 2: 16, 3: 42, 4: 8, 5: 10, 6: 12, 7: 11, 8: 12,
        9: 12, 10: 11, 11: 12, 12: 14, 13: 12, 14: 10, 15: 22, 16: 14
    })

    xlsx_path = Path(output_path).with_name(Path(output_path).stem + "_INTERNAL.xlsx")
    wb.save(xlsx_path)
    return str(xlsx_path)


def draw_invoice(inv, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=0.55 * inch,
        rightMargin=0.55 * inch,
        topMargin=0.38 * inch,
        bottomMargin=1.00 * inch,
    )
    doc._gg_invoice = inv

    elements = []
    temp_images = []

    label_style = ParagraphStyle(
        "label_style",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        textColor=DARK,
        leading=10,
    )
    value_style = ParagraphStyle(
        "value_style",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=BLACK,
        leading=10,
    )
    cell_style = ParagraphStyle(
        "cell_style",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        textColor=BLACK,
        leading=10,
    )
    note_style = ParagraphStyle(
        "note_style",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.3,
        textColor=DARK,
        leading=10,
    )
    invoice_title_style = ParagraphStyle(
        "invoice_title_style",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=17,
        textColor=BLACK,
        alignment=1,
        spaceAfter=6,
    )

    # (1) Bigger centered logo
    if LOGO_PATH.exists():
        logo = Image(str(LOGO_PATH), width=2.10 * inch, height=1.35 * inch)
        logo.hAlign = "CENTER"
        elements.append(logo)
        elements.append(Spacer(1, 6))

    # (2) Left block indented to left page margin
    client_rows = [
        [Paragraph("Tel. | Mob.:", label_style), Paragraph(inv.get("client_phone", ""), value_style)],
        [Paragraph("Name:", label_style), Paragraph(inv.get("client_name", ""), value_style)],
    ]

    if inv.get("client_email"):
        client_rows.append([Paragraph("Email:", label_style), Paragraph(inv.get("client_email", ""), value_style)])

    for idx, line in enumerate(inv.get("client_address", []) or []):
        client_rows.append([
            Paragraph("Del. Address:" if idx == 0 else "", label_style),
            Paragraph(line, value_style),
        ])

    client_tbl = Table(client_rows, colWidths=[0.95 * inch, 3.45 * inch])
    client_tbl.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(client_tbl)
    elements.append(Spacer(1, 10))

    # Invoice title
    elements.append(Paragraph("Invoice", invoice_title_style))

    # (3) Meta table left aligned full width
    meta_data = [
        ["Invoice Date:", "Invoice", "Client No:", "Your Reference:"],
        [inv.get("date", ""), inv.get("number", ""), str(inv.get("client_no", "")), inv.get("reference", "")],
    ]
    meta_tbl = Table(meta_data, colWidths=[1.35 * inch, 1.30 * inch, 1.25 * inch, 3.05 * inch])
    meta_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, 1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(meta_tbl)
    elements.append(Spacer(1, 12))

    # Item table
    headers = ["ITEM NO.", "DESCRIPTION", "EST. DEL.", "TYPE", "QTY", "UNIT PRICE", "DISC.", "TOTAL", "PHOTO"]
    rows = [headers]

    for item in inv.get("items", []):
        qty = item.get("qty", 0)
        unit_price = item.get("unit_price", 0)
        disc = item.get("discount", 0)
        line_total = qty * unit_price * (1 - disc)

        photo_cell = ""
        img_path = _decode_image(item.get("image", ""))

        if img_path:
            temp_images.append(img_path)
            try:
                img = Image(img_path, width=0.82 * inch, height=0.52 * inch)
                img.hAlign = "CENTER"

                # (4) white background behind image cell so no dark box effect
                photo_cell = Table([[img]], colWidths=[0.90 * inch], rowHeights=[0.60 * inch])
                photo_cell.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
            except Exception:
                photo_cell = ""

        rows.append([
            Paragraph(str(item.get("no", "")), cell_style),
            Paragraph(item.get("description", ""), cell_style),
            Paragraph(item.get("delivery", ""), cell_style),
            Paragraph(item.get("unit", ""), cell_style),
            Paragraph(str(qty), cell_style),
            Paragraph(usd(unit_price), cell_style),
            Paragraph(f"{disc * 100:.0f}%" if disc else "", cell_style),
            Paragraph(usd(line_total), cell_style),
            photo_cell,
        ])

    item_tbl = Table(
        rows,
        colWidths=[0.65 * inch, 2.05 * inch, 0.95 * inch, 0.55 * inch, 0.35 * inch, 0.75 * inch, 0.45 * inch, 0.60 * inch, 0.95 * inch]
    )
    item_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER),
        ("TOPPADDING", (0, 1), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (3, 1), (8, -1), "CENTER"),
    ]))
    elements.append(item_tbl)
    elements.append(Spacer(1, 10))

    subtotal = sum(i.get("qty", 0) * i.get("unit_price", 0) * (1 - i.get("discount", 0)) for i in inv.get("items", []))
    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_amt = (subtotal + delivery_charge) * (inv.get("tax_rate", 0) or 0)
    total = subtotal + delivery_charge + tax_amt

    if inv.get("delivery_type"):
        elements.append(Paragraph(f"<b>{inv.get('delivery_type')}</b>", cell_style))
        elements.append(Spacer(1, 4))

    # (5) line after SubTotal and after Total
    totals_tbl = Table([
        ["SubTotal", usd(subtotal)],
        ["Delivery Charge", usd(delivery_charge)],
        ["Sales Tax", usd(tax_amt)],
        ["Total", usd(total)],
    ], colWidths=[1.7 * inch, 1.0 * inch])
    totals_tbl.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 0.8, BLACK),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, BORDER),
        ("LINEBELOW", (0, 3), (-1, 3), 0.8, BLACK),
        ("FONTNAME", (0, 0), (-1, -2), "Helvetica"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    totals_wrap = Table([["", totals_tbl]], colWidths=[4.5 * inch, 2.2 * inch])
    totals_wrap.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(totals_wrap)
    elements.append(Spacer(1, 12))

    pay_map = {
        "standard": (
            f"Payment is via check, bank transfer, or credit card. Please note that credit card payments incur a 3% processing fee. "
            f"Please refer to order no. {inv.get('number','')} and client no. {inv.get('client_no','')} with your payment."
        ),
        "advance": (
            f"Payment is paid in advance. Please refer to order no. {inv.get('number','')} and client no. {inv.get('client_no','')} with your payment."
        ),
        "installments": (
            f"Payment is in installments. Please refer to order no. {inv.get('number','')} and client no. {inv.get('client_no','')} with your payment."
        ),
    }

    elements.append(Paragraph(pay_map.get(inv.get("payment_terms", "standard"), pay_map["standard"]), note_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(
        "Bank account details #: 930283558  Routing number: 061092387.  Zelle email: rana_salah@goldenglam.nl",
        note_style
    ))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(
        "All quote(s), (provisional) order(s) (confirmations), sales and deliveries are subject to the Golden Glam Terms of orders and payments, the Golden Glam Reseller Terms and the CBM General Sales Terms and Conditions. US law applies.",
        note_style
    ))

    doc.build(elements, onFirstPage=_footer, onLaterPages=_footer)

    xlsx_path = _write_internal_excel(inv, output_path)

    for img in temp_images:
        try:
            os.remove(img)
        except Exception:
            pass

    return xlsx_path
