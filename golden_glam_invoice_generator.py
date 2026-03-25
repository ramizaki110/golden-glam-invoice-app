import os
import base64
import tempfile
from pathlib import Path

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
BORDER = colors.HexColor("#e0dedd")
GOLD = colors.HexColor("#b8963e")
BG = colors.HexColor("#faf9f8")


def usd(v: float) -> str:
    return f"${v:,.0f}"


def usd2(v: float) -> str:
    return f"${v:,.2f}"


def _decode_image(image_value: str) -> str | None:
    if not image_value:
        return None
    if not image_value.startswith("data:image"):
        return image_value if os.path.exists(image_value) else None

    try:
        header, b64 = image_value.split(",", 1)
        ext = "jpg"
        if "png" in header.lower():
            ext = "png"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
        tmp.write(base64.b64decode(b64))
        tmp.close()
        return tmp.name
    except Exception:
        return None


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(BORDER)
    canvas.line(doc.leftMargin, 0.65 * inch, letter[0] - doc.rightMargin, 0.65 * inch)

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(BLACK)
    canvas.drawCentredString(letter[0] / 2, 0.48 * inch, "GOLDEN GLAM INTERIORS LLC")

    canvas.setFont("Helvetica", 6)
    canvas.setFillColor(DARK)
    canvas.drawCentredString(
        letter[0] / 2,
        0.34 * inch,
        "Address: 828 Highland Ln Ne, Apt. 2204, Atlanta, GA 30306  |  Phone: 770-375-7343",
    )
    canvas.drawCentredString(
        letter[0] / 2,
        0.22 * inch,
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
    total_cost = 0

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
        total_cost += ext_cost

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
            cost_disc,
            unit_cost,
            ext_cost,
            profit,
            gm,
            item.get("delivery", ""),
            "Yes" if item.get("image") else "",
        ])

    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_rate = inv.get("tax_rate", 0) or 0
    tax_amt = (subtotal + delivery_charge) * tax_rate
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
        for col_idx in [7, 10]:
            ws.cell(row_idx, col_idx).number_format = "0.0%"
        ws.cell(row_idx, 14).number_format = "0.0%"

    _autosize(ws, {
        1: 14, 2: 16, 3: 42, 4: 8, 5: 10, 6: 12, 7: 11, 8: 12,
        9: 12, 10: 11, 11: 12, 12: 14, 13: 12, 14: 10, 15: 22, 16: 14
    })

    ws2 = wb.create_sheet("Vendor Numbers")
    hdr2 = ["Item No", "Description", "Vendor No", "Raw Cost", "Cost Disc %", "Unit Cost"]
    ws2.append(hdr2)

    for c in range(1, len(hdr2) + 1):
        cell = ws2.cell(1, c)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for item in inv["items"]:
        ws2.append([
            item.get("no", ""),
            item.get("description", ""),
            item.get("vendor_no", ""),
            item.get("raw_cost", 0) or 0,
            (item.get("cost_disc", 0) or 0) / 100,
            item.get("cost", 0) or 0,
        ])

    for r in ws2.iter_rows():
        for cell in r:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx in range(2, ws2.max_row + 1):
        ws2.cell(row_idx, 4).number_format = "$#,##0"
        ws2.cell(row_idx, 5).number_format = "0.0%"
        ws2.cell(row_idx, 6).number_format = "$#,##0"

    _autosize(ws2, {1: 14, 2: 42, 3: 18, 4: 12, 5: 12, 6: 12})

    xlsx_path = Path(output_path).with_name(Path(output_path).stem + "_INTERNAL.xlsx")
    wb.save(xlsx_path)
    return str(xlsx_path)


def draw_invoice(inv, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=0.6 * inch,
        rightMargin=0.6 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.9 * inch,
    )

    elements = []
    temp_images = []

    title_style = ParagraphStyle(
        "gg_title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=BLACK,
        spaceAfter=8,
        alignment=1,
    )
    small_label = ParagraphStyle(
        "gg_label",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        textColor=DARK,
        leading=10,
    )
    small_val = ParagraphStyle(
        "gg_val",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=BLACK,
        leading=10,
    )
    normal = ParagraphStyle(
        "gg_normal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        textColor=BLACK,
        leading=11,
    )
    note_style = ParagraphStyle(
        "gg_note",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        textColor=DARK,
        leading=10,
    )

    # Header
    header_parts = []
    if LOGO_PATH.exists():
        header_parts.append(Image(str(LOGO_PATH), width=1.9 * inch, height=0.8 * inch))
    else:
        header_parts.append(Paragraph("Golden Glam", title_style))

    header_parts.append(Paragraph("Invoice", title_style))
    header_tbl = Table([[header_parts[0], header_parts[1]]], colWidths=[2.3 * inch, 4.5 * inch])
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(header_tbl)
    elements.append(Spacer(1, 8))

    # Client + Meta
    client_lines = [
        [Paragraph("Tel. | Mob.:", small_label), Paragraph(inv.get("client_phone", ""), small_val)],
        [Paragraph("Name:", small_label), Paragraph(inv.get("client_name", ""), small_val)],
    ]
    if inv.get("client_email"):
        client_lines.append([Paragraph("Email:", small_label), Paragraph(inv.get("client_email", ""), small_val)])

    addr = inv.get("client_address", []) or []
    for idx, line in enumerate(addr):
        client_lines.append([
            Paragraph("Del. Address:" if idx == 0 else "", small_label),
            Paragraph(line, small_val),
        ])

    client_tbl = Table(client_lines, colWidths=[0.95 * inch, 2.6 * inch])
    client_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))

    meta_data = [
        ["Invoice Date:", "Invoice", "Client No:", "Your Reference:"],
        [inv.get("date", ""), inv.get("number", ""), str(inv.get("client_no", "")), inv.get("reference", "")],
    ]
    meta_tbl = Table(meta_data, colWidths=[1.0 * inch, 1.15 * inch, 0.95 * inch, 1.65 * inch])
    meta_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, 1), 8.5),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    top_tbl = Table([[client_tbl, meta_tbl]], colWidths=[3.7 * inch, 3.2 * inch])
    top_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(top_tbl)
    elements.append(Spacer(1, 12))

    # Items
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
                photo_cell = Image(img_path, width=0.9 * inch, height=0.7 * inch)
            except Exception:
                photo_cell = ""

        rows.append([
            Paragraph(str(item.get("no", "")), normal),
            Paragraph(item.get("description", ""), normal),
            Paragraph(item.get("delivery", ""), normal),
            Paragraph(item.get("unit", ""), normal),
            Paragraph(str(qty), normal),
            Paragraph(usd(unit_price), normal),
            Paragraph(f"{disc*100:.0f}%" if disc else "", normal),
            Paragraph(usd(line_total), normal),
            photo_cell,
        ])

    item_tbl = Table(
        rows,
        colWidths=[0.8 * inch, 2.15 * inch, 1.0 * inch, 0.65 * inch, 0.4 * inch, 0.8 * inch, 0.5 * inch, 0.7 * inch, 1.0 * inch]
    )
    item_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (3, 1), (7, -1), "CENTER"),
    ]))
    elements.append(item_tbl)
    elements.append(Spacer(1, 10))

    # Totals
    subtotal = sum(i.get("qty", 0) * i.get("unit_price", 0) * (1 - i.get("discount", 0)) for i in inv.get("items", []))
    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_amt = (subtotal + delivery_charge) * (inv.get("tax_rate", 0) or 0)
    total = subtotal + delivery_charge + tax_amt

    if inv.get("delivery_type"):
        elements.append(Paragraph(f"<b>{inv.get('delivery_type')}</b>", normal))
        elements.append(Spacer(1, 4))

    totals_tbl = Table([
        ["SubTotal", usd(subtotal)],
        ["Delivery Charge", usd(delivery_charge)],
        ["Sales Tax", usd(tax_amt)],
        ["Total", usd(total)],
    ], colWidths=[1.7 * inch, 1.1 * inch])
    totals_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -2), "Helvetica"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LINEABOVE", (0, -1), (-1, -1), 0.8, BLACK),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    totals_wrap = Table([["", totals_tbl]], colWidths=[4.5 * inch, 2.3 * inch])
    totals_wrap.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(totals_wrap)
    elements.append(Spacer(1, 12))

    # Notes / payment
    if inv.get("notes"):
        elements.append(Paragraph(f"<b>Note:</b> {inv.get('notes')}", note_style))
        elements.append(Spacer(1, 8))

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
        "All quote(s), (provisional) order(s) (confirmations), sales and deliveries are subject to the Golden Glam Terms of orders and payments, "
        "the Golden Glam Reseller Terms and the CBM General Sales Terms and Conditions. US law applies.",
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
