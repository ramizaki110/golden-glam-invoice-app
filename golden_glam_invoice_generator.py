import os
import io
import base64
import tempfile
import datetime
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

try:
    from PIL import Image as PILImage
except Exception:
    PILImage = None

# Romio
# ──────────────────────────────────────────────────────────────────────────────
# Layout / Styling
# ──────────────────────────────────────────────────────────────────────────────
PAGE_W, PAGE_H = letter
MARGIN_L = 0.65 * inch
MARGIN_R = 0.65 * inch
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R
RIGHT_X = MARGIN_L + CONTENT_W
FOOTER_H = 0.85 * inch

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "golden_glam_logo_final.png"

BLACK = colors.Color(0.14, 0.12, 0.11)
DARK_GRAY = colors.Color(0.28, 0.25, 0.22)
MID_GRAY = colors.Color(0.50, 0.47, 0.44)
LIGHT_GRAY = colors.Color(0.88, 0.86, 0.84)
HDR_BG = colors.Color(0.14, 0.12, 0.11)
ROW_TINT = colors.Color(0.98, 0.97, 0.96)

BASE_COL_W = [
    0.85 * inch,   # item no
    1.95 * inch,   # description
    1.15 * inch,   # est del
    0.70 * inch,   # type
    0.40 * inch,   # qty
    0.90 * inch,   # unit price
    0.58 * inch,   # disc
    0.78 * inch,   # total amount
]
IMG_COL_W = 1.35 * inch

HDR_LABELS = [
    "ITEM\nNO.",
    "DESCRIPTION",
    "ESTIMATED\nDEL. DATE",
    "TYPE",
    "QTY",
    "UNIT\nPRICE",
    "DISC.",
    "TOTAL\nAMOUNT",
]


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def usd(v):
    return f"${v:,.0f}"


def usd2(v):
    return f"${v:,.2f}"


def draw_hline(c, x1, x2, y, width=0.5, col=None):
    c.setLineWidth(width)
    c.setStrokeColor(col or LIGHT_GRAY)
    c.line(x1, y, x2, y)


def draw_logo(c, cx, y_top, w, h):
    if LOGO_PATH.exists():
        c.drawImage(
            str(LOGO_PATH),
            cx - w / 2,
            y_top - h,
            width=w,
            height=h,
            mask="auto",
            preserveAspectRatio=True,
        )


def _safe_date(date_str):
    try:
        return datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%m-%d-%Y")
    except Exception:
        return date_str


def _decode_base64_images(inv):
    """
    Converts any item["image"] data:image/...;base64,... into temp files
    so ReportLab can draw them. Returns list of temp file paths to clean up.
    """
    temp_files = []
    for item in inv["items"]:
        img = item.get("image", "") or ""
        if img.startswith("data:image"):
            try:
                header, b64 = img.split(",", 1)
                ext = "jpg"
                if "png" in header.lower():
                    ext = "png"
                data = base64.b64decode(b64)
                tmp = tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False)
                tmp.write(data)
                tmp.close()
                item["image"] = tmp.name
                temp_files.append(tmp.name)
            except Exception:
                item["image"] = ""
    return temp_files


# ──────────────────────────────────────────────────────────────────────────────
# PDF sections
# ──────────────────────────────────────────────────────────────────────────────
def draw_footer(c, page_num, total_pages, inv_number):
    base = 0.05 * inch
    draw_hline(c, MARGIN_L, RIGHT_X, base + 0.68 * inch, width=0.4, col=LIGHT_GRAY)

    c.setFont("Helvetica-Bold", 7.5)
    c.setFillColor(BLACK)
    c.drawCentredString(PAGE_W / 2, base + 0.52 * inch, "GOLDEN GLAM INTERIORS LLC")

    c.setFont("Helvetica", 6.5)
    c.setFillColor(DARK_GRAY)
    c.drawCentredString(
        PAGE_W / 2,
        base + 0.37 * inch,
        "Address: 828 Highland Ln Ne, Apt. 2204, Atlanta, GA 30306  |  Phone: 770-375-7343",
    )
    c.drawCentredString(
        PAGE_W / 2,
        base + 0.23 * inch,
        "Bank #: 930283558  |  Routing: 061092387  |  Zelle: rana_salah@goldenglam.nl  |  E-mail: sales@goldenglam.nl",
    )

    c.setFont("Helvetica", 6)
    c.setFillColor(MID_GRAY)
    c.drawRightString(
        RIGHT_X,
        base + 0.08 * inch,
        f"Invoice {inv_number}   •   Page {page_num} of {total_pages}",
    )


def draw_page1_header(c, inv):
    y = PAGE_H - 0.35 * inch
    draw_logo(c, PAGE_W / 2, y, 2.8 * inch, 1.2 * inch)

    y -= 1.2 * inch + 0.45 * inch
    lx, vx, lh = MARGIN_L, MARGIN_L + 0.82 * inch, 0.168 * inch

    def lv(cy, lbl, val, bold=False):
        c.setFont("Helvetica", 7.5)
        c.setFillColor(DARK_GRAY)
        c.drawString(lx, cy, lbl)

        c.setFont("Helvetica-Bold" if bold else "Helvetica", 7.5)
        c.setFillColor(BLACK)
        c.drawString(vx, cy, val or "")

    lv(y, "Tel. | Mob.:", inv.get("client_phone", ""))
    y -= lh
    lv(y, "Name:", inv.get("client_name", ""), bold=True)

    if inv.get("client_email"):
        y -= lh
        lv(y, "Email:", inv["client_email"])

    y -= lh
    addr_lines = inv.get("client_address", []) or []
    for i, line in enumerate(addr_lines):
        lv(y, "Del. Address:" if i == 0 else "", line)
        y -= lh

    return y - 0.10 * inch


def draw_cont_header(c, inv):
    y = PAGE_H - 0.32 * inch
    draw_logo(c, MARGIN_L + 0.55 * inch, y, 1.1 * inch, 0.55 * inch)

    c.setFont("Helvetica", 7)
    c.setFillColor(DARK_GRAY)
    c.drawRightString(
        RIGHT_X,
        y - 0.14 * inch,
        f"{inv['client_name']}  |  Invoice {inv['number']}  |  {inv['date']}  (continued)",
    )

    y -= 0.55 * inch + 0.12 * inch
    draw_hline(c, MARGIN_L, RIGHT_X, y, width=0.4, col=LIGHT_GRAY)
    return y - 0.18 * inch


def draw_invoice_meta(c, inv, y):
    bar_h = 0.22 * inch

    c.setFillColor(BLACK)
    c.rect(MARGIN_L, y - bar_h, CONTENT_W, bar_h, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(PAGE_W / 2, y - bar_h + 0.055 * inch, "Invoice")
    y -= bar_h + 0.13 * inch

    c.setFont("Helvetica", 7.5)
    c.setFillColor(DARK_GRAY)
    c.drawString(MARGIN_L, y, "Invoice Date:")
    c.drawString(MARGIN_L + 2.0 * inch, y, "Invoice")
    c.drawString(MARGIN_L + 3.8 * inch, y, "Client No:")
    c.drawString(MARGIN_L + 5.4 * inch, y, "Your Reference:")

    y -= 0.17 * inch
    c.setFont("Helvetica-Bold", 8.5)
    c.setFillColor(BLACK)
    c.drawString(MARGIN_L, y, _safe_date(inv.get("date", "")))
    c.drawString(MARGIN_L + 2.0 * inch, y, inv.get("number", ""))
    c.drawString(MARGIN_L + 3.8 * inch, y, str(inv.get("client_no", "")))
    if inv.get("reference"):
        c.drawString(MARGIN_L + 5.4 * inch, y, inv["reference"])

    y -= 0.13 * inch
    draw_hline(c, MARGIN_L, RIGHT_X, y, width=0.5, col=DARK_GRAY)
    return y - 0.18 * inch


def build_all_items_table(items):
    scale = (CONTENT_W - IMG_COL_W) / sum(BASE_COL_W)
    col_w = [w * scale for w in BASE_COL_W] + [IMG_COL_W]

    hdr = HDR_LABELS + ["PHOTO"]

    desc_style = ParagraphStyle(
        "desc",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        wordWrap="CJK",
    )
    cell_style = ParagraphStyle("cell", fontName="Helvetica", fontSize=8, leading=10)
    cell_style_c = ParagraphStyle("cellc", fontName="Helvetica", fontSize=8, leading=10, alignment=1)
    cell_style_r = ParagraphStyle("cellr", fontName="Helvetica", fontSize=8, leading=10, alignment=2)
    cell_style_b = ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=8, leading=10, alignment=2)

    def wrap_h(para, w):
        _, h = para.wrapOn(None, w - 10, 500)
        return h

    rows = [hdr]
    row_heights = [0.50 * inch]
    item_row_indices = []

    for item in items:
        disc = item.get("discount", 0.0)
        has_img = bool(item.get("image")) and os.path.exists(item["image"])

        cells = [
            Paragraph(str(item.get("no", "")), cell_style_c),
            Paragraph(item.get("description", ""), desc_style),
            Paragraph(item.get("delivery", ""), cell_style),
            Paragraph(item.get("unit", "Piece"), cell_style_c),
            Paragraph(str(item.get("qty", 1)), cell_style_c),
            Paragraph(usd(item.get("unit_price", 0)), cell_style_r),
            Paragraph(f"{disc * 100:.0f}%" if disc > 0 else "", cell_style_c),
            Paragraph(usd(item["_line_total"]), cell_style_b),
            "",
        ]

        cell_heights = [wrap_h(c, col_w[i]) for i, c in enumerate(cells[:-1])]
        min_h = 0.95 * inch if has_img else 0.42 * inch
        row_h = max(min_h, max(cell_heights) + 0.16 * inch)

        item_row_indices.append(len(rows))
        rows.append(cells)
        row_heights.append(row_h)

    t = Table(rows, colWidths=col_w, rowHeights=row_heights)
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), HDR_BG),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),

            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),

            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),

            ("LINEAFTER", (0, 1), (-2, -1), 0.3, colors.Color(0.84, 0.82, 0.80)),
            ("LINEBELOW", (0, 1), (-1, -1), 0.3, colors.Color(0.84, 0.82, 0.80)),
        ])
    )
    return t, col_w, row_heights, item_row_indices


# ──────────────────────────────────────────────────────────────────────────────
# Internal Excel
# ──────────────────────────────────────────────────────────────────────────────
def _autosize(ws, widths=None):
    if widths:
        for col_idx, width in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        return

    for col_cells in ws.columns:
        length = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                length = max(length, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(length + 2, 40)


def _write_internal_workbook(inv, pdf_path):
    xlsx_path = str(Path(pdf_path).with_name(Path(pdf_path).stem + "_INTERNAL.xlsx"))

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

    for col in range(1, len(headers) + 1):
        cell = ws.cell(1, col)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for item in inv["items"]:
        line_total = item["_line_total"]
        qty = item.get("qty", 1)
        unit_cost = float(item.get("cost", 0) or 0)
        ext_cost = unit_cost * qty
        profit = line_total - ext_cost if unit_cost else None
        gm = (profit / line_total) if (profit is not None and line_total) else None

        ws.append([
            item.get("no", ""),
            item.get("vendor_no", ""),
            item.get("description", ""),
            qty,
            item.get("unit_price", 0),
            (item.get("discount", 0) * 100.0) if item.get("discount") else 0,
            line_total,
            item.get("raw_cost", 0) or 0,
            item.get("cost_disc", 0) or 0,
            unit_cost,
            ext_cost,
            profit,
            gm,
            item.get("delivery", ""),
            "Yes" if item.get("image") else "",
        ])

    row = ws.max_row + 2
    subtotal = sum(i["_line_total"] for i in inv["items"])
    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_rate = inv.get("tax_rate", 0) or 0
    tax_amt = (subtotal + delivery_charge) * tax_rate
    total = subtotal + delivery_charge + tax_amt

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

    # Formats
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in [5, 7, 8, 10, 11, 12]:
            ws.cell(row_idx, col_idx).number_format = "$#,##0.00"
        ws.cell(row_idx, 13).number_format = "0.0%"

    _autosize(ws, {
        1: 14, 2: 14, 3: 42, 4: 8, 5: 12, 6: 11, 7: 12,
        8: 12, 9: 11, 10: 12, 11: 14, 12: 12, 13: 10, 14: 22, 15: 14
    })

    # Vendor mapping sheet
    ws2 = wb.create_sheet("Vendor Numbers")
    hdr2 = ["Item No", "Description", "Vendor No", "Raw Cost", "Cost Disc %", "Unit Cost"]
    ws2.append(hdr2)
    for col in range(1, len(hdr2) + 1):
        cell = ws2.cell(1, col)
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
            item.get("cost_disc", 0) or 0,
            item.get("cost", 0) or 0,
        ])

    for r in ws2.iter_rows():
        for cell in r:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row_idx in range(2, ws2.max_row + 1):
        ws2.cell(row_idx, 4).number_format = "$#,##0.00"
        ws2.cell(row_idx, 6).number_format = "$#,##0.00"

    _autosize(ws2, {1: 14, 2: 42, 3: 18, 4: 12, 5: 11, 6: 12})

    wb.save(xlsx_path)
    return xlsx_path


# ──────────────────────────────────────────────────────────────────────────────
# Main draw function
# ──────────────────────────────────────────────────────────────────────────────
def draw_invoice(inv, output_path):
    temp_files = _decode_base64_images(inv)

    for item in inv["items"]:
        disc = item.get("discount", 0) or 0
        item["_line_total"] = item.get("qty", 1) * item.get("unit_price", 0) * (1 - disc)

    subtotal = sum(i["_line_total"] for i in inv["items"])
    delivery_charge = inv.get("delivery_charge", 0) or 0
    tax_rate = inv.get("tax_rate", 0) or 0
    tax_amt = (subtotal + delivery_charge) * tax_rate
    grand_total = subtotal + delivery_charge + tax_amt

    c = rl_canvas.Canvas(output_path, pagesize=letter)
    c.setTitle(f"Invoice {inv['number']} - Golden Glam Interiors")

    page_num = 1

    def new_page():
        nonlocal page_num
        draw_footer(c, page_num, "...", inv["number"])
        c.showPage()
        page_num += 1
        return draw_cont_header(c, inv)

    y = draw_page1_header(c, inv)
    y = draw_invoice_meta(c, inv, y)

    t, col_w, row_heights, item_row_indices = build_all_items_table(inv["items"])
    table_h = sum(row_heights)
    if y - table_h < FOOTER_H + 0.15 * inch:
        y = new_page()

    t.wrapOn(c, sum(col_w), table_h)
    t.drawOn(c, MARGIN_L, y - table_h)

    # Draw product photos inside last column
    base_w = sum(col_w[:-1])
    img_x = MARGIN_L + base_w
    cell_w = col_w[-1]

    for item, ri in zip(inv["items"], item_row_indices):
        img_path = item.get("image")
        if not img_path or not os.path.exists(img_path) or PILImage is None:
            continue

        row_top = y - sum(row_heights[:ri])
        cell_h = row_heights[ri]
        img_yb = row_top - cell_h

        try:
            with PILImage.open(img_path) as pimg:
                pw, ph = pimg.size

            scale = max(cell_w / pw, cell_h / ph)
            draw_w = pw * scale
            draw_h = ph * scale
            cx = img_x + (cell_w - draw_w) / 2
            cy = img_yb + (cell_h - draw_h) / 2

            c.saveState()
            p = c.beginPath()
            p.rect(img_x, img_yb, cell_w, cell_h)
            c.clipPath(p, stroke=0, fill=0)
            c.drawImage(img_path, cx, cy, width=draw_w, height=draw_h, mask="auto")
            c.restoreState()
        except Exception:
            pass

    y -= table_h + 0.10 * inch

    # Totals
    TOTALS_LEFT = MARGIN_L + 3.50 * inch
    if y - 1.65 * inch < FOOTER_H + 0.15 * inch:
        y = new_page()

    if inv.get("delivery_type"):
        y -= 0.18 * inch
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(BLACK)
        c.drawRightString(RIGHT_X, y, inv["delivery_type"])
        draw_hline(c, TOTALS_LEFT, RIGHT_X, y - 0.04 * inch, width=0.8, col=BLACK)
        y -= 0.23 * inch

    def tot_row(cy, label, value, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 8.5 if bold else 8)
        c.setFillColor(BLACK)
        c.drawString(TOTALS_LEFT, cy, label)
        c.drawRightString(RIGHT_X, cy, value)

    tot_row(y, "SubTotal", usd(subtotal))
    y -= 0.21 * inch
    tot_row(y, "Delivery Charge", usd(delivery_charge))
    y -= 0.21 * inch
    tot_row(y, "Sales Tax", usd(tax_amt))
    y -= 0.07 * inch
    draw_hline(c, TOTALS_LEFT, RIGHT_X, y, width=0.8, col=BLACK)
    y -= 0.21 * inch
    tot_row(y, "Total", usd(grand_total), bold=True)
    y -= 0.38 * inch

    # Notes
    if inv.get("notes"):
        sty = ParagraphStyle(
            "n",
            fontName="Helvetica",
            fontSize=7.5,
            leading=11,
            textColor=DARK_GRAY,
        )
        p = Paragraph(f"<b>Note:</b> {inv['notes']}", sty)
        _, h = p.wrapOn(c, CONTENT_W, 2 * inch)
        if y - h < FOOTER_H + 0.1 * inch:
            y = new_page()
        p.drawOn(c, MARGIN_L, y - h)
        y -= h + 0.18 * inch

    pay_map = {
        "standard": (
            f"Payment is via check, bank transfer, or credit card. Please note that credit card "
            f"payments incur a 3% processing fee. Please refer to order no. {inv['number']} "
            f"and client no. {inv['client_no']} with your payment."
        ),
        "advance": (
            f"Payment is paid in advance. Please refer to order no. {inv['number']} "
            f"and client no. {inv['client_no']} with your payment."
        ),
        "installments": (
            f"Payment is in installments. Please refer to order no. {inv['number']} "
            f"and client no. {inv['client_no']} with your payment."
        ),
    }

    sty2 = ParagraphStyle(
        "p",
        fontName="Helvetica",
        fontSize=7,
        leading=10,
        textColor=DARK_GRAY,
    )
    pp = Paragraph(pay_map.get(inv.get("payment_terms", "standard"), pay_map["standard"]), sty2)
    _, ph = pp.wrapOn(c, CONTENT_W, 2 * inch)
    if y - ph < FOOTER_H + 0.1 * inch:
        y = new_page()
    pp.drawOn(c, MARGIN_L, y - ph)
    y -= ph + 0.10 * inch

    c.setFont("Helvetica-Bold", 7)
    c.setFillColor(BLACK)
    c.drawString(
        MARGIN_L,
        y,
        "Bank account details #: 930283558  Routing number: 061092387.  Zelle email: rana_salah@goldenglam.nl",
    )
    y -= 0.18 * inch

    legal = (
        "All quote(s), (provisional) order(s) (confirmations), sales and deliveries are subject "
        "to the Golden Glam Terms of orders and payments, the Golden Glam Reseller Terms and the "
        "CBM General Sales Terms and Conditions, which may be consulted at www.goldenglam.nl. "
        "US law applies. The courts in the Hague have exclusive jurisdiction."
    )
    sty3 = ParagraphStyle(
        "l",
        fontName="Helvetica",
        fontSize=6.5,
        leading=9,
        textColor=MID_GRAY,
    )
    lp = Paragraph(legal, sty3)
    _, lh = lp.wrapOn(c, CONTENT_W, 2 * inch)
    if y - lh < FOOTER_H + 0.1 * inch:
        y = new_page()
    lp.drawOn(c, MARGIN_L, y - lh)
    y -= lh + 0.13 * inch

    terms_map = {
        "advance": "Term of Delivery: Delivered at Place.  Term of payment: Paid in advance",
        "installments": "Term of Delivery: Delivered at Place.  Term of payment: Payment in installments",
    }
    if inv.get("payment_terms") in terms_map:
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(DARK_GRAY)
        c.drawString(MARGIN_L, y, terms_map[inv["payment_terms"]])
        y -= 0.21 * inch

    draw_hline(c, MARGIN_L, RIGHT_X, y, width=0.4)
    y -= 0.15 * inch
    c.setFont("Helvetica", 7)
    c.setFillColor(MID_GRAY)
    c.drawString(MARGIN_L, y, f"Print date:    {inv.get('date', '')}")

    draw_footer(c, page_num, page_num, inv["number"])
    c.save()

    _write_internal_workbook(inv, output_path)

    for f in temp_files:
        try:
            os.remove(f)
        except Exception:
            pass


# Optional compatibility helper
def generate_from_summary(summary_path, output_dir):
    raise NotImplementedError(
        "The hosted app should call draw_invoice(...) through golden_glam_invoice_webapp.py, "
        "not generate_from_summary(...) directly."
    )
