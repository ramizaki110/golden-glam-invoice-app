from __future__ import annotations

import os
import re
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from openpyxl import load_workbook

from golden_glam_invoice_generator import draw_invoice

BASE_DIR = Path(__file__).resolve().parent
HTML_FILE = BASE_DIR / "GoldenGlam_InvoiceGenerator_hosted.html"
OUTPUT_DIR = BASE_DIR / "generated_invoices"
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)


def _clean_currency(x: str) -> float:
    return float(x.replace(",", "").strip())


def parse_summary(summary_text: str) -> dict:
    m = re.search(r"GG-INV\|([^|]+)\|([^|]+)\|ref:(.*)", summary_text)
    if not m:
        raise ValueError("Could not read invoice number/date from summary.")
    number, date, ref = m.groups()

    client_name = ""
    client_no = ""
    client_phone = ""
    client_email = ""

    m = re.search(r"client:(.*?)\|no:(.*?)\|ph:(.*?)(?:\|email:(.*))?$", summary_text, re.M)
    if not m:
        raise ValueError("Could not read client section from summary.")
    client_name = m.group(1).strip()
    client_no = m.group(2).strip()
    client_phone = m.group(3).strip()
    client_email = (m.group(4) or "").strip()

    m = re.search(r"addr:(.+)", summary_text)
    addr_line = m.group(1).strip() if m else ""
    client_address = [p.strip() for p in addr_line.split(",") if p.strip()]

    m = re.search(r"file:(.+)", summary_text)
    base_filename = m.group(1).strip() if m else f"GG__{number}_{client_name.upper().replace(' ', '_')}"

    delivery_type = ""
    delivery_charge = 0.0
    tax_rate = 0.0
    m = re.search(
        r"\n([^\n|]*Delivery(?: \([^)]+\))?)\|sub:\$(-?[0-9.,]+)\|del:\$(-?[0-9.,]+)\|tax\((\d+(?:\.\d+)?)%\):\$(-?[0-9.,]+)\|total:\$(-?[0-9.,]+)",
        summary_text
    )
    if m:
        delivery_type = m.group(1).strip()
        delivery_charge = _clean_currency(m.group(3))
        tax_rate = float(m.group(4)) / 100.0

    payment_terms = "standard"
    low = summary_text.lower()
    if "pay:paid in advance" in low:
        payment_terms = "advance"
    elif "pay:payment in installments" in low or "pay:installments" in low:
        payment_terms = "installments"

    notes = ""
    m = re.search(r"\nnotes:(.+?)(?:\nINTERNAL|\Z)", summary_text, re.S)
    if m:
        notes = m.group(1).strip()

    parts = summary_text.split("ITEMS", 1)
    if len(parts) < 2:
        raise ValueError("Could not find invoice items in summary.")

    items_section = parts[1]

    if "INTERNAL" in items_section:
        items_part, internal_part = items_section.split("INTERNAL", 1)
    else:
        items_part, internal_part = items_section, ""

    # Remove the helper line:
    # (fmt: [itemNo]desc|type|qty:N|$unitPrice|disc:N%|tot:$N|del:text|Photo-base64:...)
    item_lines = items_part.splitlines()
    cleaned_lines = []
    for line in item_lines:
        if line.strip().startswith("(fmt:"):
            continue
        cleaned_lines.append(line)
    items_part = "\n".join(cleaned_lines).strip()

    item_pattern = re.compile(
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$(-?[0-9.,]+)(?:\|disc:([0-9.]+)%?)?\|tot:\$(-?[0-9.,]+)\|del:(.*?)(?:\n\s*Photo-base64:\s*(data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=\s]+))?(?=\n\[|\n[^\n|]*Delivery(?: \([^)]+\))?\||\npay:|\nnotes:|\nINTERNAL|\Z)",
        re.S,
    )

    items = []
    for g in item_pattern.finditer(items_part):
        no, desc, unit, qty, unit_price, disc, _total, delivery, photo = g.groups()
        items.append({
            "no": no.strip(),
            "description": " ".join(desc.strip().split()),
            "delivery": " ".join(delivery.strip().split()),
            "qty": int(qty),
            "unit": unit.strip(),
            "unit_price": _clean_currency(unit_price),
            "discount": (float(disc) / 100.0) if disc else 0.0,
            "image": re.sub(r"\s+", "", photo) if photo else "",
        })

    if not items:
        raise ValueError("Could not parse any line items from summary.")

    internal_map = {}
    for line in internal_part.splitlines():
        line = line.strip()
        if not line.startswith("["):
            continue

        item_no_match = re.match(r"\[([^\]]+)\]", line)
        if not item_no_match:
            continue
        item_no = item_no_match.group(1)

        vendor_no = ""
        vendor_name = ""
        raw_cost = 0.0
        cost_disc = 0.0
        cost = 0.0

        m = re.search(r'vendor_no:\s*"([^"]*)"', line)
        if m:
            vendor_no = m.group(1)

        m = re.search(r'vendor:\s*"([^"]*)"', line)
        if m:
            vendor_name = m.group(1)

        m = re.search(r"raw_cost:\s*([0-9.]+)", line)
        if m:
            raw_cost = float(m.group(1))

        m = re.search(r"cost_disc:\s*([0-9.]+)%", line)
        if m:
            cost_disc = float(m.group(1))

        m = re.search(r"cost:\s*([0-9.]+)", line)
        if m:
            cost = float(m.group(1))

        internal_map[item_no] = {
            "vendor_name": vendor_name,
            "vendor_no": vendor_no,
            "raw_cost": raw_cost,
            "cost_disc": cost_disc,
            "cost": cost,
        }

    for item in items:
        item.update(internal_map.get(item["no"], {}))

    return {
        "number": number.strip(),
        "date": date.strip(),
        "reference": ref.strip(),
        "client_name": client_name,
        "client_no": client_no,
        "client_phone": client_phone,
        "client_email": client_email,
        "client_address": client_address,
        "delivery_type": delivery_type,
        "delivery_charge": delivery_charge,
        "tax_rate": tax_rate,
        "payment_terms": payment_terms,
        "notes": notes,
        "items": items,
        "base_filename": base_filename,
    }


def reformat_excel_no_decimals(xlsx_path: Path) -> None:
    if not xlsx_path.exists():
        return

    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "$#,##0"
    wb.save(xlsx_path)


def generate_from_summary(summary_text: str) -> tuple[Path, Path]:
    invoice = parse_summary(summary_text)
    safe_name = invoice.pop("base_filename")
    pdf_path = OUTPUT_DIR / f"{safe_name}.pdf"
    xlsx_path = draw_invoice(invoice, str(pdf_path))
    xlsx_path = Path(xlsx_path)
    return pdf_path, xlsx_path


@app.get("/")
def home():
    return send_from_directory(BASE_DIR, HTML_FILE.name)


@app.post("/generate")
def generate():
    payload = request.get_json(silent=True) or {}
    summary = (payload.get("summary") or "").strip()
    if not summary:
        return jsonify({"ok": False, "error": "Missing invoice summary."}), 400

    try:
        pdf_path, xlsx_path = generate_from_summary(summary)
        return jsonify({
            "ok": True,
            "pdf_name": pdf_path.name,
            "xlsx_name": xlsx_path.name,
            "pdf_url": f"/download/{pdf_path.name}",
            "xlsx_url": f"/download/{xlsx_path.name}",
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.get("/download/<path:filename>")
def download(filename: str):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
