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
        r"\n([^\n|]+)\|sub:\$(-?[0-9.,]+)\|del:\$(-?[0-9.,]+)\|tax\((\d+(?:\.\d+)?)%\):\$(-?[0-9.,]+)\|total:\$(-?[0-9.,]+)",
        summary_text
    )
    if m:
        delivery_type = m.group(1).strip()
        delivery_charge = _clean_currency(m.group(3))
        tax_rate = float(m.group(4)) / 100.0

    payment_terms = "advance"  # default is now advance (no standard option)
    low = summary_text.lower()
    if "pay:in installments" in low or "pay:installments" in low or "pay:payment in installments" in low:
        payment_terms = "installments"
    elif "pay:paid in advance" in low:
        payment_terms = "advance"

    notes = ""
    m = re.search(r"\nnotes:(.+?)(?:\ninstallments:|\nINTERNAL|\Z)", summary_text, re.S)
    if m:
        notes = m.group(1).strip()

    # Parse installments
    installments = []
    installment_split_type = "amount"
    m_inst = re.search(r"\ninstallments:(.+?)(?:\nINTERNAL|\nITEMS|\Z)", summary_text, re.S)
    if m_inst:
        for part in m_inst.group(1).strip().split(";"):
            seg = part.strip().split("|")
            if len(seg) >= 3:
                installments.append({"date": seg[1], "val": seg[2]})
                if len(seg) >= 4:
                    installment_split_type = seg[3]

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
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$(-?[0-9.,]+)(?:\|disc:([0-9.]+)%?)?\|tot:\$(-?[0-9.,]+)(?:\|del:([^\n]*))?(?:\n\s*Photo-base64:\s*(data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+))?(?=\n\[|\n[^\n]+\|sub:\$|\npay:|\nnotes:|\ninstallments:|\nINTERNAL|\Z)",
        re.S,
    )

    items = []
    for g in item_pattern.finditer(items_part):
        no, desc, unit, qty, unit_price, disc, _total, delivery, photo = g.groups()
        items.append({
            "no": no.strip(),
            "description": " ".join(desc.strip().split()).replace('⁄', '|'),
            "delivery": " ".join((delivery or "").strip().split()).replace('⁄', '|'),
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

        m = re.search(r"(?<![a-z_])cost:\s*([0-9.]+)", line)
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
        "installments": installments,
        "installment_split_type": installment_split_type,
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



# ── Supabase persistent storage ────────────────────────────────────────────────
import json
import urllib.request
import urllib.error

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
TABLE        = "gg_data"
ROW_KEY      = "main"

EMPTY_DATA = {"clients": [], "library": [], "vendors": [], "invoices": []}

def _supabase_headers():
    return {
        "apikey":        SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type":  "application/json",
        "Prefer":        "return=representation",
    }

def _load_data() -> dict:
    if not SUPABASE_URL or not SUPABASE_KEY:
        return dict(EMPTY_DATA)
    try:
        url = f"{SUPABASE_URL}/rest/v1/{TABLE}?key=eq.{ROW_KEY}&select=value"
        req = urllib.request.Request(url, headers=_supabase_headers())
        with urllib.request.urlopen(req, timeout=8) as resp:
            rows = json.loads(resp.read())
            if rows:
                return rows[0].get("value") or dict(EMPTY_DATA)
    except Exception as e:
        print(f"[data] load error: {e}")
    return dict(EMPTY_DATA)

def _save_data(data: dict):
    if not SUPABASE_URL or not SUPABASE_KEY:
        return
    try:
        body = json.dumps({"key": ROW_KEY, "value": data}).encode()
        hdrs = {**_supabase_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"}
        url  = f"{SUPABASE_URL}/rest/v1/{TABLE}"
        req  = urllib.request.Request(url, data=body, headers=hdrs, method="POST")
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        print(f"[data] save error: {e}")


@app.get("/api/data")
def api_get_data():
    return jsonify(_load_data())


@app.post("/api/data")
def api_save_data():
    payload = request.get_json(silent=True) or {}
    data    = _load_data()
    for key in ("clients", "library", "vendors", "invoices"):
        if key in payload:
            data[key] = payload[key]
    _save_data(data)
    return jsonify({"ok": True})

@app.get("/api/photo/<path:key>")
def api_get_photo(key):
    if not SUPABASE_URL or not SUPABASE_KEY:
        return jsonify({"ok": False, "error": "No DB"}), 503
    try:
        import urllib.parse
        url = f"{SUPABASE_URL}/rest/v1/gg_photos?key=eq.{urllib.parse.quote(key)}&select=data"
        req = urllib.request.Request(url, headers=_supabase_headers())
        with urllib.request.urlopen(req, timeout=8) as resp:
            rows = json.loads(resp.read())
            if rows:
                return jsonify({"ok": True, "data": rows[0]["data"]})
            return jsonify({"ok": False, "data": ""})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.post("/api/photo")
def api_save_photo():
    if not SUPABASE_URL or not SUPABASE_KEY:
        return jsonify({"ok": False, "error": "No DB"}), 503
    payload = request.get_json(silent=True) or {}
    key  = payload.get("key", "")
    data = payload.get("data", "")
    if not key:
        return jsonify({"ok": False, "error": "Missing key"}), 400
    try:
        body = json.dumps({"key": key, "data": data}).encode()
        hdrs = {**_supabase_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"}
        url  = f"{SUPABASE_URL}/rest/v1/gg_photos"
        req  = urllib.request.Request(url, data=body, headers=hdrs, method="POST")
        urllib.request.urlopen(req, timeout=10)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.delete("/api/photo/<path:key>")
def api_delete_photo(key):
    if not SUPABASE_URL or not SUPABASE_KEY:
        return jsonify({"ok": True})
    try:
        import urllib.parse
        url = f"{SUPABASE_URL}/rest/v1/gg_photos?key=eq.{urllib.parse.quote(key)}"
        req = urllib.request.Request(url, headers=_supabase_headers(), method="DELETE")
        urllib.request.urlopen(req, timeout=8)
    except Exception:
        pass
    return jsonify({"ok": True})



# ── Competitor Pricing ─────────────────────────────────────────────────────────
@app.post("/api/price-check")
def api_price_check():
    import urllib.parse as urlparse

    payload  = request.get_json(silent=True) or {}
    image_b64    = payload.get("image", "")
    product_text = payload.get("product", "").strip()
    sku          = payload.get("sku", "").strip()
    zip_code     = payload.get("zip", "30338").strip()

    GOOGLE_KEY   = os.environ.get("GOOGLE_SEARCH_API_KEY", "")
    GOOGLE_CX    = os.environ.get("GOOGLE_SEARCH_CX", "")
    if not GOOGLE_KEY or not GOOGLE_CX:
        return jsonify({"ok": False,
            "error": "Google Search API not configured. Add GOOGLE_SEARCH_API_KEY and GOOGLE_SEARCH_CX to your Render environment variables. See setup guide in Notion."}), 503

    # ── Step 1: Build search query from text (no Claude needed) ─────────────────
    identified = product_text
    if not identified and not sku:
        if image_b64:
            return jsonify({"ok": False,
                "error": "Please also type the product name or SKU in the fields below the image — this lets us search without any AI identification cost."}), 400
        return jsonify({"ok": False, "error": "Please provide a product name or SKU."}), 400

    # ── Step 2: Google Custom Search ──────────────────────────────────────────
    parts = []
    if identified: parts.append(identified)
    if sku:        parts.append(sku)
    query = " ".join(parts) + " price"

    try:
        search_url = (f"https://www.googleapis.com/customsearch/v1"
                      f"?key={GOOGLE_KEY}&cx={GOOGLE_CX}"
                      f"&q={urlparse.quote(query)}&num=10&gl=us&cr=countryUS")
        with urllib.request.urlopen(urllib.request.Request(search_url), timeout=12) as resp:
            items = json.loads(resp.read()).get("items", [])
    except Exception as e:
        return jsonify({"ok": False, "error": f"Google search failed: {e}"}), 500

    if not items:
        return jsonify({"ok": True, "identified_product": "", "results": [],
                        "floor": None, "ceiling": None, "suggested": None, "query_used": query})

    # ── Step 3: Extract prices from search snippets (regex, no Claude needed) ─
    import re
    results = []
    for it in items:
        snippet = it.get("snippet", "") + " " + it.get("title", "")
        prices = re.findall(r'\$([\d,]+(?:\.\d{2})?)', snippet)
        if prices:
            try:
                price = float(prices[0].replace(",", ""))
                domain = urlparse.urlparse(it["link"]).netloc.replace("www.", "")
                results.append({"retailer": domain, "price": price,
                                "url": it["link"], "title": it.get("title",""), "in_stock": True})
            except Exception:
                pass

    # ── Step 4: Compute price range ───────────────────────────────────────────
    prices = [r["price"] for r in results if isinstance(r.get("price"), (int, float)) and r["price"] > 0]
    floor_p   = round(min(prices), 2) if prices else None
    ceiling_p = round(max(prices), 2) if prices else None
    avg_p     = sum(prices) / len(prices) if prices else None
    # Suggested = 5% above average, rounded to nearest $5
    suggested = (round(avg_p * 1.05 / 5) * 5) if avg_p else None

    return jsonify({"ok": True, "identified_product": "", "results": results,
                    "floor": floor_p, "ceiling": ceiling_p, "suggested": suggested, "query_used": query})

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
