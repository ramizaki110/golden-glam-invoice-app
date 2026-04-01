from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
import urllib.parse
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from openpyxl import load_workbook

from golden_glam_invoice_generator import draw_invoice

BASE_DIR   = Path(__file__).resolve().parent
HTML_FILE  = BASE_DIR / "GoldenGlam_InvoiceGenerator_hosted.html"
OUTPUT_DIR = BASE_DIR / "generated_invoices"
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)


# ── Invoice parsing helpers ────────────────────────────────────────────────────

def _clean_currency(x: str) -> float:
    return float(x.replace(",", "").strip())


def parse_summary(summary_text: str) -> dict:
    m = re.search(r"GG-INV\|([^|]+)\|([^|]+)\|ref:(.*)", summary_text)
    if not m:
        raise ValueError("Could not read invoice number/date from summary.")
    number, date, ref = m.groups()

    m = re.search(r"client:(.*?)\|no:(.*?)\|ph:(.*?)(?:\|email:(.*))?$", summary_text, re.M)
    if not m:
        raise ValueError("Could not read client section from summary.")
    client_name  = m.group(1).strip()
    client_no    = m.group(2).strip()
    client_phone = m.group(3).strip()
    client_email = (m.group(4) or "").strip()

    m = re.search(r"addr:(.+)", summary_text)
    addr_line      = m.group(1).strip() if m else ""
    client_address = [p.strip() for p in addr_line.split(",") if p.strip()]

    m = re.search(r"file:(.+)", summary_text)
    base_filename = m.group(1).strip() if m else f"GG__{number}_{client_name.upper().replace(' ', '_')}"

    delivery_type   = ""
    delivery_charge = 0.0
    tax_rate        = 0.0
    m = re.search(
        r"\n([^\n|]+)\|sub:\$(-?[0-9.,]+)\|del:\$(-?[0-9.,]+)\|tax\((\d+(?:\.\d+)?)%\):\$(-?[0-9.,]+)\|total:\$(-?[0-9.,]+)",
        summary_text
    )
    if m:
        delivery_type   = m.group(1).strip()
        delivery_charge = _clean_currency(m.group(3))
        tax_rate        = float(m.group(4)) / 100.0

    payment_terms = "advance"
    low = summary_text.lower()
    if "pay:in installments" in low or "pay:installments" in low or "pay:payment in installments" in low:
        payment_terms = "installments"
    elif "pay:paid in advance" in low:
        payment_terms = "advance"

    notes = ""
    m = re.search(r"\nnotes:(.+?)(?:\ninstallments:|\nINTERNAL|\Z)", summary_text, re.S)
    if m:
        notes = m.group(1).strip()

    installments          = []
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

    item_lines    = items_part.splitlines()
    cleaned_lines = [l for l in item_lines if not l.strip().startswith("(fmt:")]
    items_part    = "\n".join(cleaned_lines).strip()

    item_pattern = re.compile(
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$(-?[0-9.,]+)(?:\|disc:([0-9.]+)%?)?\|tot:\$(-?[0-9.,]+)(?:\|del:([^\n]*))?(?:\n\s*Photo-base64:\s*(data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+))?(?=\n\[|\n[^\n]+\|sub:\$|\npay:|\nnotes:|\ninstallments:|\nINTERNAL|\Z)",
        re.S,
    )

    items = []
    for g in item_pattern.finditer(items_part):
        no, desc, unit, qty, unit_price, disc, _total, delivery, photo = g.groups()
        items.append({
            "no":          no.strip(),
            "description": " ".join(desc.strip().split()).replace("⁄", "|"),
            "delivery":    " ".join((delivery or "").strip().split()).replace("⁄", "|"),
            "qty":         int(qty),
            "unit":        unit.strip(),
            "unit_price":  _clean_currency(unit_price),
            "discount":    (float(disc) / 100.0) if disc else 0.0,
            "image":       re.sub(r"\s+", "", photo) if photo else "",
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
        item_no    = item_no_match.group(1)
        vendor_no  = ""
        vendor_name = ""
        raw_cost   = 0.0
        cost_disc  = 0.0
        cost       = 0.0
        m = re.search(r'vendor_no:\s*"([^"]*)"', line)
        if m: vendor_no = m.group(1)
        m = re.search(r'vendor:\s*"([^"]*)"', line)
        if m: vendor_name = m.group(1)
        m = re.search(r"raw_cost:\s*([0-9.]+)", line)
        if m: raw_cost = float(m.group(1))
        m = re.search(r"cost_disc:\s*([0-9.]+)%", line)
        if m: cost_disc = float(m.group(1))
        m = re.search(r"(?<![a-z_])cost:\s*([0-9.]+)", line)
        if m: cost = float(m.group(1))
        internal_map[item_no] = {
            "vendor_name": vendor_name, "vendor_no": vendor_no,
            "raw_cost": raw_cost, "cost_disc": cost_disc, "cost": cost,
        }

    for item in items:
        item.update(internal_map.get(item["no"], {}))

    return {
        "number": number.strip(), "date": date.strip(), "reference": ref.strip(),
        "client_name": client_name, "client_no": client_no,
        "client_phone": client_phone, "client_email": client_email,
        "client_address": client_address, "delivery_type": delivery_type,
        "delivery_charge": delivery_charge, "tax_rate": tax_rate,
        "payment_terms": payment_terms, "notes": notes,
        "installments": installments, "installment_split_type": installment_split_type,
        "items": items, "base_filename": base_filename,
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
    invoice   = parse_summary(summary_text)
    safe_name = invoice.pop("base_filename")
    pdf_path  = OUTPUT_DIR / f"{safe_name}.pdf"
    xlsx_path = draw_invoice(invoice, str(pdf_path))
    return pdf_path, Path(xlsx_path)


# ── Supabase persistent storage ────────────────────────────────────────────────

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
TABLE        = "gg_data"
ROW_KEY      = "main"
EMPTY_DATA   = {"clients": [], "library": [], "vendors": [], "invoices": []}


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
        url = f"{SUPABASE_URL}/rest/v1/gg_photos?key=eq.{urllib.parse.quote(key)}"
        req = urllib.request.Request(url, headers=_supabase_headers(), method="DELETE")
        urllib.request.urlopen(req, timeout=8)
    except Exception:
        pass
    return jsonify({"ok": True})


# ── Competitor Pricing ─────────────────────────────────────────────────────────

FURNITURE_RETAILERS = {
    "wayfair.com", "westelm.com", "perigold.com", "potterybarn.com",
    "crateandbarrel.com", "cb2.com", "rh.com", "restorationhardware.com",
    "serenaandlily.com", "article.com", "worldmarket.com", "target.com",
    "amazon.com", "homedepot.com", "lowes.com", "ikea.com",
    "ballarddesigns.com", "arhaus.com", "roomandboard.com", "zgallerie.com",
    "hayneedle.com", "overstock.com", "allmodern.com", "jossandmain.com",
    "anthropologie.com", "burkedecor.com", "lumens.com", "ylighting.com",
    "circafurniture.com", "furniturerow.com", "ethanallen.com",
    "livingspaces.com", "castlery.com", "tuftandneedle.com",
    "pier1.com", "zgallerie.com", "highfashionhome.com", "onekingslane.com",
    "jaysonhome.com", "lazyboy.com", "havenly.com", "interior-define.com",
}

# Major retailers for site-targeted search queries
RETAILER_SITES = (
    "site:wayfair.com OR site:westelm.com OR site:perigold.com OR "
    "site:potterybarn.com OR site:crateandbarrel.com OR site:rh.com OR "
    "site:allmodern.com OR site:jossandmain.com OR site:arhaus.com OR "
    "site:article.com OR site:serenaandlily.com OR site:cb2.com OR "
    "site:roomandboard.com OR site:ballarddesigns.com OR site:zgallerie.com OR "
    "site:livingspaces.com OR site:castlery.com OR site:ethanallen.com OR "
    "site:onekingslane.com OR site:highfashionhome.com OR site:anthropologie.com"
)


def _brave_search(query: str, brave_key: str, count: int = 10) -> list:
    url = (
        f"https://api.search.brave.com/res/v1/web/search"
        f"?q={urllib.parse.quote(query)}&count={count}&country=us"
    )
    req = urllib.request.Request(url, headers={
        "X-Subscription-Token": brave_key,
        "Accept": "application/json",
    })
    with urllib.request.urlopen(req, timeout=12) as resp:
        return json.loads(resp.read()).get("web", {}).get("results", [])


def _identify_from_image(image_b64: str, anthropic_key: str) -> dict:
    """
    Call Claude Haiku Vision once and return BOTH:
      - product_name: brand + product name if identifiable (e.g. "Four Hands Portia Outdoor Occasional Chair")
      - visual_terms: material/style/shape keywords for cross-retailer matching
                      (e.g. "woven rattan wicker curved S-shape accent chair outdoor natural")
    Returns {"product_name": str, "visual_terms": str}
    """
    if "," in image_b64:
        media_type_part, data = image_b64.split(",", 1)
        media_type = media_type_part.split(":")[1].split(";")[0] if ":" in media_type_part else "image/jpeg"
    else:
        data, media_type = image_b64, "image/jpeg"

    prompt = (
        "You are a furniture product identifier for a US interior design pricing tool.\n"
        "Analyse this furniture or home decor image and respond with ONLY a JSON object — no markdown, no explanation.\n\n"
        "JSON format:\n"
        "{\n"
        '  "product_name": "<brand and product name if you can identify it, otherwise empty string>",\n'
        '  "visual_terms": "<6-12 keywords: material, construction style, color, shape, indoor/outdoor, furniture type>"\n'
        "}\n\n"
        "Examples:\n"
        '{"product_name": "Four Hands Portia Outdoor Occasional Chair", "visual_terms": "woven rattan wicker curved S-shape accent chair outdoor natural tan"}\n'
        '{"product_name": "Arteriors Dita Side Table", "visual_terms": "hammered brass gold drum side table metal round accent"}\n'
        '{"product_name": "", "visual_terms": "upholstered linen curved sofa sectional ivory cream modern organic"}\n\n'
        "Only return the JSON object. No other text."
    )

    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 120,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": data}},
                {"type": "text", "text": prompt},
            ]
        }]
    }).encode()

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "x-api-key":         anthropic_key,
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = json.loads(resp.read())["content"][0]["text"].strip()
        # Strip markdown fences if model wraps in ```json
        raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
        try:
            parsed = json.loads(raw)
            return {
                "product_name": parsed.get("product_name", "").strip(),
                "visual_terms": parsed.get("visual_terms", "").strip(),
            }
        except json.JSONDecodeError:
            # Haiku didn't return valid JSON — treat whole response as product name
            print(f"[vision] JSON parse failed, raw='{raw}'")
            return {"product_name": raw[:120], "visual_terms": ""}


def _extract_prices(raw_results: list, seen_urls: set) -> list:
    """Parse price from snippet/title and return structured result rows."""
    out = []
    for it in raw_results:
        url = it.get("url", "")
        if not url or url in seen_urls:
            continue
        seen_urls.add(url)
        snippet = it.get("description", "") + " " + it.get("title", "")
        prices  = re.findall(r'\$([\d,]+(?:\.\d{2})?)', snippet)
        if not prices:
            continue
        try:
            price = float(prices[0].replace(",", ""))
            if price <= 0 or price > 100_000:
                continue
            domain       = urllib.parse.urlparse(url).netloc.replace("www.", "")
            is_reputable = any(r in domain for r in FURNITURE_RETAILERS)
            out.append({
                "retailer":  domain,
                "price":     price,
                "url":       url,
                "title":     it.get("title", ""),
                "reputable": is_reputable,
            })
        except Exception:
            pass
    return out


@app.post("/api/price-check")
def api_price_check():
    try:
        return _api_price_check_inner()
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "error": f"Server error: {str(e)}"}), 500

def _api_price_check_inner():
    payload      = request.get_json(silent=True) or {}
    image_b64    = payload.get("image", "")
    product_text = payload.get("product", "").strip()
    sku          = payload.get("sku", "").strip()

    BRAVE_KEY     = os.environ.get("BRAVE_API_KEY", "")
    ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

    if not BRAVE_KEY:
        return jsonify({"ok": False,
            "error": "Brave Search API not configured. Add BRAVE_API_KEY to Render environment variables."}), 503

    # ── Step 1: Image identification (Haiku, ~$0.001/call) ────────────────────
    product_name  = ""   # brand + product name (e.g. "Four Hands Portia Outdoor Occasional Chair")
    visual_terms  = ""   # material/style keywords (e.g. "woven rattan wicker curved chair outdoor")
    image_used    = False
    vision_error  = ""

    if image_b64 and ANTHROPIC_KEY:
        try:
            result       = _identify_from_image(image_b64, ANTHROPIC_KEY)
            product_name = result["product_name"]
            visual_terms = result["visual_terms"]
            image_used   = True
            print(f"[vision] name='{product_name}' visual='{visual_terms}'")
        except Exception as e:
            vision_error = str(e)
            print(f"[vision] failed: {e}")

    # If user typed a name, that always takes precedence over Haiku's name guess
    if product_text:
        product_name = product_text

    # Must have something to search with
    if not product_name and not visual_terms and not sku:
        if image_b64 and vision_error:
            return jsonify({"ok": False,
                "error": f"Could not identify product from image. Please type the product name below the image. (Detail: {vision_error})"}), 400
        if image_b64 and not ANTHROPIC_KEY:
            return jsonify({"ok": False,
                "error": "Image identification requires ANTHROPIC_API_KEY in Render environment variables."}), 503
        return jsonify({"ok": False,
            "error": "Please provide a product name, SKU, or upload an image."}), 400

    # ── Step 2: Build search queries ──────────────────────────────────────────
    # We run up to 4 searches to maximise coverage across retailers that name
    # the same product differently:
    #
    # Q1  Product name (general web)        — finds exact name matches anywhere
    # Q2  Product name + retailer sites     — finds it at major retailers by name
    # Q3  Visual terms + retailer sites     — finds RENAMED versions at major retailers
    # Q4  SKU exact match                   — direct lookup if SKU provided
    #
    queries = []

    if product_name:
        q1 = f"{product_name} price buy"
        q2 = f"{product_name} ({RETAILER_SITES})"
        queries += [q1, q2]
    elif visual_terms:
        # No name — fall back to visual-only searches
        q1 = f"{visual_terms} price buy"
        queries.append(q1)

    if visual_terms and product_name:
        # Use visual terms to catch retailers that renamed the product
        q3 = f"{visual_terms} ({RETAILER_SITES})"
        queries.append(q3)

    if sku:
        q4 = f'"{sku}" price buy furniture'
        queries.append(q4)

    primary_query = queries[0] if queries else f"{product_name or visual_terms} price"

    # ── Step 3: Run all searches, deduplicate ─────────────────────────────────
    seen_urls   = set()
    all_raw     = []
    for q in queries:
        try:
            all_raw += _brave_search(q, BRAVE_KEY, count=10)
        except Exception as e:
            print(f"[brave] query failed '{q[:60]}': {e}")

    results = _extract_prices(all_raw, seen_urls)

    if not results:
        identified = product_name or visual_terms or sku
        return jsonify({
            "ok": True, "identified_product": identified,
            "image_used": image_used, "results": [],
            "floor": None, "ceiling": None, "suggested": None,
            "query_used": primary_query,
        })

    # ── Step 4: Sort — reputable retailers first, then by price ───────────────
    results.sort(key=lambda r: (not r["reputable"], r["price"]))

    # ── Step 5: Price range & suggested retail ────────────────────────────────
    prices    = [r["price"] for r in results]
    floor_p   = round(min(prices), 2)
    ceiling_p = round(max(prices), 2)
    avg_p     = sum(prices) / len(prices)
    suggested = round(avg_p * 1.05 / 5) * 5   # 5% above avg, rounded to nearest $5

    identified_label = product_name or visual_terms or sku

    return jsonify({
        "ok":                 True,
        "identified_product": identified_label,
        "image_used":         image_used,
        "product_name":       product_name,
        "visual_terms":       visual_terms,
        "results":            results,
        "floor":              floor_p,
        "ceiling":            ceiling_p,
        "suggested":          suggested,
        "query_used":         primary_query,
    })


# ── Core routes ────────────────────────────────────────────────────────────────

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
            "ok":        True,
            "pdf_name":  pdf_path.name,
            "xlsx_name": xlsx_path.name,
            "pdf_url":   f"/download/{pdf_path.name}",
            "xlsx_url":  f"/download/{xlsx_path.name}",
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.get("/download/<path:filename>")
def download(filename: str):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
