from __future__ import annotations

import base64
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from openpyxl import load_workbook

from golden_glam_invoice_generator import draw_invoice

BASE_DIR   = Path(__file__).resolve().parent
HTML_FILE  = BASE_DIR / "GoldenGlam_InvoiceGenerator_hosted.html"
OUTPUT_DIR = BASE_DIR / "generated_invoices"
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max upload

# ── Temp image hosting via Supabase Storage ───────────────────────────────────
STORAGE_BUCKET = "temp-images"
_bucket_ensured = False


def _ensure_bucket():
    global _bucket_ensured
    if _bucket_ensured or not SUPABASE_URL or not _storage_key():
        return
    try:
        key  = _storage_key()
        body = json.dumps({"id": STORAGE_BUCKET, "name": STORAGE_BUCKET, "public": True}).encode()
        req  = urllib.request.Request(
            f"{SUPABASE_URL}/storage/v1/bucket",
            data=body,
            headers={
                "apikey":        key,
                "Authorization": f"Bearer {key}",
                "Content-Type":  "application/json",
            },
            method="POST"
        )
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        if "409" not in str(e) and "already exists" not in str(e).lower():
            print(f"[storage] bucket create warning: {e}")
    _bucket_ensured = True


def _storage_key():
    # Service key bypasses RLS — required for storage write operations
    return SUPABASE_SERVICE_KEY or SUPABASE_KEY


def _upload_temp_image(img_bytes: bytes):
    if not SUPABASE_URL or not _storage_key():
        print("[storage] no Supabase credentials available")
        return None, None
    _ensure_bucket()
    filename = f"tmp_{uuid.uuid4().hex}.jpg"
    try:
        key = _storage_key()
        req = urllib.request.Request(
            f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{filename}",
            data=img_bytes,
            headers={
                "apikey":        key,
                "Authorization": f"Bearer {key}",
                "Content-Type":  "image/jpeg",
                "Cache-Control": "no-cache",
            },
            method="POST"
        )
        urllib.request.urlopen(req, timeout=10)
        public_url = f"{SUPABASE_URL}/storage/v1/object/public/{STORAGE_BUCKET}/{filename}"
        print(f"[storage] uploaded OK: {filename}")
        return public_url, filename
    except Exception as e:
        print(f"[storage] upload FAILED: {e}")
        return None, None


def _delete_temp_image(filename: str):
    if not SUPABASE_URL or not _storage_key() or not filename:
        return
    try:
        key  = _storage_key()
        body = json.dumps({"prefixes": [filename]}).encode()
        req  = urllib.request.Request(
            f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}",
            data=body,
            headers={
                "apikey":        key,
                "Authorization": f"Bearer {key}",
                "Content-Type":  "application/json",
            },
            method="DELETE"
        )
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        print(f"[storage] delete failed: {e}")


# ── Invoice parsing ────────────────────────────────────────────────────────────

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
    base_filename = m.group(1).strip() if m else \
        f"GG__{number}_{client_name.upper().replace(' ', '_')}"

    delivery_type = ""
    delivery_charge = 0.0
    tax_rate = 0.0
    m = re.search(
        r"\n([^\n|]+)\|sub:\$(-?[0-9.,]+)\|del:\$(-?[0-9.,]+)"
        r"\|tax\((\d+(?:\.\d+)?)%\):\$(-?[0-9.,]+)\|total:\$(-?[0-9.,]+)",
        summary_text
    )
    if m:
        delivery_type   = m.group(1).strip()
        delivery_charge = _clean_currency(m.group(3))
        tax_rate        = float(m.group(4)) / 100.0

    payment_terms = "advance"
    low = summary_text.lower()
    if "pay:in installments" in low or "pay:installments" in low \
            or "pay:payment in installments" in low:
        payment_terms = "installments"
    elif "pay:paid in advance" in low:
        payment_terms = "advance"

    notes = ""
    m = re.search(r"\nnotes:(.+?)(?:\ninstallments:|\nINTERNAL|\Z)", summary_text, re.S)
    if m:
        notes = m.group(1).strip()

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

    cleaned = [l for l in items_part.splitlines() if not l.strip().startswith("(fmt:")]
    items_part = "\n".join(cleaned).strip()

    item_pattern = re.compile(
        r"\[([^\]]+)\](.*?)\|([^|\n]+)\|qty:(\d+)\|\$(-?[0-9.,]+)"
        r"(?:\|disc:([0-9.]+)%?)?\|tot:\$(-?[0-9.,]+)(?:\|del:([^\n]*))?"
        r"(?:\n\s*Photo-base64:\s*(data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+))?"
        r"(?=\n\[|\n[^\n]+\|sub:\$|\npay:|\nnotes:|\ninstallments:|\nINTERNAL|\Z)",
        re.S,
    )
    items = []
    for g in item_pattern.finditer(items_part):
        no, desc, unit, qty, unit_price, disc, _tot, delivery, photo = g.groups()
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
        im = re.match(r"\[([^\]]+)\]", line)
        if not im:
            continue
        item_no = im.group(1)
        vendor_no = vendor_name = ""
        raw_cost = cost_disc = cost = 0.0
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


def generate_from_summary(summary_text: str) -> tuple[Path, Path]:
    invoice   = parse_summary(summary_text)
    safe_name = invoice.pop("base_filename")
    pdf_path  = OUTPUT_DIR / f"{safe_name}.pdf"
    xlsx_path = draw_invoice(invoice, str(pdf_path))
    return pdf_path, Path(xlsx_path)


# ── Supabase ───────────────────────────────────────────────────────────────────

SUPABASE_URL         = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY         = os.environ.get("SUPABASE_KEY", "")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
TABLE        = "gg_data"
ROW_KEY      = "main"
EMPTY_DATA   = {"clients": [], "library": [], "vendors": [], "invoices": []}


def _sb_headers():
    return {
        "apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json", "Prefer": "return=representation",
    }


def _load_data() -> dict:
    if not SUPABASE_URL or not SUPABASE_KEY:
        return dict(EMPTY_DATA)
    try:
        url = f"{SUPABASE_URL}/rest/v1/{TABLE}?key=eq.{ROW_KEY}&select=value"
        with urllib.request.urlopen(urllib.request.Request(url, headers=_sb_headers()), timeout=8) as r:
            rows = json.loads(r.read())
            return rows[0].get("value") or dict(EMPTY_DATA) if rows else dict(EMPTY_DATA)
    except Exception as e:
        print(f"[data] load error: {e}")
    return dict(EMPTY_DATA)


def _save_data(data: dict):
    if not SUPABASE_URL or not SUPABASE_KEY:
        return
    try:
        body = json.dumps({"key": ROW_KEY, "value": data}).encode()
        hdrs = {**_sb_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"}
        req  = urllib.request.Request(f"{SUPABASE_URL}/rest/v1/{TABLE}",
                                      data=body, headers=hdrs, method="POST")
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
        with urllib.request.urlopen(urllib.request.Request(url, headers=_sb_headers()), timeout=8) as r:
            rows = json.loads(r.read())
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
        hdrs = {**_sb_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"}
        req  = urllib.request.Request(f"{SUPABASE_URL}/rest/v1/gg_photos",
                                      data=body, headers=hdrs, method="POST")
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
        urllib.request.urlopen(
            urllib.request.Request(url, headers=_sb_headers(), method="DELETE"), timeout=8)
    except Exception:
        pass
    return jsonify({"ok": True})


# ── Competitor Pricing ─────────────────────────────────────────────────────────

RENDER_URL = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")

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
    "onekingslane.com", "highfashionhome.com", "jaysonhome.com",
}


def _serpapi_get(params: dict, serpapi_key: str, timeout: int = 20) -> dict:
    params["api_key"] = serpapi_key
    url = "https://serpapi.com/search?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())


def _parse_price(price_str: str) -> float | None:
    """Convert '$1,099.00' or '1099' to float. Returns None if unparseable."""
    if not price_str:
        return None
    cleaned = re.sub(r"[^\d.]", "", str(price_str))
    try:
        val = float(cleaned)
        return val if 1 < val < 100_000 else None
    except ValueError:
        return None


def _shopping_results_to_rows(items: list) -> list:
    """Convert SerpAPI shopping_results into our standard row format."""
    rows = []
    for it in items:
        price = _parse_price(it.get("price", ""))
        if price is None:
            continue
        link   = it.get("link") or it.get("product_link") or ""
        source = it.get("source", "")
        domain = urllib.parse.urlparse(link).netloc.replace("www.", "") if link else source
        rows.append({
            "retailer":  domain or source,
            "price":     price,
            "url":       link,
            "title":     it.get("title", ""),
            "thumbnail": it.get("thumbnail", ""),
            "reputable": any(r in domain for r in FURNITURE_RETAILERS),
        })
    return rows


def _lens_results_to_rows(matches: list) -> list:
    """Convert SerpAPI Google Lens visual_matches into our standard row format."""
    rows = []
    for it in matches:
        price = _parse_price(it.get("price", {}).get("value") if isinstance(it.get("price"), dict) else it.get("price", ""))
        if price is None:
            continue
        link   = it.get("link", "")
        domain = urllib.parse.urlparse(link).netloc.replace("www.", "")
        rows.append({
            "retailer":  domain,
            "price":     price,
            "url":       link,
            "title":     it.get("title", ""),
            "thumbnail": it.get("thumbnail", ""),
            "reputable": any(r in domain for r in FURNITURE_RETAILERS),
        })
    return rows


@app.post("/api/price-check")
def api_price_check():
    try:
        return _price_check_inner()
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"ok": False, "error": f"Server error: {str(e)}"}), 500


def _price_check_inner():
    payload      = request.get_json(silent=True) or {}
    image_b64    = payload.get("image", "")
    product_text = payload.get("product", "").strip()
    sku          = payload.get("sku", "").strip()

    SERPAPI_KEY = os.environ.get("SERPAPI_KEY", "")
    if not SERPAPI_KEY:
        return jsonify({"ok": False,
            "error": "SerpAPI not configured. Add SERPAPI_KEY to Render environment variables."}), 503

    results      = []
    product_name = product_text  # what we'll show the user and use for Shopping
    lens_name    = ""            # name extracted from Lens top result
    image_used   = False

    # ── Step 1: Google Lens via Supabase Storage (external, no deadlock) ────────
    if image_b64:
        image_used = True

        if "," in image_b64:
            _, b64_data = image_b64.split(",", 1)
        else:
            b64_data = image_b64

        img_bytes = None
        try:
            img_bytes = base64.b64decode(b64_data)
        except Exception as e:
            print(f"[lens] base64 decode failed: {e}")

        image_url     = None
        temp_filename = None
        if img_bytes:
            image_url, temp_filename = _upload_temp_image(img_bytes)

        if image_url:
            try:
                lens_data = _serpapi_get({
                    "engine":  "google_lens",
                    "url":     image_url,
                    "hl":      "en",
                    "country": "us",
                }, SERPAPI_KEY, timeout=25)
                visual_matches = lens_data.get("visual_matches", [])
                if visual_matches:
                    lens_name = visual_matches[0].get("title", "")
                    print(f"[lens] top match: '{lens_name}'")
                results += _lens_results_to_rows(visual_matches)
                print(f"[lens] {len(visual_matches)} visual matches")
            except Exception as e:
                print(f"[lens] SerpAPI call failed: {e}")
            finally:
                _delete_temp_image(temp_filename)
        else:
            print("[lens] skipping Lens — no image URL available")

    # ── Step 2: Determine Shopping query ──────────────────────────────────────
    # Priority: user-typed text > lens-identified name > SKU
    shopping_query = product_text or lens_name or sku
    if not shopping_query:
        if image_used:
            # Lens ran but couldn't identify — search generically for furniture
            shopping_query = "furniture home decor"
            print("[shopping] Lens returned no name, falling back to generic search")
        else:
            return jsonify({"ok": False,
                "error": "Please provide a product name, SKU, or upload an image."}), 400

    if sku and sku not in shopping_query:
        shopping_query = f"{shopping_query} {sku}"

    product_name = shopping_query

    # ── Step 3: Google Shopping ───────────────────────────────────────────────
    try:
        shopping_data = _serpapi_get({
            "engine": "google_shopping",
            "q":      shopping_query,
            "gl":     "us",
            "hl":     "en",
            "num":    "20",
        }, SERPAPI_KEY, timeout=15)

        shopping_rows = _shopping_results_to_rows(
            shopping_data.get("shopping_results", [])
        )
        results += shopping_rows
        print(f"[shopping] query='{shopping_query}' found {len(shopping_rows)} priced results")

    except Exception as e:
        print(f"[shopping] failed: {e}")

    # ── Step 4: Deduplicate by URL ────────────────────────────────────────────
    seen, deduped = set(), []
    for r in results:
        key = r["url"] or r["title"]
        if key and key not in seen:
            seen.add(key)
            deduped.append(r)

    # Sort: reputable retailers first, then by price
    deduped.sort(key=lambda r: (not r["reputable"], r["price"]))

    # ── Step 5: Price range ───────────────────────────────────────────────────
    if not deduped:
        return jsonify({
            "ok": True, "identified_product": product_name,
            "image_used": image_used, "results": [],
            "floor": None, "ceiling": None, "suggested": None,
            "query_used": shopping_query,
        })

    prices    = [r["price"] for r in deduped]
    floor_p   = round(min(prices), 2)
    ceiling_p = round(max(prices), 2)
    avg_p     = sum(prices) / len(prices)
    suggested = round(avg_p * 1.05 / 5) * 5   # 5% above avg, nearest $5

    return jsonify({
        "ok":                 True,
        "identified_product": product_name,
        "lens_name":          lens_name,
        "image_used":         image_used,
        "results":            deduped,
        "floor":              floor_p,
        "ceiling":            ceiling_p,
        "suggested":          suggested,
        "query_used":         shopping_query,
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
            "ok": True,
            "pdf_name":  pdf_path.name,  "xlsx_name": xlsx_path.name,
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
