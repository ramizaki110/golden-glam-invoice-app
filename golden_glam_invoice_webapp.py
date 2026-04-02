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
        # 400 or 409 both mean bucket already exists — that's fine
        err = str(e)
        if "400" not in err and "409" not in err and "already exists" not in err.lower():
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
                "cache-control": "3600",
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

# Reputable retailers — both domains and display names used in Shopping results
FURNITURE_RETAILERS = {
    # Domains
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
    "luluandgeorgia.com", "topmodern.com", "luxedecor.com", "2modern.com",
    "abccarpet.com", "stashhomefurniture.com", "graysonliving.com",
    "mcgeeandco.com", "shopatironwood.com", "lauradesignco.com",
    "laylagrayce.com", "cityhome.com", "fineline.com",
    # Display names (as returned by Google Shopping source field)
    "wayfair", "west elm", "pottery barn", "crate and barrel", "cb2",
    "restoration hardware", "rh", "serena & lily", "serena and lily",
    "article", "world market", "ballard designs", "arhaus",
    "room & board", "room and board", "z gallerie", "all modern",
    "joss & main", "joss and main", "anthropologie", "burke decor",
    "lumens", "one kings lane", "high fashion home", "jayson home",
    "lulu and georgia", "topmodern", "luxedecor", "2modern",
    "abc carpet & home", "abc carpet and home", "stash home furniture",
    "grayson living", "mcgee & co", "mcgee and co", "layla grayce",
    "city home", "fine line furniture", "perigold", "perigold by wayfair", "hayneedle",
    "overstock", "target", "ikea", "amazon",
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
        link      = it.get("link") or it.get("product_link") or ""
        source    = it.get("source", "")  # actual retailer name e.g. "Wayfair"
        # Use source name directly — Shopping links often go via Google redirects
        domain    = urllib.parse.urlparse(link).netloc.replace("www.", "") if link else ""
        retailer  = source or domain  # prefer source field
        # Check reputable against both display name and domain
        src_lower = source.lower()
        dom_lower = domain.lower()
        is_rep    = any(r in src_lower or r in dom_lower for r in FURNITURE_RETAILERS)
        rows.append({
            "retailer":    retailer,
            "price":       price,
            "url":         link,
            "title":       it.get("title", ""),
            "thumbnail":   it.get("thumbnail", ""),
            "reputable":   is_rep,
            "source_type": "shopping",
        })
    return rows


def _lens_results_to_rows(matches: list) -> list:
    """Convert SerpAPI Google Lens visual_matches into our standard row format."""
    rows = []
    for it in matches:
        price_raw = it.get("price")
        if isinstance(price_raw, dict):
            price = _parse_price(price_raw.get("value", ""))
        else:
            price = _parse_price(price_raw or "")
        if price is None:
            continue
        link     = it.get("link", "")
        domain   = urllib.parse.urlparse(link).netloc.replace("www.", "")
        source   = it.get("source", domain)
        rows.append({
            "retailer":    source or domain,
            "price":       price,
            "url":         link,
            "title":       it.get("title", ""),
            "thumbnail":   it.get("thumbnail", ""),
            "reputable":   any(r in domain for r in FURNITURE_RETAILERS),
            "source_type": "lens",
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
                    # Consensus product name: find the most frequent significant word
                    # across top 10 Lens matches — much more reliable than top-1
                    NOISE = {"the","and","for","with","from","this","that",
                             "chair","table","sofa","lamp","desk","bench",
                             "outdoor","indoor","patio","set","piece","home",
                             "decor","collection","series","style","wood",
                             "furniture","upholstered","dining","accent","side"}
                    from collections import Counter
                    word_counts = Counter()
                    for vm in visual_matches[:10]:
                        t = vm.get("title","").lower()
                        for w in t.split():
                            w = w.strip(".,()-/")
                            if len(w) > 3 and w not in NOISE:
                                word_counts[w] += 1

                    # Top word that appears in multiple matches = most reliable identifier
                def _clean_title(t):
                    import re as _r
                    for pat in [
                        r'\s*[|\-]\s*(wayfair|pottery barn|west elm|perigold|crate.*barrel|restoration hardware|overstock|amazon|target|walmart|houzz).*$',
                        r"\s*(you'?ll love|shop now|best sellers|on sale|free shipping).*$",
                    ]:
                        t = _r.sub(pat, '', t, flags=_r.IGNORECASE).strip()
                    return t
                    consensus_words = [w for w,c in word_counts.most_common(5) if c >= 2]
                    if consensus_words:
                        # Build query from consensus words + top match title for context
                        lens_name = _clean_title(visual_matches[0].get("title",""))
                        if consensus_words[0] not in lens_name.lower():
                            for vm in visual_matches[:5]:
                                c = _clean_title(vm.get("title",""))
                                if consensus_words[0] in c.lower():
                                    lens_name = c
                                    break
                    else:
                        lens_name = _clean_title(visual_matches[0].get("title",""))

                    # Add color if detectable from matches
                    COLOR_WORDS = ["natural","white","black","brown","grey","gray",
                                   "beige","cream","tan","teak","vintage","walnut",
                                   "ivory","charcoal","espresso","rattan","antique"]
                    detected_color = next(
                        (cw for vm in visual_matches[:5]
                         for cw in COLOR_WORDS
                         if cw in vm.get("title","").lower()
                         and cw not in lens_name.lower()),
                        ""
                    )
                    if detected_color:
                        lens_name = f"{lens_name} {detected_color}"

                    print(f"[lens] consensus={consensus_words[:3]} final_name='{lens_name}'")
                results += _lens_results_to_rows(visual_matches)
                print(f"[lens] {len(visual_matches)} visual matches")
            except Exception as e:
                print(f"[lens] SerpAPI call failed: {e}")
            finally:
                _delete_temp_image(temp_filename)
        else:
            print("[lens] skipping Lens — no image URL available")

    # ── Step 2: Determine Shopping query ──────────────────────────────────────
    # Priority: typed product name > SKU > lens-identified name
    # Typed text is ALWAYS authoritative — never override with Lens
    if product_text and sku:
        shopping_query = f"{product_text} {sku}"
    elif product_text:
        shopping_query = product_text
    elif sku:
        shopping_query = sku
    elif lens_name:
        shopping_query = lens_name
    elif image_used:
        # Lens ran but got nothing useful — don't search garbage terms
        return jsonify({
            "ok":    False,
            "error": "Google Lens could not identify this product. Please type the product name or brand in the field above and search again.",
            "lens_failed": True,
        }), 400
    else:
        return jsonify({"ok": False,
            "error": "Please provide a product name, SKU, or upload an image."}), 400

    product_name = shopping_query
    print(f"[shopping] query: '{shopping_query}' (typed={bool(product_text)}, sku={bool(sku)}, lens={bool(lens_name)})")

    # ── Step 3: Google Shopping — two passes ────────────────────────────────
    # Pass 1: General search
    try:
        shopping_data = _serpapi_get({
            "engine": "google_shopping",
            "q":      shopping_query,
            "gl":     "us",
            "hl":     "en",
            "num":    "20",
        }, SERPAPI_KEY, timeout=15)
        shopping_rows = _shopping_results_to_rows(shopping_data.get("shopping_results", []))
        results += shopping_rows
        print(f"[shopping] pass1 found {len(shopping_rows)} results")
    except Exception as e:
        print(f"[shopping] pass1 failed: {e}")

    # Pass 2: Targeted at reputable furniture retailers to surface Perigold,
    # West Elm, Pottery Barn etc. that may not rank organically
    # Pass 2 & 3: Targeted searches at top reputable retailers
    PREMIUM_SITES  = "site:perigold.com OR site:westelm.com OR site:rh.com OR site:potterybarn.com OR site:crateandbarrel.com OR site:cb2.com OR site:serenaandlily.com OR site:arhaus.com"
    BOUTIQUE_SITES = "site:topmodern.com OR site:2modern.com OR site:luxedecor.com OR site:luluandgeorgia.com OR site:onekingslane.com OR site:laylagrayce.com OR site:roomandboard.com OR site:burkedecor.com OR site:highfashionhome.com"

    for sites, label in [(PREMIUM_SITES, "premium"), (BOUTIQUE_SITES, "boutique")]:
        try:
            td = _serpapi_get({
                "engine": "google_shopping",
                "q":      f"{shopping_query} ({sites})",
                "gl":     "us",
                "hl":     "en",
                "num":    "10",
            }, SERPAPI_KEY, timeout=15)
            rows = _shopping_results_to_rows(td.get("shopping_results", []))
            results += rows
            print(f"[shopping] pass-{label} found {len(rows)} results")
        except Exception as e:
            print(f"[shopping] pass-{label} failed: {e}")

    # ── Step 4: Deduplicate by URL ────────────────────────────────────────────
    seen, deduped = set(), []
    for r in results:
        key = r["url"] or r["title"]
        if key and key not in seen:
            seen.add(key)
            deduped.append(r)

    # ── Step 4b: Filter Shopping by relevance to identified product ──────────
    STOP_WORDS = {"with","from","that","this","and","for","the",
                  "home","decor","set","piece","modern","collection"}
    if lens_name or product_text:
        ref_name  = (lens_name or product_text).lower()
        key_words = [w.strip(".,()-") for w in ref_name.split()
                     if len(w) > 3 and w.strip(".,()-") not in STOP_WORDS]
        if key_words:
            primary_kw = key_words[0]   # most specific word (e.g. "portia")
            filtered   = []
            for r in deduped:
                title_lower = r.get("title","").lower()
                if r.get("source_type") == "lens":
                    filtered.append(r)   # always keep Lens visual matches
                elif primary_kw in title_lower:
                    filtered.append(r)   # exact primary word match (best)
                elif any(kw in title_lower for kw in key_words[1:4]):
                    filtered.append(r)   # any secondary keyword match (fallback)
            # Only apply filter if it keeps enough results
            deduped = filtered if len(filtered) >= 4 else deduped

    # ── Step 4c: Deduplicate by retailer (keep lowest price per retailer) ────
    seen_ret = {}
    for r in deduped:
        key = r.get("retailer","").lower().strip()
        if key not in seen_ret or r["price"] < seen_ret[key]["price"]:
            seen_ret[key] = r
    deduped = list(seen_ret.values())

    # ── Step 4d: Sort — Lens first, then reputable, then price ───────────────
    # Remove Walmart — not a reputable furniture source for interior design
    deduped = [r for r in deduped if "walmart" not in r.get("retailer","").lower()]
    # Sort: Lens matches first, reputable retailers next, then by price
    deduped.sort(key=lambda r: (r.get("source_type")!="lens", not r.get("reputable",False), r["price"]))

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


# ── Delivery Estimate ─────────────────────────────────────────────────────────


def _fetch_page_shipping(url: str, zip_code: str, timeout: int = 8) -> str:
    # Fetch a product page and look for shipping/delivery text
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml",
        })
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            try:
                text = raw.decode("utf-8")
            except Exception:
                text = raw.decode("latin-1", errors="ignore")
        # Strip HTML tags
        import re
        plain = re.sub(r"<[^>]+>", " ", text)
        plain = re.sub(r"\s+", " ", plain)
        plain_lower = plain.lower()
        # Look for shipping/delivery snippets
        patterns = [
            r"free\s+(?:standard\s+)?(?:shipping|delivery)[^.]{0,80}",
            r"(?:white\s+glove|threshold|room\s+of\s+choice)[^.]{0,100}",
            r"(?:shipping|delivery)[^.]{0,60}?\$[\d,]+(?:\.\d{2})?",
            r"estimated\s+(?:shipping|delivery)[^.]{0,80}",
            r"ships?\s+(?:in|within|to)[^.]{0,60}",
        ]
        for pat in patterns:
            m = re.search(pat, plain_lower)
            if m:
                snippet = plain[m.start():m.end()].strip()
                if 8 < len(snippet) < 160:
                    return snippet.capitalize()
    except Exception as e:
        print(f"[delivery] fetch failed for {url[:50]}: {e}")
    return ""


@app.post("/api/delivery-estimate")
def api_delivery_estimate():
    try:
        return _delivery_estimate_inner()
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


def _delivery_estimate_inner():
    payload       = request.get_json(silent=True) or {}
    product_name  = payload.get("product", "").strip()
    address       = payload.get("address", "").strip()
    zip_code      = payload.get("zip", "30030").strip()
    retailer_urls = payload.get("urls", [])

    SERPAPI_KEY = os.environ.get("SERPAPI_KEY", "")
    if not SERPAPI_KEY:
        return jsonify({"ok": False, "error": "SerpAPI not configured"}), 503
    if not product_name:
        return jsonify({"ok": False, "error": "No product name provided"}), 400

    delivery_addr = address or f"zip {zip_code}"
    import re as _re
    city_state_m  = _re.search(r'([A-Za-z ]+),\s*([A-Z]{2})\s+\d{5}', delivery_addr)
    city_state    = f"{city_state_m.group(1).strip()}, {city_state_m.group(2)}" if city_state_m else zip_code

    results = []
    seen    = set()

    # Extract retailer names from the URLs passed from frontend
    retailer_names = []
    for url in retailer_urls[:8]:
        if url and "google.com" not in url:
            domain = urllib.parse.urlparse(url).netloc.replace("www.", "")
            # Convert domain to readable name
            name = domain.split(".")[0].replace("-", " ").title()
            retailer_names.append((name, domain, url))

    # Search per-retailer delivery policy — more accurate than generic product search
    for ret_name, domain, url in retailer_names[:5]:
        if domain.lower() in seen:
            continue
        query = f"{ret_name} furniture white glove delivery fee {zip_code}"
        try:
            data = _serpapi_get({
                "engine": "google",
                "q":      query,
                "gl":     "us",
                "hl":     "en",
                "num":    "3",
            }, SERPAPI_KEY, timeout=10)
            # Look for delivery cost in organic results
            shipping_info = ""
            for r in data.get("organic_results", []):
                snippet = r.get("snippet", "")
                snip_low = snippet.lower()
                if any(w in snip_low for w in ["deliver","shipping","freight","white glove"]):
                    # Extract the relevant part
                    import re as _re2
                    m = _re2.search(r"[^.]*(?:deliver|ship|freight|white glove)[^.]{0,120}", snip_low)
                    if m:
                        shipping_info = snippet[m.start():m.end()].strip().capitalize()
                        break
            if not shipping_info:
                shipping_info = "Contact retailer for delivery quote"
            results.append({
                "retailer": ret_name,
                "shipping": shipping_info,
                "url":      url,
                "reliable": "contact" not in shipping_info.lower(),
            })
            seen.add(domain.lower())
        except Exception as e:
            print(f"[delivery] {ret_name} query failed: {e}")

    # If no retailer URLs, do a general furniture delivery cost search
    if not results:
        try:
            data = _serpapi_get({
                "engine": "google_shopping",
                "q":      f"{product_name} white glove delivery {city_state}",
                "gl":     "us",
                "hl":     "en",
                "num":    "10",
            }, SERPAPI_KEY, timeout=15)
            for it in data.get("shopping_results", []):
                source   = it.get("source", "")
                if source.lower() in seen or "walmart" in source.lower():
                    continue
                delivery = it.get("delivery") or ""
                for ext in (it.get("extensions") or []):
                    if any(w in ext.lower() for w in ["white glove","threshold","freight","room of choice"]):
                        delivery = ext
                        break
                if not delivery:
                    continue
                results.append({
                    "retailer": source,
                    "shipping": delivery,
                    "url":      it.get("link",""),
                    "reliable": any(w in delivery.lower() for w in ["white glove","threshold","freight"]),
                })
                seen.add(source.lower())
        except Exception as e:
            print(f"[delivery] fallback failed: {e}")

    return jsonify({
        "ok":      True,
        "results": results[:8],
        "address": delivery_addr,
        "zip":     zip_code,
        "note":    "Furniture delivery to a residential address typically ranges $99–$399 (threshold) to $199–$499 (white glove with assembly). Verify exact cost with each retailer.",
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
