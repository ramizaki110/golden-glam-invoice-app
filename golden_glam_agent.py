#!/usr/bin/env python3
"""
Golden Glam Invoice Automation
================================
Scrapes Uttermost B2B portal -> saves invoice directly to app via API.

Usage:
  python golden_glam_agent.py --item "Riccardi Floor Lamp Blue"
  python golden_glam_agent.py
"""

import argparse, base64, datetime, json, os, random, re, sys, urllib.request, urllib.parse
from pathlib import Path

APP_URL           = "https://golden-glam-invoice-app.onrender.com"
CREDENTIALS_FILE  = Path(__file__).parent / "credentials.json"
UTTERMOST_SESSION = Path(__file__).parent / "uttermost_session.json"

SUPPLIERS = {
    "Uttermost": {},
    "Surya": {"note": "Has reCAPTCHA — log in manually first"},
}

# ── Credentials ────────────────────────────────────────────────────────────────
def load_credentials():
    if not CREDENTIALS_FILE.exists():
        CREDENTIALS_FILE.write_text(json.dumps({
            "uttermost": {"username": "YOUR_EMAIL", "password": "YOUR_PASSWORD"}
        }, indent=2))
        print("Created credentials.json — fill in your details.")
        sys.exit(1)
    return json.loads(CREDENTIALS_FILE.read_text())

# ── Helpers ────────────────────────────────────────────────────────────────────
def prompt_float(prompt):
    while True:
        v = input(prompt).strip().replace('$', '').replace(',', '')
        try: return float(v)
        except ValueError: print("  Enter a number e.g. 499")

def prompt_int(prompt, default=1):
    v = input(prompt).strip()
    try: return int(v) if v else default
    except ValueError: return default

def rand_item_no():
    return str(random.randint(100000, 999999))

# ── Item info ──────────────────────────────────────────────────────────────────
def get_item_info(item_name_arg=None):
    info = {}
    info['item_name']     = item_name_arg or input("Item name: ").strip()
    info['retail_price']  = prompt_float("Retail price (from website, $): ")
    info['selling_price'] = prompt_float("Rana's selling/discounted price ($): ")
    info['qty']           = prompt_int("Quantity [1]: ", 1)
    return info

# ── Supplier selection ─────────────────────────────────────────────────────────
def select_supplier():
    names = list(SUPPLIERS.keys()) + ["Other"]
    print("\n=== Supplier ===")
    for i, n in enumerate(names, 1):
        note = SUPPLIERS.get(n, {}).get('note', '')
        print(f"  {i}. {n}" + (f"  ⚠  {note}" if note else ""))
    while True:
        c = input("Select: ").strip()
        try:
            idx = int(c) - 1
            if 0 <= idx < len(names):
                return input("Supplier name: ").strip() if names[idx] == "Other" else names[idx]
        except ValueError: pass

# ── App API ────────────────────────────────────────────────────────────────────
def fetch_app_data():
    try:
        req = urllib.request.Request(f"{APP_URL}/api/data")
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"  Warning: Could not reach app ({e})")
        return {"clients": [], "invoices": [], "library": [], "vendors": []}

def api_post(payload):
    body = json.dumps(payload).encode()
    req = urllib.request.Request(
        f"{APP_URL}/api/data", data=body,
        headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())

def api_post_photo(key, data_url):
    body = json.dumps({"key": key, "data": data_url.replace(" ", "")}).encode()
    req = urllib.request.Request(
        f"{APP_URL}/api/photo", data=body,
        headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())

# ── Client selection ───────────────────────────────────────────────────────────
def select_client(clients):
    print("\n=== Client ===")
    if clients:
        for i, c in enumerate(clients, 1):
            print(f"  {i}. {c.get('name', '')}  #{c.get('no', '')}")
        print(f"  {len(clients)+1}. Enter manually")
        while True:
            c = input("Select: ").strip()
            try:
                idx = int(c) - 1
                if idx == len(clients): break
                if 0 <= idx < len(clients): return clients[idx]
            except ValueError: pass
    return {
        "name":  input("  Name: ").strip(),
        "no":    input("  Client No.: ").strip(),
        "phone": input("  Phone: ").strip(),
        "email": input("  Email: ").strip(),
        "addr":  input("  Address: ").strip(),
    }

# ── Invoice numbering ──────────────────────────────────────────────────────────
def next_invoice_number(invoices):
    nums = []
    for inv in invoices:
        n = str(inv.get('inv_num', '') if isinstance(inv, dict) else inv)
        m = re.search(r'SI(\d+)', n)
        if m:
            try: nums.append(int(m.group(1)))
            except ValueError: pass
    return f"SI{max(nums)+1}" if nums else "SI1000172"

# ── Uttermost session ──────────────────────────────────────────────────────────
def save_uttermost_session():
    from playwright.sync_api import sync_playwright
    print("\n[Uttermost] First-time setup — browser will open.")
    print("Log in to Uttermost, then press Enter here.")
    input("Press Enter to open browser... ")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context(viewport={"width": 1280, "height": 800})
        page = ctx.new_page()
        page.goto("https://uttermost.com/sign-in", wait_until="networkidle")
        input("\nLogged in? Press Enter to save session: ")
        ctx.storage_state(path=str(UTTERMOST_SESSION))
        browser.close()
    print("[Uttermost] Session saved ✓")

# ── Scrape Uttermost ───────────────────────────────────────────────────────────
def scrape_uttermost(item_name):
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

    if not UTTERMOST_SESSION.exists():
        save_uttermost_session()

    result = {
        'sku': '', 'cost_price': 0.0, 'image_url': '',
        'dimensions': '', 'weight': ''
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context(
            storage_state=str(UTTERMOST_SESSION),
            viewport={"width": 1280, "height": 800}
        )
        page = ctx.new_page()
        try:
            # Verify session
            page.goto("https://uttermost.com", wait_until="networkidle", timeout=30000)
            page.wait_for_timeout(2000)
            if (page.locator("text=Dashboard").count() == 0 and
                    page.locator("a[href*='dashboard']").count() == 0):
                print("[Uttermost] Session expired — log in again")
                browser.close()
                UTTERMOST_SESSION.unlink(missing_ok=True)
                save_uttermost_session()
                return scrape_uttermost(item_name)
            print("[Uttermost] Session valid ✓")

            # Search
            q = urllib.parse.quote(item_name)
            page.goto(f"https://uttermost.com/search.html?query={q}",
                      wait_until="networkidle", timeout=20000)
            page.wait_for_timeout(3000)

            # Score name words for URL matching
            name_words = [w.lower() for w in re.findall(r'[a-zA-Z]+', item_name) if len(w) > 2]

            # Parse full HTML to find all product URLs (direct /slug-SKU paths)
            # This catches URLs in JSON blobs and data attributes, not just visible links
            best_href = None
            best_score = -1
            try:
                html = page.content()
                # Match /slug-NNNNN anywhere in the HTML
                raw = re.findall(r'["\']/([\w][\w-]+-\d{4,6})["\'/? ]', html)
                for path in dict.fromkeys(raw):
                    if 'collection' in path.lower():
                        continue
                    score = sum(1 for w in name_words if w in path.lower())
                    if score > best_score:
                        best_score = score
                        best_href = '/' + path
                print(f"[Uttermost] Best URL match: {best_href} (score {best_score}/{len(name_words)})")
            except Exception as e:
                print(f"[Uttermost] HTML parse error: {e}")

            clicked = False
            if best_href and best_score >= 1:
                page.goto("https://uttermost.com" + best_href,
                          wait_until="networkidle", timeout=20000)
                page.wait_for_timeout(3000)
                print(f"[Uttermost] Product: {page.url[:80]}")
                clicked = True

            if not clicked:
                # Fallback: click first result card
                for sel in [".item-name-LPg a", ".item-root-Chs a"]:
                    try:
                        page.click(sel, timeout=5000)
                        page.wait_for_load_state("networkidle", timeout=15000)
                        page.wait_for_timeout(2000)
                        print(f"[Uttermost] Product (fallback): {page.url[:80]}")
                        clicked = True
                        break
                    except PWTimeout:
                        continue

            # ── Get page body text ────────────────────────────────────────
            try:
                body_text = page.inner_text("body")
            except Exception:
                body_text = ""

            # ── SKU from page text ────────────────────────────────────────
            # "SKU: 30751" is shown on all Uttermost product pages
            sku_m = re.search(r'SKU[:\s]+(\d{4,6})', body_text, re.IGNORECASE)
            if sku_m:
                result['sku'] = sku_m.group(1)
                print(f"[Uttermost] SKU from page: {result['sku']}")

            # Also try from URL
            if not result['sku']:
                url_m = re.search(r'-(\d{4,6})/?$', page.url)
                if url_m:
                    result['sku'] = url_m.group(1)
                    print(f"[Uttermost] SKU from URL: {result['sku']}")

            # ── Navigate to specific product URL if on collection page ────
            # Collection pages mix multiple variants' prices — specific URL has correct dealer price
            if result['sku'] and ('collection' in page.url.lower() or '?' in page.url):
                slug = re.sub(r'[^a-z0-9]+', '-', item_name.lower()).strip('-')
                specific_url = f"https://uttermost.com/{slug}-{result['sku']}"
                print(f"[Uttermost] Navigating to specific product URL...")
                try:
                    resp = page.goto(specific_url, wait_until="networkidle", timeout=20000)
                    page.wait_for_timeout(2000)
                    # Force a full reload — Uttermost is a React SPA and client-side navigation
                    # leaves old DOM content (old prices, old images) until reload clears it
                    if resp and resp.ok:
                        page.reload(wait_until="networkidle", timeout=20000)
                        page.wait_for_timeout(3000)
                        print(f"[Uttermost] Specific product loaded ✓ — {page.url[:70]}")
                        body_text = page.inner_text("body")
                    else:
                        print(f"[Uttermost] Specific URL failed (status {resp.status if resp else 'N/A'})")
                except Exception as nav_e:
                    print(f"[Uttermost] Specific URL error: {nav_e}")

            # ── Prices ────────────────────────────────────────────────────
            # Uttermost B2B layout: dealer price appears before "Suggested retail price"
            # Find "Suggested retail price" in the text, then take the last $ amount before it
            try:
                retail_pos = re.search(r'Suggested\s+retail\s+price', body_text, re.IGNORECASE)
                if retail_pos:
                    text_before = body_text[:retail_pos.start()]
                    price_hits  = re.findall(r'\$([0-9,]+(?:\.\d{2})?)', text_before)
                    if price_hits:
                        # Last dollar amount before "Suggested retail price" = dealer price
                        dealer = float(price_hits[-1].replace(',', ''))
                        if 10 < dealer < 10000:
                            result['cost_price'] = dealer
                    # Also grab the retail price from the same line
                    retail_val_m = re.search(
                        r'Suggested\s+retail\s+price\s*\$([0-9,]+(?:\.\d{2})?)',
                        body_text, re.IGNORECASE
                    )
                    retail_val = float(retail_val_m.group(1).replace(',','')) if retail_val_m else 0
                    print(f"[Uttermost] Cost: ${result['cost_price']}  Retail: ${retail_val}")
                else:
                    # Fallback: take min of all prices on page
                    prices = []
                    for el in page.query_selector_all(
                            "span[class*='price'], div[class*='price'], [class*='Price']"):
                        for pm in re.finditer(r'\$([0-9,]+(?:\.\d{2})?)',
                                              el.text_content() or ''):
                            try:
                                v = float(pm.group(1).replace(',', ''))
                                if 10 < v < 100000:
                                    prices.append(v)
                            except ValueError:
                                pass
                    if prices:
                        result['cost_price'] = sorted(set(prices))[0]
                        print(f"[Uttermost] Cost (fallback min): ${result['cost_price']}")
            except Exception as pe:
                print(f"[Uttermost] Price error: {pe}")

            # ── Dimensions and weight ─────────────────────────────────────
            try:
                m = re.search(r'Dimensions[:\s]+([^\n]+)', body_text, re.IGNORECASE)
                if m:
                    result['dimensions'] = m.group(1).strip().split('UPC')[0].strip()
                m = re.search(r'Weight[:\s]+([^\n]+)', body_text, re.IGNORECASE)
                if m:
                    result['weight'] = m.group(1).strip().split('\n')[0].strip()
                if result['dimensions']: print(f"[Uttermost] Dims: {result['dimensions']}")
                if result['weight']:     print(f"[Uttermost] Weight: {result['weight']}")
            except Exception:
                pass

            # ── Images ────────────────────────────────────────────────────
            # All Riccardi variants share the same CDN folder (32 images).
            # Filter by SKU in the URL path to get only this specific product's images.
            # Then prefer white-background shot (2nd image in Uttermost's convention).
            try:
                img_els = page.query_selector_all("img[src*='/media/catalog']")
                all_imgs = []
                for el in img_els:
                    src = el.get_attribute("src") or ""
                    if src and 'media/catalog' in src and not src.endswith('.svg'):
                        if src.startswith('/'):
                            src = "https://uttermost.com" + src
                        all_imgs.append(src)
                all_imgs = list(dict.fromkeys(all_imgs))

                # Prefer images that contain the SKU number in the URL
                sku = result.get('sku', '')
                sku_imgs = [c for c in all_imgs if sku and sku in c] if sku else []
                candidates = sku_imgs if sku_imgs else all_imgs
                print(f"[Uttermost] {len(all_imgs)} total images, "
                      f"{len(sku_imgs)} SKU-specific ({sku})")

                if candidates:
                    # Prefer image with 'white'/'global' in filename
                    chosen = next(
                        (c for c in candidates
                         if any(x in c.lower() for x in ['white','global','clean','_2_','_02_'])),
                        None
                    )
                    # Fall back to 2nd candidate (white BG convention), then 1st
                    if not chosen:
                        chosen = candidates[1] if len(candidates) > 1 else candidates[0]
                    result['image_url'] = chosen
                    print(f"[Uttermost] Image #{candidates.index(chosen)+1}/{len(candidates)} ✓")
            except Exception as e:
                print(f"[Uttermost] Image error: {e}")

        except Exception as e:
            print(f"[Uttermost] Error: {e}")
        finally:
            browser.close()
    return result

def scrape_supplier(supplier, item_name):
    if supplier.lower() == 'uttermost':
        return scrape_uttermost(item_name)
    return {
        'sku':        input(f"  [{supplier}] SKU: ").strip(),
        'cost_price': prompt_float(f"  [{supplier}] Cost price ($): "),
        'image_url':  '',
        'dimensions': input(f"  [{supplier}] Dimensions (optional): ").strip(),
        'weight':     '',
    }

# ── Download image ─────────────────────────────────────────────────────────────
def download_image(url):
    if not url:
        return ''
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as r:
            ct = r.headers.get('Content-Type', 'image/jpeg').split(';')[0]
            return f"data:{ct};base64,{base64.b64encode(r.read()).decode()}"
    except Exception as e:
        print(f"  Image download failed: {e}")
        return ''

# ── Build invoice ──────────────────────────────────────────────────────────────
def build_invoice(invoice_number, client, items_data, today_iso):
    """
    Build invoice JSON matching app's captureCurrentInvoice() format exactly.
    Confirmed from GoldenGlam_InvoiceGenerator_hosted.html source code.
    """
    line_items = []
    for it in items_data:
        item_no = rand_item_no()
        desc    = it['item_name']
        dims    = it.get('dimensions', '') or it.get('supplier_dims', '')
        if dims:
            desc = f"{desc} | {dims}"
        retail   = it.get('retail_price', it['selling_price'])
        selling  = it['selling_price']
        disc_pct = round((1 - selling / retail) * 100, 1) if retail and retail != selling else 0
        cost     = it.get('cost_price', 0)

        line_items.append({
            "no":        item_no,
            "desc":      desc,
            "vendor":    it.get('supplier', ''),
            "vno":       it.get('sku', ''),
            "qty":       str(it.get('qty', 1)),
            "price":     str(retail),
            "disc":      str(disc_pct),
            "net":       str(round(selling, 2)),
            "raw_cost":  str(cost),
            "cost_disc": "",
            "cost":      str(cost),
            "del":       "In Stock",
            "unit":      "Piece",
            "photo":     "",  # saved separately via /api/photo
        })
        it['_item_no'] = item_no

    return {
        "inv_num":        invoice_number,
        "inv_date":       today_iso,
        "inv_ref":        "",
        "client_name":    client.get('name', ''),
        "client_no":      client.get('no', ''),
        "client_phone":   client.get('phone', ''),
        "client_email":   client.get('email', ''),
        "client_addr":    client.get('addr', client.get('address', '')),
        "delivery_type":  "Standard Delivery",
        "delivery_charge":"0",
        "tax_rate":       "0",
        "payment_terms":  "advance",
        "inv_notes":      "",
        "savedAt":        datetime.datetime.now(datetime.timezone.utc).isoformat().replace("+00:00", "Z"),
        "installment_split_type": "amount",
        "installments":   [],
        "lineItems":      line_items,
    }

def save_invoice(invoice_obj, items_data, app_data):
    invoices = app_data.get('invoices', [])
    invoices.insert(0, invoice_obj)  # newest first
    result = api_post({
        "clients":  app_data.get('clients', []),
        "library":  app_data.get('library', []),
        "vendors":  app_data.get('vendors', []),
        "invoices": invoices,
    })
    if not result.get('ok'):
        return False
    # Upload photos separately
    for it, li in zip(items_data, invoice_obj['lineItems']):
        img = it.get('image_b64', '')
        if img and img.startswith('data:'):
            key = f"{invoice_obj['inv_num']}_{li['no']}"
            try:
                api_post_photo(key, img)
                print(f"  Photo saved ✓ ({key})")
            except Exception as e:
                print(f"  Photo upload failed ({e})")
    return True

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--item', '-i')
    args = parser.parse_args()

    load_credentials()
    today_iso = datetime.date.today().isoformat()

    print("\n" + "═" * 50)
    print("  ✦  Golden Glam Invoice")
    print("═" * 50)

    print("\nConnecting to app...")
    app_data = fetch_app_data()
    clients  = app_data.get('clients', [])
    invoice_number = next_invoice_number(app_data.get('invoices', []))
    print(f"Next invoice: {invoice_number}")

    client = select_client(clients)
    print(f"✓ Client: {client.get('name', '')}")

    items_data = []
    first = True
    while True:
        print(f"\n{'─' * 40}")
        print(f"Item {len(items_data)+1}")
        item_info = get_item_info(args.item if first else None)
        first = False
        supplier = select_supplier()
        item_info['supplier'] = supplier

        print(f"\n  Item:     {item_info['item_name']}")
        print(f"  Retail:   ${item_info['retail_price']}")
        print(f"  Selling:  ${item_info['selling_price']}")
        print(f"  Qty:      {item_info.get('qty', 1)}")
        print(f"  Supplier: {supplier}")

        if input("\nScrape supplier for cost/SKU/image? (y/n): ").strip().lower() == 'y':
            s = scrape_supplier(supplier, item_info['item_name'])

            sku_ovr  = input(f"  SKU [{s.get('sku', '')}] (Enter to keep): ").strip()
            cost_ovr = input(f"  Cost [${s.get('cost_price', 0)}] (Enter to keep): ").strip()
            if sku_ovr:  s['sku'] = sku_ovr
            if cost_ovr: s['cost_price'] = float(cost_ovr.replace('$', ''))

            item_info.update({
                'sku':          s.get('sku', ''),
                'cost_price':   s.get('cost_price', 0),
                'dimensions':   s.get('dimensions', ''),
                'supplier_dims':s.get('dimensions', ''),
                'weight':       s.get('weight', ''),
            })
            if s.get('image_url'):
                print("  Downloading image...")
                item_info['image_b64'] = download_image(s['image_url'])
        else:
            item_info.update({
                'sku':          input("  SKU: ").strip(),
                'cost_price':   prompt_float("  Cost price ($): "),
                'image_b64':    '',
                'dimensions':   input("  Dimensions (optional): ").strip(),
                'supplier_dims':'',
            })

        items_data.append(item_info)

        if input("\nAdd another item? (y/n): ").strip().lower() != 'y':
            break

    print(f"\n{'─' * 50}")
    print(f"Invoice: {invoice_number}  |  Client: {client.get('name', '')}")
    for it in items_data:
        print(f"  • {it['item_name']} x{it.get('qty', 1)} "
              f"@ ${it['selling_price']} (cost ${it.get('cost_price', 0)}) [{it.get('supplier', '')}]")
    total = sum(it['selling_price'] * it.get('qty', 1) for it in items_data)
    print(f"  Total: ${total:.2f}")
    print(f"{'─' * 50}")

    if input("\nSave invoice to app? (y/n): ").strip().lower() != 'y':
        print("Aborted.")
        sys.exit(0)

    print("Saving...")
    invoice_obj = build_invoice(invoice_number, client, items_data, today_iso)
    try:
        ok = save_invoice(invoice_obj, items_data, app_data)
        if ok:
            print(f"\n✅  Invoice {invoice_number} saved!")
            print(f"    Open {APP_URL} → Invoice Library")
        else:
            raise Exception("API returned ok=false")
    except Exception as e:
        backup = Path(f"invoice_{invoice_number}_backup.json")
        backup.write_text(json.dumps(invoice_obj, indent=2))
        print(f"\n❌  Save failed ({e})\n    Backed up to {backup.name}")

if __name__ == "__main__":
    main()
