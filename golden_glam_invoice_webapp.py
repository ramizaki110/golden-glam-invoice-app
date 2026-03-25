from flask import Flask, request, jsonify, send_from_directory
from pathlib import Path
import tempfile
import os

from golden_glam_invoice_generator import generate_from_summary

BASE_DIR = Path(__file__).resolve().parent
HTML_FILE = "GoldenGlam_InvoiceGenerator_hosted.html"
OUTPUT_DIR = BASE_DIR / "generated_invoices"
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, HTML_FILE)


@app.post("/generate")
def generate():
    data = request.get_json(silent=True) or {}
    summary = (data.get("summary") or "").strip()

    if not summary:
        return jsonify({"ok": False, "error": "Missing invoice summary."}), 400

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            suffix=".txt",
            delete=False,
            encoding="utf-8",
            dir=OUTPUT_DIR
        ) as f:
            f.write(summary)
            temp_path = f.name

        pdf_path, excel_path = generate_from_summary(temp_path, str(OUTPUT_DIR))

        return jsonify({
            "ok": True,
            "pdf_name": os.path.basename(pdf_path),
            "xlsx_name": os.path.basename(excel_path),
            "pdf_url": f"/download/{os.path.basename(pdf_path)}",
            "xlsx_url": f"/download/{os.path.basename(excel_path)}",
        })

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


@app.get("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
