import os
import sqlite3
import tempfile
from datetime import date
from pathlib import Path
from threading import Lock

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from web_file_convertor import DocumentConverter, HAS_PDFIUM, pdfium
import traceback
import sys

@app.route('/convert', methods=['POST'])
def convert():
    try:
        # Aapka existing conversion logic yahan hai
        ...
    except Exception as e:
        print("CRITICAL ERROR IN CONVERT ROUTE:", file=sys.stderr)
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# pdf2image fallback import check
try:
    from pdf2image import pdfinfo_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

APP_DIR = Path(__file__).resolve().parent
TEMP_DIR = APP_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
DB_PATH = Path(os.environ.get("QUOTA_DB_PATH", str(APP_DIR / "quota.sqlite3")))
ALLOWED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"}
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "20"))
DAILY_PAGE_LIMIT = int(os.environ.get("DAILY_PAGE_LIMIT", "5"))

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "smart-office-converter-secret")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# Blogger (mindsofaking.in) se request allow karne ke liye CORS
CORS(app, resources={r"/*": {"origins": "*"}})

converter = DocumentConverter()
db_lock = Lock()

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS daily_usage (
            user_key TEXT NOT NULL, usage_date TEXT NOT NULL, pages INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(user_key, usage_date))""")
init_db()

def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS

def count_pages(path: Path) -> int:
    ext = path.suffix.lower()
    if ext != ".pdf":
        return 1
    if HAS_PDFIUM:
        pdf = pdfium.PdfDocument(str(path))
        try:
            return len(pdf)
        finally:
            pdf.close()
    if HAS_PDF2IMAGE:
        info = pdfinfo_from_path(str(path))
        return int(info.get("Pages", 0))
    raise RuntimeError("PDF support library is not installed.")

def parse_selection(selection, total):
    selected = converter.parse_page_range(selection, total)
    if not selected:
        raise ValueError("Page selection valid nahi hai. Example: all, 1, 1-3, 1,3.")
    return selected

def reserve_usage(user_key, pages):
    today = date.today().isoformat()
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT pages FROM daily_usage WHERE user_key=? AND usage_date=?", (user_key, today)).fetchone()
        used = int(row[0]) if row else 0
        if used + pages > DAILY_PAGE_LIMIT:
            return False, used
        conn.execute(
            "INSERT INTO daily_usage(user_key, usage_date, pages) VALUES(?,?,?) "
            "ON CONFLICT(user_key, usage_date) DO UPDATE SET pages=excluded.pages",
            (user_key, today, used + pages)
        )
        return True, used

def refund_usage(user_key, pages):
    today = date.today().isoformat()
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "UPDATE daily_usage SET pages=MAX(0,pages-?) WHERE user_key=? AND usage_date=?",
            (pages, user_key, today)
        )

@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "online",
        "service": "Devlys 010 & Arial Converter Engine",
        "daily_limit_per_user": DAILY_PAGE_LIMIT
    })

@app.route("/convert", methods=["POST"])
def convert():
    upload = request.files.get("file")
    output_format = request.form.get("output_format", "word").lower()
    page_selection = request.form.get("page_selection", "all").strip() or "all"

    if not upload or not upload.filename:
        return jsonify({"error": "Pehle PDF ya image file select karein."}), 400

    filename = secure_filename(upload.filename)
    ext = Path(filename).suffix.lower()
    if not filename or not allowed_file(filename):
        return jsonify({"error": "Sirf PDF, PNG, JPG, JPEG, BMP, TIFF aur WEBP files allowed hain."}), 400

    if output_format not in {"word", "excel"}:
        return jsonify({"error": "Output format 'word' ya 'excel' hi hona chahiye."}), 400

    # User IP detection (Reverse proxy ke piche Render/Cloudflare headers support)
    user_key = request.headers.get("X-Forwarded-For", request.remote_addr or "unknown").split(",")[0].strip()
    
    input_path = None
    output_path = None
    reserved_pages = 0
    quota_reserved = False

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext, dir=TEMP_DIR) as tmp:
            input_path = Path(tmp.name)
            upload.save(tmp)

        total = count_pages(input_path)
        selected_pages = parse_selection(page_selection, total)
        reserved_pages = len(selected_pages)

        if reserved_pages > DAILY_PAGE_LIMIT:
            raise ValueError(f"Ek baar mein adhiktam {DAILY_PAGE_LIMIT} pages select kar sakte hain.")

        ok, used = reserve_usage(user_key, reserved_pages)
        if not ok:
            remaining = max(0, DAILY_PAGE_LIMIT - used)
            reserved_pages = 0
            raise ValueError(f"Aaj ki daily free limit {DAILY_PAGE_LIMIT} pages khatam ho chuki hai. Aapne {used} pages use kiye hain.")

        quota_reserved = True
        out_ext = ".docx" if output_format == "word" else ".xlsx"
        output_path = TEMP_DIR / (Path(filename).stem + "_converted" + out_ext)

        if output_format == "word":
            converter.convert_to_word(str(input_path), str(output_path), page_selection=page_selection)
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            converter.convert_to_excel(str(input_path), str(output_path), page_selection=page_selection)
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        return send_file(output_path, as_attachment=True, download_name=output_path.name, mimetype=mime)

    except Exception as exc:
        if quota_reserved and reserved_pages:
            refund_usage(user_key, reserved_pages)
        return jsonify({"error": str(exc)}), 500

    finally:
        if input_path and input_path.exists():
            try:
                input_path.unlink()
            except OSError:
                pass

@app.errorhandler(413)
def too_large(_error):
    return jsonify({"error": f"File size {MAX_UPLOAD_MB} MB se kam honi chahiye."}), 413

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
