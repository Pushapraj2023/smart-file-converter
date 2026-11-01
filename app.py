import os
import sys
import sqlite3
import tempfile
import traceback
from datetime import date
from pathlib import Path
from threading import Lock

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename

from web_file_convertor import DocumentConverter

app = Flask(__name__)
CORS(app)

APP_DIR = Path(__file__).resolve().parent
TEMP_DIR = APP_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
DB_PATH = APP_DIR / "usage.db"
db_lock = Lock()

def get_client_ip(req):
    forwarded = req.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return req.remote_addr or "unknown"

def check_and_update_quota(ip, requested_pages=1, limit=5):
    today = date.today().isoformat()
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usage (
                ip TEXT,
                usage_date TEXT,
                count INTEGER,
                PRIMARY KEY (ip, usage_date)
            )
        """)
        cursor.execute("SELECT count FROM usage WHERE ip = ? AND usage_date = ?", (ip, today))
        row = cursor.fetchone()
        current_count = row[0] if row else 0

        if current_count + requested_pages > limit:
            conn.close()
            return False, current_count, limit

        new_count = current_count + requested_pages
        cursor.execute("""
            INSERT INTO usage (ip, usage_date, count)
            VALUES (?, ?, ?)
            ON CONFLICT(ip, usage_date) DO UPDATE SET count = excluded.count
        """, (ip, today, new_count))
        conn.commit()
        conn.close()
        return True, new_count, limit

@app.route("/", methods=["GET"])
def health_check():
    return jsonify({"status": "online", "service": "Smart Document Converter API"})

@app.route("/convert", methods=["POST"])
def convert():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded in request."}), 400

        file = request.files["file"]
        if not file or file.filename == "":
            return jsonify({"error": "Selected file has an empty filename."}), 400

        output_format = request.form.get("output_format", "word").lower().strip()
        page_selection = request.form.get("page_selection", "all").strip()

        client_ip = get_client_ip(request)
        allowed, current, limit = check_and_update_quota(client_ip, requested_pages=1, limit=5)
        if not allowed:
            return jsonify({
                "error": f"Daily conversion quota exceeded ({limit} pages/day). Used: {current} pages."
            }), 429

        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            return jsonify({"error": "GEMINI_API_KEY environment variable is not configured on server."}), 500

        filename = secure_filename(file.filename)
        input_path = TEMP_DIR / filename
        file.save(str(input_path))

        converter = DocumentConverter(api_key=api_key)

        if output_format == "excel":
            out_path = converter.convert_to_excel(str(input_path), page_selection=page_selection)
            mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            out_path = converter.convert_to_word(str(input_path), page_selection=page_selection)
            mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        return send_file(
            out_path,
            as_attachment=True,
            download_name=Path(out_path).name,
            mimetype=mimetype
        )

    except Exception as e:
        print("=" * 60, file=sys.stderr)
        print("CRITICAL BACKEND ERROR IN CONVERT ROUTE:", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        print("=" * 60, file=sys.stderr)
        return jsonify({"error": f"Internal conversion error: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
