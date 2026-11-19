import io
import os
import uuid
import re
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from doc_generator import (
    generate_kheralu_order_docx,
    generate_batch_docx,
    generate_kheralu_order_docx_bytes,
    generate_batch_docx_bytes
)
from pdf_generator import (
    generate_pdf_document,
    merge_pdf_files,
    generate_pdf_bytes,
    generate_batch_pdf_bytes
)
from generate_template import create_base_template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Vercel serverless functions have a read-only filesystem except for /tmp
if os.environ.get("VERCEL"):
    DOCS_DIR = "/tmp/generated_docs"
else:
    DOCS_DIR = os.path.join(BASE_DIR, "generated_docs")

TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, "templates"))

os.makedirs(DOCS_DIR, exist_ok=True)

if not os.path.exists(TEMPLATE_PATH):
    try:
        create_base_template(TEMPLATE_PATH)
    except Exception as e:
        print(f"Base template notice: {e}")

@app.after_request
def add_expose_headers(response):
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return response

def sanitize_filename_part(text):
    return re.sub(r'[/\\:\*\?"<>\| ]', '_', str(text).strip())

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/generate", methods=["POST"])
@app.route("/api/generate_pdf", methods=["POST"])
@app.route("/api/generate_docx", methods=["POST"])
def generate_single_doc():
    """Generates single 2-page document and streams it directly in memory."""
    try:
        data = request.get_json(silent=True) or {}
        if not data:
            return jsonify({"status": "error", "message": "અમાન્ય ડેટા પ્રદાન કરવામાં આવ્યો છે."}), 400

        # Determine requested format from path, query parameter, or JSON payload
        fmt = (request.args.get("format") or data.get("format") or ("docx" if request.path.endswith("_docx") else "pdf")).lower()

        reg_no = sanitize_filename_part(data.get("reg_no", "Order"))
        today_date = sanitize_filename_part(data.get("today_date", "date"))
        base_name = f"Sudhara_Hukam_{reg_no}_{today_date}"

        if fmt == "docx":
            stream = generate_kheralu_order_docx_bytes(data)
            filename = f"{base_name}.docx"
            return send_file(
                stream,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            stream = generate_pdf_bytes(data)
            filename = f"{base_name}.pdf"
            return send_file(
                stream,
                as_attachment=True,
                download_name=filename,
                mimetype="application/pdf"
            )
    except Exception as e:
        print(f"Error generating single document: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/generate_batch", methods=["POST"])
@app.route("/api/generate_batch_pdf", methods=["POST"])
@app.route("/api/generate_batch_docx", methods=["POST"])
def generate_batch_docs():
    """Generates a combined batch document and streams it directly in memory."""
    try:
        payload = request.get_json(silent=True) or {}
        if not payload or "batch" not in payload or not isinstance(payload["batch"], list) or len(payload["batch"]) == 0:
            return jsonify({"status": "error", "message": "બેચ ક્વિમાં કોઈ અરજીઓ નથી."}), 400

        batch_list = payload["batch"]
        batch_count = len(batch_list)
        unique_id = str(uuid.uuid4())[:6]

        fmt = (request.args.get("format") or payload.get("format") or ("docx" if request.path.endswith("_docx") else "pdf")).lower()
        base_name = f"Combined_Sudhara_Hukam_Batch_{batch_count}_{unique_id}"

        if fmt == "docx":
            stream = generate_batch_docx_bytes(batch_list)
            filename = f"{base_name}.docx"
            return send_file(
                stream,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            stream = generate_batch_pdf_bytes(batch_list)
            filename = f"{base_name}.pdf"
            return send_file(
                stream,
                as_attachment=True,
                download_name=filename,
                mimetype="application/pdf"
            )
    except Exception as e:
        print(f"Error generating batch document: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/download/<filename>")
def download_file(filename):
    """Legacy route fallback."""
    filepath = os.path.join(DOCS_DIR, filename)
    if os.path.exists(filepath):
        return send_from_directory(DOCS_DIR, filename, as_attachment=True)
    return jsonify({"status": "error", "message": "File not found"}), 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
