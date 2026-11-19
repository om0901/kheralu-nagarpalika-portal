import os
import uuid
import re
from flask import Flask, render_template, request, jsonify, send_from_directory
from doc_generator import generate_kheralu_order_docx, generate_batch_docx
from pdf_generator import generate_pdf_document, merge_pdf_files
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

def sanitize_filename_part(text):
    return re.sub(r'[/\\:\*\?"<>\| ]', '_', str(text).strip())

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/generate", methods=["POST"])
def generate_single_doc():
    """Generates single 2-page PDF and Word document."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "અમાન્ય ડેટા પ્રદાન કરવામાં આવ્યો છે."}), 400

        reg_no = sanitize_filename_part(data.get("reg_no", "Order"))
        today_date = sanitize_filename_part(data.get("today_date", "date"))
        
        base_name = f"Sudhara_Hukam_{reg_no}_{today_date}"
        docx_filename = f"{base_name}.docx"
        pdf_filename = f"{base_name}.pdf"

        docx_filepath = os.path.join(DOCS_DIR, docx_filename)
        pdf_filepath = os.path.join(DOCS_DIR, pdf_filename)

        # 1. Generate Word Document (.docx)
        generate_kheralu_order_docx(data, docx_filepath)

        # 2. Generate PDF Document (.pdf)
        pdf_generated = False
        try:
            generate_pdf_document(data, docx_filepath, pdf_filepath)
            pdf_generated = os.path.exists(pdf_filepath)
        except Exception as pdf_err:
            print(f"Warning: PDF generation fallback error: {pdf_err}")

        return jsonify({
            "status": "success",
            "docx_filename": docx_filename,
            "docx_download_url": f"/download/{docx_filename}",
            "pdf_filename": pdf_filename if pdf_generated else None,
            "pdf_download_url": f"/download/{pdf_filename}" if pdf_generated else None
        })
    except Exception as e:
        print(f"Error generating single document: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/generate_batch", methods=["POST"])
def generate_batch_docs():
    """Generates a SINGLE COMBINED PDF and Word file containing all applications in the batch queue."""
    try:
        payload = request.get_json()
        if not payload or "batch" not in payload or not isinstance(payload["batch"], list) or len(payload["batch"]) == 0:
            return jsonify({"status": "error", "message": "બેચ ક્વિમાં કોઈ અરજીઓ નથી."}), 400

        batch_list = payload["batch"]
        batch_count = len(batch_list)
        unique_id = str(uuid.uuid4())[:6]

        combined_base_name = f"Combined_Sudhara_Hukam_Batch_{batch_count}_{unique_id}"
        combined_pdf_filename = f"{combined_base_name}.pdf"
        combined_docx_filename = f"{combined_base_name}.docx"

        combined_pdf_filepath = os.path.join(DOCS_DIR, combined_pdf_filename)
        combined_docx_filepath = os.path.join(DOCS_DIR, combined_docx_filename)

        # 1. Generate Single Combined Word Document (.docx)
        generate_batch_docx(batch_list, combined_docx_filepath)

        # 2. Generate Single PDF for each item and merge them into a single PDF file
        single_pdf_paths = []
        for idx, item in enumerate(batch_list, start=1):
            temp_docx = os.path.join(DOCS_DIR, f"temp_batch_{unique_id}_{idx}.docx")
            temp_pdf = os.path.join(DOCS_DIR, f"temp_batch_{unique_id}_{idx}.pdf")
            try:
                generate_kheralu_order_docx(item, temp_docx)
                generate_pdf_document(item, temp_docx, temp_pdf)
                single_pdf_paths.append(temp_pdf)
            except Exception as single_err:
                print(f"Error rendering single batch item {idx}: {single_err}")

        combined_pdf_generated = False
        if len(single_pdf_paths) > 0:
            try:
                merge_pdf_files(single_pdf_paths, combined_pdf_filepath)
                combined_pdf_generated = os.path.exists(combined_pdf_filepath)
            except Exception as merge_err:
                print(f"PDF merge error: {merge_err}")

        # Clean up temporary single files
        for pdf_p in single_pdf_paths:
            if os.path.exists(pdf_p):
                os.remove(pdf_p)
            docx_p = pdf_p.replace(".pdf", ".docx")
            if os.path.exists(docx_p):
                os.remove(docx_p)

        return jsonify({
            "status": "success",
            "batch_count": batch_count,
            "combined_pdf_filename": combined_pdf_filename if combined_pdf_generated else None,
            "combined_pdf_download_url": f"/download/{combined_pdf_filename}" if combined_pdf_generated else None,
            "combined_docx_filename": combined_docx_filename,
            "combined_docx_download_url": f"/download/{combined_docx_filename}"
        })
    except Exception as e:
        print(f"Error generating batch combined document: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(DOCS_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
