import os
import subprocess
import shutil
from jinja2 import Template
import pypdfium2

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from doc_generator import generate_kheralu_order_docx

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
FALLBACK_EDGE_PATH = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"

def get_browser_executable():
    """Finds available Chrome/Chromium/Edge browser executable across Windows and Linux."""
    # Check PATH first (works on Linux and configured Windows)
    for name in ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge", "microsoft-edge"]:
        path = shutil.which(name)
        if path:
            return path
            
    # Check standard Windows paths
    candidates = [
        EDGE_PATH,
        FALLBACK_EDGE_PATH,
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
            
    return None

def render_html_template(data, template_path):
    """Renders Jinja2 HTML template with user input data."""
    with open(template_path, "r", encoding="utf-8") as f:
        html_content = f.read()
        
    template = Template(html_content)
    
    DEFAULT_FALLBACK = "-------------------"
    def get_corr_val(key):
        val = data.get(key)
        if not val or not str(val).strip():
            return DEFAULT_FALLBACK
        return str(val).strip()

    context = {
        "today_date": data.get("today_date", ""),
        "application_date": data.get("application_date", ""),
        "applicant_name": data.get("applicant_name", ""),
        "applicant_address": data.get("applicant_address", ""),
        "reg_no": data.get("reg_no", ""),
        "dob": data.get("dob", ""),
        "reg_date": data.get("reg_date", ""),
        "child_name": data.get("child_name", ""),
        "father_name": data.get("father_name", ""),
        "surname": data.get("surname", ""),
        "mother_name": data.get("mother_name", ""),
        "address": data.get("address", ""),
        "corr_child_name": get_corr_val("corr_child_name"),
        "corr_father_name": get_corr_val("corr_father_name"),
        "corr_surname": get_corr_val("corr_surname"),
        "corr_mother_name": get_corr_val("corr_mother_name"),
        "corr_address": get_corr_val("corr_address")
    }
    
    return template.render(**context)

def convert_html_to_pdf_edge(html_path, pdf_path):
    """Converts HTML file to PDF using headless Chrome/Chromium/Edge print-to-pdf."""
    browser_exe = get_browser_executable()
    if not browser_exe:
        return False

    abs_html = os.path.abspath(html_path)
    abs_pdf = os.path.abspath(pdf_path)
    
    file_url = f"file:///{abs_html.replace('\\', '/')}"
    cmd = [
        browser_exe,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--no-pdf-header-footer",
        f"--print-to-pdf={abs_pdf}",
        file_url
    ]
    
    try:
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20)
        return os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0
    except Exception as e:
        print(f"Browser PDF generation failed: {e}")
        return False

def convert_docx_to_pdf_word(docx_path, pdf_path):
    """Fallback: Converts .docx file to .pdf using MS Word COM automation (Windows only)."""
    if not HAS_WIN32COM:
        return False
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17)
        doc.Close()
        word.Quit()
        return os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0
    except Exception as e:
        print(f"Word COM conversion failed: {e}")
        return False

def generate_pdf_document(data, docx_filepath, pdf_filepath):
    """
    Generates exact 2-page PDF document.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    html_template_path = os.path.join(base_dir, "templates", "order_print_template.html")
    temp_html_path = pdf_filepath.replace(".pdf", "_temp.html")
    
    pdf_success = False
    
    # 1. Primary Strategy: HTML to PDF via Edge Headless
    if os.path.exists(html_template_path):
        try:
            rendered_html = render_html_template(data, html_template_path)
            with open(temp_html_path, "w", encoding="utf-8") as f:
                f.write(rendered_html)
                
            pdf_success = convert_html_to_pdf_edge(temp_html_path, pdf_filepath)
            
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
        except Exception as err:
            print(f"HTML PDF Strategy Error: {err}")

    # 2. Secondary Strategy: Word COM PDF Conversion
    if not pdf_success and HAS_WIN32COM:
        if not os.path.exists(docx_filepath):
            generate_kheralu_order_docx(data, docx_filepath)
        pdf_success = convert_docx_to_pdf_word(docx_filepath, pdf_filepath)
        
    if not pdf_success or not os.path.exists(pdf_filepath):
        # Guarantee Word .docx is generated so user always has Word file download
        if not os.path.exists(docx_filepath):
            generate_kheralu_order_docx(data, docx_filepath)
        print("Note: PDF conversion unavailable on this environment, Word document generated.")
        
    return pdf_filepath

def merge_pdf_files(pdf_filepath_list, output_combined_pdf_path):
    """
    Merges multiple single 2-page PDF files into a single combined PDF document.
    """
    combined_pdf = pypdfium2.PdfDocument.new()
    for single_pdf_path in pdf_filepath_list:
        if os.path.exists(single_pdf_path):
            src_pdf = pypdfium2.PdfDocument(single_pdf_path)
            combined_pdf.import_pages(src_pdf)
            src_pdf.close()
            
    combined_pdf.save(output_combined_pdf_path)
    combined_pdf.close()
    return output_combined_pdf_path

def generate_pdf_bytes(data):
    """
    Generates exact 2-page PDF document directly as an in-memory BytesIO stream.
    Cleans up any temporary intermediate files immediately within the same request.
    """
    import io
    import uuid
    import tempfile

    base_dir = os.path.dirname(os.path.abspath(__file__))
    html_template_path = os.path.join(base_dir, "templates", "order_print_template.html")

    if not os.path.exists(html_template_path):
        return None

    temp_dir = tempfile.gettempdir()
    unique_suffix = uuid.uuid4().hex[:8]
    temp_html_path = os.path.join(temp_dir, f"order_{unique_suffix}.html")
    temp_pdf_path = os.path.join(temp_dir, f"order_{unique_suffix}.pdf")
    temp_docx_path = os.path.join(temp_dir, f"order_{unique_suffix}.docx")

    try:
        # 1. Primary Strategy: HTML to PDF via Headless Browser
        rendered_html = render_html_template(data, html_template_path)
        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(rendered_html)

        pdf_success = convert_html_to_pdf_edge(temp_html_path, temp_pdf_path)

        # 2. Secondary Strategy: Word COM (Windows only)
        if not pdf_success and HAS_WIN32COM:
            generate_kheralu_order_docx(data, temp_docx_path)
            pdf_success = convert_docx_to_pdf_word(temp_docx_path, temp_pdf_path)

        if pdf_success and os.path.exists(temp_pdf_path) and os.path.getsize(temp_pdf_path) > 0:
            with open(temp_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            buf = io.BytesIO(pdf_bytes)
            buf.seek(0)
            return buf

        return None
    finally:
        for p in [temp_html_path, temp_pdf_path, temp_docx_path]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

def generate_batch_pdf_bytes(batch_list):
    """
    Generates a single combined PDF document for the entire batch directly as an in-memory BytesIO stream.
    """
    import io
    combined_pdf = pypdfium2.PdfDocument.new()
    opened_docs = []

    try:
        for item in batch_list:
            single_buf = generate_pdf_bytes(item)
            if single_buf:
                src_doc = pypdfium2.PdfDocument(single_buf.getvalue())
                opened_docs.append(src_doc)
                combined_pdf.import_pages(src_doc)

        if len(combined_pdf) == 0:
            return None

        buf = io.BytesIO()
        combined_pdf.save(buf)
        buf.seek(0)
        return buf
    finally:
        for d in opened_docs:
            try:
                d.close()
            except Exception:
                pass
        try:
            combined_pdf.close()
        except Exception:
            pass
