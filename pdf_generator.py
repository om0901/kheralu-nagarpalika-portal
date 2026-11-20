import io
import os
import warnings
from pypdf import PdfWriter, PdfReader
from fpdf import FPDF

# Suppress spurious namespace warning between pypdf and fpdf
warnings.filterwarnings("ignore", category=UserWarning, module="fpdf")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(BASE_DIR, "fonts")
REG_FONT_PATH = os.path.join(FONTS_DIR, "NotoSansGujarati-Regular.ttf")
BOLD_FONT_PATH = os.path.join(FONTS_DIR, "NotoSansGujarati-Bold.ttf")

DEFAULT_FALLBACK = "-------------------"

def create_configured_pdf():
    """
    Creates an FPDF instance configured for A4 portrait with HarfBuzz complex
    text shaping enabled for 100% accurate Gujarati conjuncts and matras.
    """
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(15, 10, 15)
    pdf.set_auto_page_break(False)

    # Enable HarfBuzz text shaping (solves matra / conjunct rendering bugs)
    pdf.set_text_shaping(use_shaping_engine=True, script="gujr", language="guj")

    if os.path.exists(REG_FONT_PATH):
        pdf.add_font("NotoGujarati", fname=REG_FONT_PATH)
        pdf.add_font("NotoGujarati", style="B", fname=BOLD_FONT_PATH if os.path.exists(BOLD_FONT_PATH) else REG_FONT_PATH)
    else:
        win_reg = r"C:\Windows\Fonts\Nirmala.ttf"
        win_bold = r"C:\Windows\Fonts\NirmalaB.ttf"
        pdf.add_font("NotoGujarati", fname=win_reg)
        pdf.add_font("NotoGujarati", style="B", fname=win_bold if os.path.exists(win_bold) else win_reg)

    return pdf

def gv(v):
    if not v or not str(v).strip():
        return DEFAULT_FALLBACK
    return str(v).strip()

def render_order_pages(pdf, data):
    """
    Renders exactly 2 pages for a single Kheralu Nagarpalika Sudhara Hukam application.
    """
    today_date = gv(data.get("today_date"))
    app_date = gv(data.get("application_date"))
    app_name = gv(data.get("applicant_name"))
    app_addr = gv(data.get("applicant_address"))

    rows_data = [
        ("1", "નોંધણી ક્રમાંક", gv(data.get("reg_no")), DEFAULT_FALLBACK),
        ("2", "જન્મ તારીખ", gv(data.get("dob")), DEFAULT_FALLBACK),
        ("3", "નોંધણી તારીખ", gv(data.get("reg_date")), DEFAULT_FALLBACK),
        ("4", "બાળકનુ નામ\n(ભાઈ,કુમાર,બેન)", gv(data.get("child_name")), gv(data.get("corr_child_name"))),
        ("5", "પિતાનુ નામ", gv(data.get("father_name")), gv(data.get("corr_father_name"))),
        ("6", "અટક", gv(data.get("surname")), gv(data.get("corr_surname"))),
        ("7", "માતાનુ નામ", gv(data.get("mother_name")), gv(data.get("corr_mother_name"))),
        ("8", "સરનામુ", gv(data.get("address")), gv(data.get("corr_address"))),
    ]

    def render_table():
        with pdf.table(
            col_widths=(12, 52, 52, 64),
            text_align="CENTER",
            markdown=True,
            line_height=4.8,
            padding=1.2
        ) as table:
            hdr = table.row()
            hdr.cell("**અ.નં**", colspan=1)
            hdr.cell("**રજીસ્ટરમા નોધાયેલ વિગત**", colspan=2)
            hdr.cell("**જન્મનોંધમાં સુધારો કરવાપત્ર વિગત**", colspan=1)
            for sr, desc, orig, corr in rows_data:
                r = table.row()
                r.cell(sr)
                r.cell(f"**{desc}**")
                r.cell(orig)
                is_changed = (corr != DEFAULT_FALLBACK)
                r.cell(f"**{corr}**" if is_changed else corr)

    # =========================================================
    # PAGE 1
    # =========================================================
    pdf.add_page()
    pdf.set_y(10)
    pdf.set_font("NotoGujarati", style="B", size=11)
    pdf.cell(w=0, h=5, text="(૧)", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("NotoGujarati", style="B", size=10)
    pdf.cell(w=0, h=5, text="જન્મ-મરણ શાખા", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="ખેરાલુ નગરપાલિકા", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text=f"તા. {today_date}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_font("NotoGujarati", style="BU", size=10.5)
    pdf.cell(w=0, h=5.5, text="સાદરરજુ.", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    p1_body = (
        "જન્મ-મરણ નોધણી અધિનિયમ, ૧૯૬૯ ની કલમ(૧૫) અને ગુજરાત જન્મ અને મરણ નોધણી નિયમો ૨૦૦૪ ના નિયમ(૧૧) "
        "તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર "
        "મુજબ અરજદારશ્રી દ્વારા જન્મની નોંધમાં સુધારો કરવા અરજી કરેલ છે. અરજદારે અરજીસાથે સામેલ રાખેલ સાધનિક "
        "પુરાવા જેવાકે L.C. આધારકાર્ડ, રેશનકાર્ડ, ચુંટણીકાર્ડ, સોગંદનામુ, ડોકટરનું પ્રમાણપત્ર, અસલજન્મ પ્રમાણપત્ર, "
        "વિગેરે જોતા નીચેની વિગતે જન્મ નીનોંધ માં સુધારો કરવાપાત્ર છે."
    )
    pdf.set_font("NotoGujarati", size=9.5)
    pdf.multi_cell(w=0, h=4.8, text="        " + p1_body, align="J", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    render_table()

    # Signature box Page 1
    pdf.ln(16)
    pdf.set_font("NotoGujarati", style="B", size=10)
    pdf.cell(w=0, h=5, text="ક્લાર્ક", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="જન્મ-મરણ", align="R", new_x="LMARGIN", new_y="NEXT")

    # Pin "(પાછળ) પાનાં નં - ૨" strictly to the bottom-right corner of Page 1
    pdf.set_y(280)
    pdf.set_font("NotoGujarati", style="B", size=9.5)
    pdf.cell(w=0, h=6, text="(પાછળ) પાનાં નં - ૨", align="R", new_x="LMARGIN", new_y="NEXT")

    # =========================================================
    # PAGE 2
    # =========================================================
    pdf.add_page()
    pdf.set_y(10)
    pdf.set_font("NotoGujarati", style="B", size=11)
    pdf.cell(w=0, h=5, text="(૨)", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("NotoGujarati", style="B", size=10)
    pdf.cell(w=0, h=5, text="ખેરાલુ નગરપાલિકા", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="જા.નં.", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="તા. ", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.5)

    pdf.set_font("NotoGujarati", style="BU", size=10.5)
    pdf.cell(w=0, h=5.5, text="વંચાણે લીધુ.", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.5)

    preambles = [
        f"1. અરજદાર શ્રી {app_name} રહે. {app_addr} તા. {app_date} ની અરજી તથા રજુકરેલ પુરાવા",
        "2. જન્મ-મરણ નોંધણી અધિનીયમ ૧૯૬૯ની કલમ-૧૫",
        "3. ગુજરાત જન્મ-મરણ નોધણીનિયમો ૨૦૦૪ ના નિયમ-૧૧ તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર",
        "4. રજીસ્ટ્રાર શ્રી (જન્મ-મરણ) અને ચીફ ઓફીસર શ્રી ખેરાલુ નગરપાલિકા કચેરી હુંકમ નં.૨૧૬૬/૨૪-૨૫ તા.૦૫/૧૨/૨૦૨૪",
        f"5. જન્મ-મરણ શાખાની તા. {today_date} ની નોંધ"
    ]

    pdf.set_font("NotoGujarati", size=9)
    for p in preambles:
        pdf.multi_cell(w=0, h=4.2, text=p, align="L", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(0.8)

    pdf.ln(1.5)
    pdf.set_font("NotoGujarati", style="BU", size=10.5)
    pdf.cell(w=0, h=5.5, text="કચેરી હુકમ", align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.5)

    p2_body = (
        f"વંચાણે લીધેલ આમુખ. ૧ થી અરજદાર શ્રી {app_name}, ધ્વારા તા {app_date} ની અરજી તથા સામેલ પુરાવા તેમજ "
        "વંચાણે લીધેલ આમુખ ૨ થી ૩ જન્મ-મરણ નોંધણી નિયમોની કલમ થી મળેલ સત્તાનીરૂએ વંચાણે લીધેલ આમુખ ૪ ના "
        "ચીફ ઓફીસર શ્રીના કચેરી હુકમ થી મળેલ સત્તાનીરૂએ જન્મ નોંધણી રજીસ્ટર મા નીચેમુજબ અરજદાર ના "
        "પુત્ર-પુત્રી/પૌત્ર-પૌત્રી/ભાણા-ભાણી ના જન્મની નોંધમા સુધારો કરવા હુકમ કરું છુ."
    )
    pdf.set_font("NotoGujarati", size=9.5)
    pdf.multi_cell(w=0, h=4.8, text="        " + p2_body, align="J", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    render_table()

    # Signature box Page 2
    pdf.ln(16)
    pdf.set_font("NotoGujarati", style="B", size=10)
    pdf.cell(w=0, h=5, text="સબરજીસ્ટ્રાર", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="જન્મ-મરણ નોંધણી", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(w=0, h=5, text="ખેરાલુ નગરપાલિકા", align="R", new_x="LMARGIN", new_y="NEXT")

def generate_pdf_bytes(data):
    """
    Generates exact 2-page PDF document directly as an in-memory BytesIO stream
    with HarfBuzz complex text shaping for accurate Gujarati conjuncts.
    """
    pdf = create_configured_pdf()
    render_order_pages(pdf, data)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf

def generate_batch_pdf_bytes(batch_list):
    """
    Generates a single combined multi-page PDF for the entire batch directly in memory.
    Each item occupies exactly 2 pages.
    """
    pdf = create_configured_pdf()
    for item in batch_list:
        render_order_pages(pdf, item)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf

def merge_pdf_files(pdf_input_list, output_target):
    """
    Merges multiple PDF file paths or BytesIO streams into output_target using pypdf.
    output_target can be a file path string or a file-like object (io.BytesIO).
    """
    writer = PdfWriter()
    for item in pdf_input_list:
        if isinstance(item, str):
            reader = PdfReader(item)
        else:
            item.seek(0)
            reader = PdfReader(item)
        for page in reader.pages:
            writer.add_page(page)

    if isinstance(output_target, str):
        with open(output_target, "wb") as f:
            writer.write(f)
        return output_target
    else:
        writer.write(output_target)
        output_target.seek(0)
        return output_target

def generate_pdf_document(data, docx_filepath, pdf_filepath):
    """
    Legacy helper: Generates PDF document to disk.
    """
    buf = generate_pdf_bytes(data)
    with open(pdf_filepath, "wb") as f:
        f.write(buf.getvalue())
    return pdf_filepath
