import io
import os
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(BASE_DIR, "fonts")

DEFAULT_FALLBACK = "-------------------"

# ---------------------------------------------------------
# FONT REGISTRATION
# ---------------------------------------------------------
FONT_REGULAR = "NotoSansGujarati"
FONT_BOLD = "NotoSansGujarati-Bold"

def register_gujarati_fonts():
    """Registers Gujarati TrueType fonts with ReportLab."""
    reg_path = os.path.join(FONTS_DIR, "NotoSansGujarati-Regular.ttf")
    bold_path = os.path.join(FONTS_DIR, "NotoSansGujarati-Bold.ttf")

    # Fallback to Windows system fonts if repository fonts are missing
    if not os.path.exists(reg_path):
        win_reg = r"C:\Windows\Fonts\Nirmala.ttf"
        win_bold = r"C:\Windows\Fonts\NirmalaB.ttf"
        if os.path.exists(win_reg):
            reg_path = win_reg
            bold_path = win_bold if os.path.exists(win_bold) else win_reg

    if os.path.exists(reg_path):
        try:
            pdfmetrics.registerFont(TTFont(FONT_REGULAR, reg_path))
            pdfmetrics.registerFont(TTFont(FONT_BOLD, bold_path if os.path.exists(bold_path) else reg_path))
            return True
        except Exception as e:
            print(f"Font registration warning: {e}")
            return False
    return False

# Register fonts upon module import
register_gujarati_fonts()

def escape_xml(text):
    """Escapes XML entities so ReportLab Paragraphs render safely without XML syntax errors."""
    if text is None:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def build_order_table(data, style_cell, style_cell_bold):
    """
    Builds the EXACT 4-column table structure:
      - Col 0 (width 30): અ.નં
      - Col 1 & 2 merged (width 150 + 150 = 300): રજીસ્ટરમા નોધાયેલ વિગત
      - Col 3 (width 180): જન્મનોંધમાં સુધારો કરવાપત્ર વિગત
    """
    def gv(val):
        if not val or not str(val).strip():
            return DEFAULT_FALLBACK
        return str(val).strip()

    row_defs = [
        ("1", "નોંધણી ક્રમાંક", gv(data.get("reg_no")), DEFAULT_FALLBACK),
        ("2", "જન્મ તારીખ", gv(data.get("dob")), DEFAULT_FALLBACK),
        ("3", "નોંધણી તારીખ", gv(data.get("reg_date")), DEFAULT_FALLBACK),
        ("4", "બાળકનુ નામ<br/>(ભાઈ,કુમાર,બેન)", gv(data.get("child_name")), gv(data.get("corr_child_name"))),
        ("5", "પિતાનુ નામ", gv(data.get("father_name")), gv(data.get("corr_father_name"))),
        ("6", "અટક", gv(data.get("surname")), gv(data.get("corr_surname"))),
        ("7", "માતાનુ નામ", gv(data.get("mother_name")), gv(data.get("corr_mother_name"))),
        ("8", "સરનામુ", gv(data.get("address")), gv(data.get("corr_address")))
    ]

    table_data = [
        [
            Paragraph("<b>અ.નં</b>", style_cell_bold),
            Paragraph("<b>રજીસ્ટરમા નોધાયેલ વિગત</b>", style_cell_bold),
            "",
            Paragraph("<b>જન્મનોંધમાં સુધારો કરવાપત્ર વિગત</b>", style_cell_bold)
        ]
    ]

    for sr, field_desc, orig_val, corr_val in row_defs:
        is_changed = (corr_val != DEFAULT_FALLBACK)
        corr_style = style_cell_bold if is_changed else style_cell
        
        table_data.append([
            Paragraph(escape_xml(sr), style_cell),
            Paragraph(f"<b>{field_desc}</b>", style_cell_bold),
            Paragraph(escape_xml(orig_val), style_cell),
            Paragraph(escape_xml(corr_val), corr_style)
        ])

    table = Table(table_data, colWidths=[30, 150, 150, 180])
    table.setStyle(TableStyle([
        ('SPAN', (1, 0), (2, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F8FAFC')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))
    return table

def render_order_story(data):
    """
    Constructs the exact 2-page flowable story for a single Kheralu Nagarpalika Sudhara Hukam order.
    """
    today_date = escape_xml(data.get("today_date", ""))
    app_date = escape_xml(data.get("application_date", ""))
    app_name = escape_xml(data.get("applicant_name", ""))
    app_addr = escape_xml(data.get("applicant_address", ""))

    style_center = ParagraphStyle('Center', fontName=FONT_BOLD, fontSize=10.5, leading=14, alignment=TA_CENTER)
    style_right = ParagraphStyle('Right', fontName=FONT_BOLD, fontSize=9.5, leading=13, alignment=TA_RIGHT)
    style_left = ParagraphStyle('Left', fontName=FONT_REGULAR, fontSize=9, leading=12.5, alignment=TA_LEFT)
    style_left_bold = ParagraphStyle('LeftBold', fontName=FONT_BOLD, fontSize=9.5, leading=13, alignment=TA_LEFT)
    style_body = ParagraphStyle('Body', fontName=FONT_REGULAR, fontSize=9, leading=12.5, alignment=TA_JUSTIFY, firstLineIndent=20)
    style_cell = ParagraphStyle('Cell', fontName=FONT_REGULAR, fontSize=8.5, leading=11, alignment=TA_CENTER)
    style_cell_bold = ParagraphStyle('CellBold', fontName=FONT_BOLD, fontSize=8.5, leading=11, alignment=TA_CENTER)

    story = []

    # =========================================================
    # PAGE 1
    # =========================================================
    story.append(Paragraph("(૧)", style_center))
    story.append(Spacer(1, 2))
    story.append(Paragraph("જન્મ-મરણ શાખા", style_right))
    story.append(Paragraph("ખેરાલુ નગરપાલિકા", style_right))
    story.append(Paragraph(f"તા. {today_date}", style_right))
    story.append(Spacer(1, 4))
    story.append(Paragraph("<u><b>સાદરરજુ,</b></u>", style_left_bold))
    story.append(Spacer(1, 3))

    p1_body_text = (
        "જન્મ-મરણ નોધણી અધિનિયમ, ૧૯૬૯ ની કલમ(૧૫) અને ગુજરાત જન્મ અને મરણ નોધણી નિયમો ૨૦૦૪ ના નિયમ(૧૧) "
        "તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર "
        "મુજબ અરજદારશ્રી દ્વારા જન્મની નોંધમાં સુધારો કરવા અરજી કરેલ છે. અરજદારે અરજીસાથે સામેલ રાખેલ સાધનિક "
        "પુરાવા જેવાકે L.C. આધારકાર્ડ, રેશનકાર્ડ, ચુંટણીકાર્ડ, સોગંદનામુ, ડોકટરનું પ્રમાણપત્ર, અસલજન્મ પ્રમાણપત્ર, "
        "વિગેરે જોતા નીચેની વિગતે જન્મ નીનોંધ માં સુધારો કરવાપાત્ર છે."
    )
    story.append(Paragraph(p1_body_text, style_body))
    story.append(Spacer(1, 6))

    story.append(build_order_table(data, style_cell, style_cell_bold))
    story.append(Spacer(1, 45))

    story.append(Paragraph("ક્લાર્ક", style_right))
    story.append(Paragraph("જન્મ-મરણ", style_right))
    story.append(Spacer(1, 35))
    story.append(Paragraph("(પાછળ) પાનં -૨", style_right))

    # =========================================================
    # PAGE 2
    # =========================================================
    story.append(PageBreak())

    story.append(Paragraph("(૨)", style_center))
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>ખેરાલુ નગરપાલિકા</b>", style_left_bold))
    story.append(Paragraph("<b>જા.નં.</b>", style_left_bold))
    story.append(Paragraph("<b>તા. </b>", style_left_bold))
    story.append(Spacer(1, 4))

    story.append(Paragraph("<u><b>વંચાણે લીધુ.</b></u>", style_left_bold))
    story.append(Spacer(1, 2))

    preambles = [
        f"1. અરજદાર શ્રી {app_name} રહે. {app_addr} તા. {app_date} ની અરજી તથા રજુકરેલ પુરાવા",
        "2. જન્મ-મરણ નોંધણી અધિનીયમ ૧૯૬૯ની કલમ-૧૫",
        "3. ગુજરાત જન્મ-મરણ નોધણીનિયમો ૨૦૦૪ ના નિયમ-૧૧ તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર",
        "4. રજીસ્ટાર શ્રી (જન્મ-મરણ) અને ચીફ ઓફીસર શ્રી ખેરાલુ નગરપાલિકા કચેરી હુંકમ નં.૨૧૬૬/૨૪-૨૫ તા.૦૫/૧૨/૨૦૨૪",
        f"5. જન્મ-મરણ શાખાની તા. {today_date} ની નોંધ"
    ]

    for p in preambles:
        story.append(Paragraph(p, style_left))
        story.append(Spacer(1, 1))

    story.append(Spacer(1, 4))
    story.append(Paragraph("<u><b>કચેરી હુકમ</b></u>", style_left_bold))
    story.append(Spacer(1, 2))

    p2_body_text = (
        f"વંચાણે લીધેલ આમુખ. ૧ થી અરજદાર શ્રી {app_name}, ધ્વારા તા {app_date} ની અરજી તથા સામેલ પુરાવા "
        "તેમજ વંચાણે લીધેલ આમુખ ૨ થી ૩ જન્મ-મરણ નોંધણી નિયમોની કલમ થી મળેલ સત્તાનીરૂએ વંચાણે લીધેલ આમુખ ૪ "
        "ના ચીફ ઓફીસર શ્રીના કચેરી હુકમ થી મળેલ સત્તાનીરૂએ જન્મ નોંધણી રજીસ્ટર મા નીચેમુજબ અરજદાર ના "
        "પુત્ર-પુત્રી/પૌત્ર-પૌત્રી/ભાણા-ભાણી ના જન્મની નોંધમા સુધારો કરવા હુકમ કરું છુ."
    )
    story.append(Paragraph(p2_body_text, style_body))
    story.append(Spacer(1, 6))

    story.append(build_order_table(data, style_cell, style_cell_bold))
    story.append(Spacer(1, 45))

    story.append(Paragraph("સબરજીસ્ટાર", style_right))
    story.append(Paragraph("જન્મ-મરણ નોધણી", style_right))
    story.append(Paragraph("ખેરાલુ નગરપાલિકા", style_right))

    return story

def generate_pdf_bytes(data):
    """
    Generates exact 2-page PDF document completely in memory using ReportLab.
    Returns io.BytesIO stream.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=42.5,
        rightMargin=42.5,
        topMargin=28.3,
        bottomMargin=28.3
    )

    story = render_order_story(data)
    doc.build(story)
    buf.seek(0)
    return buf

def generate_batch_pdf_bytes(batch_list):
    """
    Generates a single combined multi-page PDF for the entire batch directly in memory.
    Uses pypdf to merge the 2-page documents seamlessly.
    """
    writer = PdfWriter()

    for item in batch_list:
        single_buf = generate_pdf_bytes(item)
        single_buf.seek(0)
        reader = PdfReader(single_buf)
        for page in reader.pages:
            writer.add_page(page)

    combined_buf = io.BytesIO()
    writer.write(combined_buf)
    combined_buf.seek(0)
    return combined_buf

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
