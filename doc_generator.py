import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

DEFAULT_FALLBACK = "-------------------"

def set_cell_background(cell, fill_hex):
    """Sets background shading of a cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=35, bottom=35, left=90, right=90):
    """Sets inner margins for a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{edge}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="000000", sz="4", val="single"):
    """Applies clean solid borders to a table."""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def add_gujarati_run(paragraph, text, font_name="Nirmala UI", font_size=10, bold=False, italic=False, underline=False):
    """Adds a run with explicit Gujarati Unicode font configuration across ascii, hAnsi, cs, and hint."""
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}" w:eastAsia="{font_name}"/>')
    rPr.append(rFonts)
    hint_elm = parse_xml(f'<w:hint {nsdecls("w")} w:val="cs"/>')
    rPr.append(hint_elm)
    return run

def add_styled_paragraph(doc, text="", font_name="Nirmala UI", font_size=10, bold=False, italic=False, underline=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, space_before=0, line_spacing=1.12, first_line_indent=None):
    """Adds a paragraph with Gujarati font formatting and optional first-line indent."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = line_spacing
    
    if first_line_indent:
        p.paragraph_format.first_line_indent = Inches(first_line_indent)
        
    if text:
        add_gujarati_run(p, text, font_name=font_name, font_size=font_size, bold=bold, italic=italic, underline=underline)
    return p

def format_cell_text(cell, text, font_name="Nirmala UI", font_size=9, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    """Formats text inside a table cell with proper Gujarati font and CENTER alignment."""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.12
    p.text = ""
    add_gujarati_run(p, text, font_name=font_name, font_size=font_size, bold=bold)

def build_exact_4col_table(doc, data):
    """
    Builds the EXACT 4-column table structure:
    Header Row:
      - Col 0: અ.નં
      - Col 1 & 2 (MERGED): રજીસ્ટરમા નોધાયેલ વિગત
      - Col 3: જન્મનોંધમાં સુધારો કરવાપત્ર વિગત
    """
    table = doc.add_table(rows=9, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, sz="4")
    
    col_widths = [Inches(0.45), Inches(1.85), Inches(1.9), Inches(2.5)]

    # Merge Col 1 and Col 2 in Header Row
    hdr_cells = table.rows[0].cells
    hdr_cells[1].merge(hdr_cells[2])
    
    hdr_cells[0].width = col_widths[0]
    table.rows[0].cells[1].width = col_widths[1] + col_widths[2]
    hdr_cells[3].width = col_widths[3]
    
    # Format Header Titles (CENTER ALIGNED)
    format_cell_text(hdr_cells[0], "અ.નં", font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell_text(table.rows[0].cells[1], "રજીસ્ટરમા નોધાયેલ વિગત", font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell_text(hdr_cells[3], "જન્મનોંધમાં સુધારો કરવાપત્ર વિગત", font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    set_cell_background(hdr_cells[0], "F8FAFC")
    set_cell_background(table.rows[0].cells[1], "F8FAFC")
    set_cell_background(hdr_cells[3], "F8FAFC")
    
    set_cell_margins(hdr_cells[0], top=40, bottom=40)
    set_cell_margins(table.rows[0].cells[1], top=40, bottom=40)
    set_cell_margins(hdr_cells[3], top=40, bottom=40)

    def get_val(val_str):
        if not val_str or not str(val_str).strip():
            return DEFAULT_FALLBACK
        return str(val_str).strip()

    row_defs = [
        ("1", "નોંધણી ક્રમાંક", get_val(data.get("reg_no")), DEFAULT_FALLBACK),
        ("2", "જન્મ તારીખ", get_val(data.get("dob")), DEFAULT_FALLBACK),
        ("3", "નોંધણી તારીખ", get_val(data.get("reg_date")), DEFAULT_FALLBACK),
        ("4", "બાળકનુ નામ\n(ભાઈ,કુમાર,બેન)", get_val(data.get("child_name")), get_val(data.get("corr_child_name"))),
        ("5", "પિતાનુ નામ", get_val(data.get("father_name")), get_val(data.get("corr_father_name"))),
        ("6", "અટક", get_val(data.get("surname")), get_val(data.get("corr_surname"))),
        ("7", "માતાનુ નામ", get_val(data.get("mother_name")), get_val(data.get("corr_mother_name"))),
        ("8", "સરનામુ", get_val(data.get("address")), get_val(data.get("corr_address")))
    ]

    for row_idx, (sr, field_desc, orig_val, corr_val) in enumerate(row_defs, start=1):
        row_cells = table.rows[row_idx].cells
        
        row_cells[0].width = col_widths[0]
        row_cells[1].width = col_widths[1]
        row_cells[2].width = col_widths[2]
        row_cells[3].width = col_widths[3]
        
        format_cell_text(row_cells[0], sr, align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell_text(row_cells[1], field_desc, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell_text(row_cells[2], orig_val, align=WD_ALIGN_PARAGRAPH.CENTER)
        
        is_changed = corr_val != DEFAULT_FALLBACK
        format_cell_text(row_cells[3], corr_val, bold=is_changed, align=WD_ALIGN_PARAGRAPH.CENTER)
        
        set_cell_margins(row_cells[0], top=30, bottom=30)
        set_cell_margins(row_cells[1], top=30, bottom=30)
        set_cell_margins(row_cells[2], top=30, bottom=30)
        set_cell_margins(row_cells[3], top=30, bottom=30)
        
    return table

def render_order_to_document(doc, data):
    """
    Appends a complete 2-page Birth Registration Correction Order to a Document instance.
    """
    today_date = data.get('today_date', '')
    app_date = data.get('application_date', '')
    app_name = data.get('applicant_name', '')
    app_addr = data.get('applicant_address', '')
    reg_no = data.get('reg_no', '')

    # PAGE 1
    add_styled_paragraph(doc, "(૧)", font_size=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    add_styled_paragraph(doc, "જન્મ-મરણ શાખા", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=1)
    add_styled_paragraph(doc, "ખેરાલુ નગરપાલિકા", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=1)
    add_styled_paragraph(doc, f"તા. {today_date}", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)
    add_styled_paragraph(doc, "સાદરરજુ.", font_size=10.5, bold=True, underline=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=4)

    p1_body_text = (
        "જન્મ-મરણ નોધણી અધિનિયમ, ૧૯૬૯ ની કલમ(૧૫) અને ગુજરાત જન્મ અને મરણ નોધણી નિયમો ૨૦૦૪ ના નિયમ(૧૧) "
        "તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર "
        "મુજબ અરજદારશ્રી દ્વારા જન્મની નોંધમાં સુધારો કરવા અરજી કરેલ છે. અરજદારે અરજીસાથે સામેલ રાખેલ સાધનિક "
        "પુરાવા જેવાકે L.C. આધારકાર્ડ, રેશનકાર્ડ, ચુંટણીકાર્ડ, સોગંદનામુ, ડોકટરનું પ્રમાણપત્ર, અસલજન્મ પ્રમાણપત્ર, "
        "વિગેરે જોતા નીચેની વિગતે જન્મ નીનોંધ માં સુધારો કરવાપાત્ર છે."
    )
    add_styled_paragraph(doc, p1_body_text, font_size=9.5, space_after=6, line_spacing=1.14, first_line_indent=0.5)

    build_exact_4col_table(doc, data)

    add_styled_paragraph(doc, "ક્લાર્ક", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=75, space_after=1)
    add_styled_paragraph(doc, "જન્મ-મરણ", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)

    add_styled_paragraph(doc, "(પાછળ) પાનાં નં - ૨", font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=60, space_after=0)

    # PAGE BREAK TO PAGE 2
    doc.add_page_break()

    # PAGE 2
    add_styled_paragraph(doc, "(૨)", font_size=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_styled_paragraph(doc, "ખેરાલુ નગરપાલિકા", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=1)
    add_styled_paragraph(doc, "જા.નં.", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=1)
    add_styled_paragraph(doc, "તા. ", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)

    add_styled_paragraph(doc, "વંચાણે લીધુ.", font_size=10.5, bold=True, underline=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=3)

    preambles = [
        f"1. અરજદાર શ્રી {app_name} રહે. {app_addr} તા. {app_date} ની અરજી તથા રજુકરેલ પુરાવા",
        "2. જન્મ-મરણ નોંધણી અધિનીયમ ૧૯૬૯ની કલમ-૧૫",
        "3. ગુજરાત જન્મ-મરણ નોધણીનિયમો ૨૦૦૪ ના નિયમ-૧૧ તથા મુખ્ય રજીસ્ટ્રાર (જન્મ-મરણ) અને કમિશ્નર શ્રી (આરોગ્ય), ગાંધીનગરના તારીખ-૨૫/૦૯/૨૦૨૪નો પરિપત્ર",
        "4. રજીસ્ટ્રાર શ્રી (જન્મ-મરણ) અને ચીફ ઓફીસર શ્રી ખેરાલુ નગરપાલિકા કચેરી હુંકમ નં.૨૧૬૬/૨૪-૨૫ તા.૦૫/૧૨/૨૦૨૪",
        f"5. જન્મ-મરણ શાખાની તા. {today_date} ની નોંધ"
    ]
    
    for item in preambles:
        add_styled_paragraph(doc, item, font_size=9, space_after=1.5, line_spacing=1.12)

    add_styled_paragraph(doc, "કચેરી હુકમ", font_size=10.5, bold=True, underline=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=4)

    hukam_text = (
        f"વંચાણે લીધેલ આમુખ. ૧ થી અરજદાર શ્રી {app_name}, ધ્વારા તા {app_date} ની અરજી તથા સામેલ પુરાવા "
        "તેમજ વંચાણે લીધેલ આમુખ ૨ થી ૩ જન્મ-મરણ નોંધણી નિયમોની કલમ થી મળેલ સત્તાનીરૂએ વંચાણે લીધેલ આમુખ ૪ "
        "ના ચીફ ઓફીસર શ્રીના કચેરી હુકમ થી મળેલ સત્તાનીરૂએ જન્મ નોંધણી રજીસ્ટર મા નીચેમુજબ અરજદાર ના "
        "પુત્ર-પુત્રી/પૌત્ર-પૌત્રી/ભાણા-ભાણી ના જન્મની નોંધમા સુધારો કરવા હુકમ કરું છુ."
    )
    add_styled_paragraph(doc, hukam_text, font_size=9, space_after=6, line_spacing=1.14, first_line_indent=0.5)

    build_exact_4col_table(doc, data)

    add_styled_paragraph(doc, "સબરજીસ્ટ્રાર", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=75, space_after=1)
    add_styled_paragraph(doc, "જન્મ-મરણ નોંધણી", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=1)
    add_styled_paragraph(doc, "ખેરાલુ નગરપાલિકા", font_size=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)

def generate_kheralu_order_docx(data, output_filepath):
    """
    Generates single 2-page Birth Registration Correction Order document.
    output_filepath can be a file path string or a file-like object (io.BytesIO).
    """
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    render_order_to_document(doc, data)
    doc.save(output_filepath)
    return output_filepath

def generate_kheralu_order_docx_bytes(data):
    """
    Generates single 2-page Word document directly into an in-memory BytesIO stream.
    """
    import io
    buf = io.BytesIO()
    generate_kheralu_order_docx(data, buf)
    buf.seek(0)
    return buf

def generate_batch_docx(batch_list, output_filepath):
    """
    Generates a single combined Word (.docx) file containing multiple 2-page applications consecutively.
    output_filepath can be a file path string or a file-like object (io.BytesIO).
    """
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    for idx, data_item in enumerate(batch_list):
        render_order_to_document(doc, data_item)
        if idx < len(batch_list) - 1:
            doc.add_page_break()  # Page break between applications

    doc.save(output_filepath)
    return output_filepath

def generate_batch_docx_bytes(batch_list):
    """
    Generates a combined Word document directly into an in-memory BytesIO stream.
    """
    import io
    buf = io.BytesIO()
    generate_batch_docx(batch_list, buf)
    buf.seek(0)
    return buf

if __name__ == "__main__":
    sample_data = {
        "today_date": "07/08/2026",
        "applicant_name": "ચૌધરી રમેશભાઈ વિષ્ણુભાઈ",
        "applicant_address": "મુ. ખેરાલુ, તા. ખેરાલુ, જી. મહેસાણા",
        "application_date": "01/08/2026",
        "reg_no": "845/2024",
        "dob": "12/04/2024",
        "reg_date": "18/04/2024",
        "child_name": "દિયા",
        "father_name": "રમેશભાઈ",
        "surname": "ચૌધરી",
        "mother_name": "કવિતાબેન",
        "address": "ખેરાલુ",
        "corr_child_name": "દિયાબેન",
        "corr_father_name": "રમેશકુમાર",
        "corr_surname": "",
        "corr_mother_name": "",
        "corr_address": ""
    }
    os.makedirs("generated_docs", exist_ok=True)
    out_file = generate_kheralu_order_docx(sample_data, "generated_docs/exact_100px_margin.docx")
    print(f"Single docx generated at: {out_file}")
