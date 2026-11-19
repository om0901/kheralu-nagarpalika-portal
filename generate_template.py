import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor

def create_base_template(output_path="template.docx"):
    """
    Creates a clean Word document template styled for Kheralu Nagarpalika.
    """
    doc = Document()
    
    # Set standard page margins (0.75 inch all around)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Configure default style font
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Hind Vadodara'
    normal_font.size = Pt(11)
    normal_font.color.rgb = RGBColor(0, 0, 0)
    
    doc.save(output_path)
    print(f"Template successfully saved to {output_path}")

if __name__ == "__main__":
    os.makedirs("generated_docs", exist_ok=True)
    create_base_template()
