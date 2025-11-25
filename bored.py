import pandas as pd
from docx import Document
from docx.shared import Pt,Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH,WD_TAB_ALIGNMENT

def add_bold_except_value(p, left_label, left_value, right_label, right_value):
    """
    In `paragraph`, add runs so that:
    - left_label is bold, left_value is not bold
    - right_label is bold, right_value is not bold
    And put left part first, then some spacing, then right part.
    """

    

    # Add left
    run1 = p.add_run(left_label)
    run1.bold = True
    run2 = p.add_run(left_value)
    run2.bold = False
    
    
    # Add spacing (you may adjust number of spaces or use tab)
    # Use tab character to try pushing the right side to the right

    p.add_run("                                       ")

    # Add right
    run3 = p.add_run(right_label)
    run3.bold = True
    run4 = p.add_run(right_value)
    run4.bold = False

    

def format_entry_docx(doc, row):
    """
    Given a Document and a row, add the entry block.
    """
    client = str(row.get("client", "")).strip()
    commodity = str(row.get("type", "")).strip() or "Units+ packages"
    nb_colis = row.get("qte", "")
    nb_colis = int(nb_colis)
    tonnage = row.get("poids", "")
    
    #nb_str = str(nb_colis)
    nb_str = f"{nb_colis:02d}"
    tonnage_str = str(tonnage)
    
    # Optionally, we can wrap all lines inside a single text frame or a table cell to better control vertical centering.
    # Here we just add paragraphs.
        
    # Top border line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=")
    run.bold = True  # border line in bold
    
    # Receiver / Commodity line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing=1
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.space_before=Pt(0)

    add_bold_except_value(
        p,
        left_label="Receiver : ",
        left_value=client + "    ",  # add some spacing if needed
        right_label="Comodity : ",
        right_value=commodity
    )
    #l=len(p.text)
    # Manifested Quantity / Tonnage line

    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing=1
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.space_before=Pt(0)

    add_bold_except_value(
        p,
        left_label="Manifested Quantity : ",
        left_value=nb_str +"  "+ commodity + "                               ",
        right_label="Tonnage : ",
        right_value=tonnage_str + "  Mt"
    )
    
    # Received line (only left side)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing=1
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.space_before=Pt(0)
    run = p.add_run("Received                ")
    run.bold = True
    run2 = p.add_run("Packaging damaged on board.")
    run2.bold = False
    
    
    # Total Received line (left only)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing=1
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.space_before=Pt(0)

    run = p.add_run("Total Received                                                             ")
    run.bold = True
    run2 = p.add_run(commodity)
    run2.bold = False
    
    # Final confirmation line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing=1
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.space_before=Pt(0)

    run = p.add_run("The Quantity Will Be confirmed after delivery Cargo.")
    run.bold = True  # the label (which is the entire sentence) in your spec is bold
    
    


def excel_to_docx_custom(input_excel, sheet_name=None,template_path="template.docx", output_docx="output.docx"):
    # Ensure openpyxl is installed; pandas will need it
    df = pd.read_excel(input_excel, sheet_name=sheet_name, engine="openpyxl")
    doc = Document(template_path)
    
    # Set default font for “Normal” style: Times New Roman, 12 pt
    style = doc.styles["Normal"]

    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)
    doc.styles["Normal"].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Optionally: to vertically center the content on a page, you might add a top margin or insert blank paragraphs
    # before and after. python-docx doesn’t have a built-in “vertical align center” setting for the page body.
   

    for idx, row in df.iterrows():
        format_entry_docx(doc, row)
        doc.add_paragraph("")  # blank line between entries
    
    doc.save(output_docx)
    print(f"Saved {output_docx}")

if __name__ == "__main__":
    excel_to_docx_custom("Book1.xlsx", sheet_name=0,template_path="template.docx", output_docx="entries.docx")
