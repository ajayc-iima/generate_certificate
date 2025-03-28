from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import RGBColor

app = FastAPI(title="IIMA Certificate Generator", description="Generate certificates efficiently using an Excel file and a Word template.")

def replace_text_preserving_format(doc, placeholder, replacement):
    """ Replaces {Name} while keeping formatting and color intact """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                text_parts = run.text.split(placeholder)
                run.text = text_parts[0]  # Keep text before {Name}

                # Create a new run with the same formatting
                new_run = paragraph.add_run(replacement)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.color.rgb = run.font.color.rgb if run.font.color else RGBColor(0, 0, 0)

                if len(text_parts) > 1:
                    run.text += text_parts[1]  # Append text after {Name}

    # Also replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            text_parts = run.text.split(placeholder)
                            run.text = text_parts[0]

                            new_run = paragraph.add_run(replacement)
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            new_run.font.color.rgb = run.font.color.rgb if run.font.color else RGBColor(0, 0, 0)

                            if len(text_parts) > 1:
                                run.text += text_parts[1]

def add_page_break(doc):
    """ Inserts a page break into the document """
    page_break = OxmlElement("w:br")
    page_break.set("w:type", "page")
    doc.add_paragraph()._element.append(page_break)

@app.post("/generate-certificates/")
async def generate_certificates(excel_file: UploadFile = File(...), docx_template: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()

    # Save uploaded files
    excel_path = os.path.join(temp_dir, excel_file.filename)
    docx_path = os.path.join(temp_dir, docx_template.filename)

    with open(excel_path, "wb") as f:
        shutil.copyfileobj(excel_file.file, f)
    with open(docx_path, "wb") as f:
        shutil.copyfileobj(docx_template.file, f)

    # Read Excel file
    df = pd.read_excel(excel_path)
    names = df.iloc[:, 0].dropna().str.strip().tolist()

    # Create final document
    final_doc = Document()

    for idx, name in enumerate(names):
        temp_doc = Document(docx_path)  # Load fresh template for each name
        replace_text_preserving_format(temp_doc, "{Name}", name)  # Replace name safely

        # Append modified template to final doc
        for element in temp_doc.element.body:
            final_doc.element.body.append(element)

        # Add page break except for the last certificate
        if idx < len(names) - 1:
            add_page_break(final_doc)

    # Save final document
    final_doc_path = os.path.join(temp_dir, "Final_Certificates.docx")
    final_doc.save(final_doc_path)

    return FileResponse(final_doc_path, filename="Final_Certificates.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
