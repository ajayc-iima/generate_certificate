from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document
from fpdf import FPDF

app = FastAPI()

@app.post("/generate-certificates/")
async def generate_certificates(excel_file: UploadFile = File(...), docx_template: UploadFile = File(...)):
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    output_dir = os.path.join(temp_dir, "certificates")
    os.makedirs(output_dir, exist_ok=True)
    
    # Save uploaded files
    excel_path = os.path.join(temp_dir, excel_file.filename)
    docx_path = os.path.join(temp_dir, docx_template.filename)
    with open(excel_path, "wb") as f:
        shutil.copyfileobj(excel_file.file, f)
    with open(docx_path, "wb") as f:
        shutil.copyfileobj(docx_template.file, f)
    
    # Read names from Excel
    df = pd.read_excel(excel_path)
    names = df.iloc[:, 0].str.strip().tolist()
    
    for name in names:
        doc = Document(docx_path)
        replace_text(doc, "{Name}", name)  # Using improved replacement function
        
        # Save DOCX certificate
        docx_output_path = os.path.join(output_dir, f"{name}_Certificate.docx")
        doc.save(docx_output_path)

        # Convert DOCX to PDF
        pdf_output_path = os.path.join(output_dir, f"{name}_Certificate.pdf")
        convert_docx_to_pdf(docx_output_path, pdf_output_path)
        
        # Remove .docx file
        os.remove(docx_output_path)
    
    # Zip all PDFs for download
    zip_path = os.path.join(temp_dir, "certificates.zip")
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_dir)
    
    return FileResponse(zip_path, filename="certificates.zip", media_type="application/zip")

def replace_text(doc, placeholder, replacement):
    """ Replaces text properly across runs without breaking formatting """
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                run.text = run.text.replace(placeholder, replacement)

    # Also replace text inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.text = run.text.replace(placeholder, replacement)

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert a Word document to a simple PDF using FPDF"""
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, txt=para.text, align='L')
    
    pdf.output(pdf_path)
