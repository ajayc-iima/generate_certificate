from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document
import subprocess

app = FastAPI()

@app.post("/generate-certificates/")
async def generate_certificates(excel_file: UploadFile = File(...), docx_template: UploadFile = File(...)):
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
        replace_text_in_doc(doc, "{Name}", name)  # Improved text replacement
        
        # Save DOCX certificate
        docx_output_path = os.path.join(output_dir, f"{name}_Certificate.docx")
        doc.save(docx_output_path)

        # Convert DOCX to PDF using LibreOffice
        pdf_output_path = os.path.join(output_dir, f"{name}_Certificate.pdf")
        convert_docx_to_pdf(docx_output_path, pdf_output_path)
        
        # Remove .docx file after conversion
        os.remove(docx_output_path)
    
    # Zip all PDFs for download
    zip_path = os.path.join(temp_dir, "certificates.zip")
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_dir)
    
    return FileResponse(zip_path, filename="certificates.zip", media_type="application/zip")


def replace_text_in_doc(doc, placeholder, replacement):
    """ Replaces text properly across runs without breaking formatting """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)
    
    # Replace text inside tables as well
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, replacement)


def convert_docx_to_pdf(docx_path, pdf_path):
    """ Converts DOCX to PDF using LibreOffice to retain formatting """
    output_dir = os.path.dirname(pdf_path)
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir])

