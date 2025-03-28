from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import pandas as pd
import os
import tempfile
import shutil
from docx import Document

app = FastAPI()

# CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/generate-certificates/")
async def generate_certificates(
    excel_file: UploadFile = File(...), 
    docx_template: UploadFile = File(...),
    program_type: str = Form(...),
    program_name: str = Form(...)
):
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
        # Create a deep copy of the original document to preserve all formatting
        doc = Document(docx_path)
        
        # Iterate through all paragraphs to find and replace the name
        for paragraph in doc.paragraphs:
            # Check if the paragraph contains the placeholder
            if '{Name}' in paragraph.text:
                # Preserve the original paragraph formatting
                paragraph.text = paragraph.text.replace('{Name}', name)
        
        # Additional replacements for program details
        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace('{ProgramType}', program_type)
            paragraph.text = paragraph.text.replace('{ProgramName}', program_name)
        
        # Save the modified document
        docx_output_path = os.path.join(output_dir, f"{name}_Certificate.docx")
        doc.save(docx_output_path)
    
    # Zip all DOCXs for download
    zip_path = os.path.join(temp_dir, "certificates.zip")
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_dir)
    
    return FileResponse(zip_path, filename=f"{program_name}_Certificates.zip", media_type="application/zip")

# Optional: Serve React frontend if needed
app.mount("/", StaticFiles(directory="frontend/build", html=True), name="static")
