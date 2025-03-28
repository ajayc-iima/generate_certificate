from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement

app = FastAPI(title="IIMA Certificate Generator", description="Generate a single certificate document efficiently using an Excel file and a Word template.")

@app.get("/", response_class=HTMLResponse)
async def serve_form():
    return """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>IIMA Certificate Generator</title>
        <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f4f4f9; }
            form { margin: 20px auto; width: 400px; padding: 25px; border-radius: 10px; background: white; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); }
            input, button { width: 100%; margin-top: 10px; padding: 12px; border-radius: 5px; border: 1px solid #ccc; }
            button { background-color: #00457C; color: white; border: none; font-size: 16px; cursor: pointer; transition: 0.3s; }
            button:hover { background-color: #002c5a; }
            #download { display: none; margin-top: 20px; padding: 10px; background-color: #28a745; color: white; text-decoration: none; border-radius: 5px; font-size: 16px; }
        </style>
    </head>
    <body>
        <h2>IIMA Certificate Generator</h2>
        <p>Upload an Excel file with participant names and a Word template with '{Name}' placeholder.</p>
        <form id="upload-form">
            <input type="file" id="excel" accept=".xlsx" required><br>
            <input type="file" id="docx" accept=".docx" required><br>
            <button type="submit">Generate Certificates</button>
        </form>
        <p id="message"></p>
        <a id="download" href="#" download>Download Certificates</a>

        <script>
            document.getElementById("upload-form").addEventListener("submit", async function(event) {
                event.preventDefault();
                let formData = new FormData();
                formData.append("excel_file", document.getElementById("excel").files[0]);
                formData.append("docx_template", document.getElementById("docx").files[0]);
                document.getElementById("message").innerText = "Processing...";
                let response = await fetch("/generate-certificates/", { method: "POST", body: formData });
                if (response.ok) {
                    let blob = await response.blob();
                    let url = window.URL.createObjectURL(blob);
                    let downloadLink = document.getElementById("download");
                    downloadLink.href = url;
                    downloadLink.style.display = "block";
                    document.getElementById("message").innerText = "✅ Certificates generated! Download below:";
                } else {
                    document.getElementById("message").innerText = "❌ Error generating certificates!";
                }
            });
        </script>
    </body>
    </html>
    """

@app.post("/generate-certificates/")
async def generate_certificates(excel_file: UploadFile = File(...), docx_template: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    output_path = os.path.join(temp_dir, "Final_Certificates.docx")

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

    # Load template once
    master_doc = Document()
    
    for i, name in enumerate(names):
        doc = Document(docx_path)
        replace_text_preserving_format(doc, "{Name}", name)

        if i > 0:
            master_doc.add_page_break()

        for element in doc.element.body:
            master_doc.element.body.append(element)

    master_doc.save(output_path)

    return FileResponse(output_path, filename="Final_Certificates.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

def replace_text_preserving_format(doc, placeholder, replacement):
    """ Replaces {Name} while preserving text formatting & color """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, replacement)
