from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document

app = FastAPI()

# Serve Upload Form
@app.get("/", response_class=HTMLResponse)
async def serve_form():
    return """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Certificate Generator</title>
        <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            form { margin: 20px auto; width: 300px; }
            input, button { width: 100%; margin-top: 10px; padding: 8px; }
            #download { display: none; margin-top: 20px; }
        </style>
    </head>
    <body>
        <h2>Upload Excel & Word Template</h2>
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

                let response = await fetch("/generate-certificates/", {
                    method: "POST",
                    body: formData
                });

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
        
        # Iterate through all paragraphs and their runs
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                # Replace the placeholder with the actual name
                run.text = run.text.replace('{Name}', name)
        
        # Save the modified document
        docx_output_path = os.path.join(output_dir, f"{name}_Certificate.docx")
        doc.save(docx_output_path)
    
    # Zip all DOCXs for download
    zip_path = os.path.join(temp_dir, "certificates.zip")
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_dir)
    
    return FileResponse(zip_path, filename="certificates.zip", media_type="application/zip")
