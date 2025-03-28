from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import os
import tempfile
import shutil
from docx import Document

app = FastAPI(title="IIMA Certificate Generator", description="Generate certificates efficiently using an Excel file and a Word template.")

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
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            form { margin: 20px auto; width: 350px; padding: 20px; border: 1px solid #ddd; border-radius: 10px; background-color: #f9f9f9; }
            input, button { width: 100%; margin-top: 10px; padding: 10px; }
            button { background-color: #00457C; color: white; border: none; cursor: pointer; }
            button:hover { background-color: #003366; }
            #download { display: none; margin-top: 20px; padding: 10px; background-color: #28a745; color: white; text-decoration: none; border-radius: 5px; }
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
    output_dir = os.path.join(temp_dir, "certificates")
    os.makedirs(output_dir, exist_ok=True)

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

    for name in names:
        doc = Document(docx_path)
        replace_text_preserving_format(doc, "{Name}", name)  # ✅ Format-safe replacement
        doc.save(os.path.join(output_dir, f"{name}_Certificate.docx"))

    # Zip all certificates
    zip_path = os.path.join(temp_dir, "certificates.zip")
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_dir)

    return FileResponse(zip_path, filename="certificates.zip", media_type="application/zip")

def replace_text_preserving_format(doc, placeholder, replacement):
    """ Replaces {Name} without losing formatting & color """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)

    # Also replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, replacement)
