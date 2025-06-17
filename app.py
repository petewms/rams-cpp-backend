from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
import tempfile
import os
import pandas as pd

app = Flask(__name__)
CORS(app)


@app.route("/generate-cpp", methods=["POST"])
def generate_cpp():
    file = request.files["quote"]
    if not file:
        return "No file uploaded", 400

    # Read Excel quote
    xls = pd.ExcelFile(file)
    df = xls.parse(xls.sheet_names[0], header=None)

    # Extract fields
    job_number = df.iloc[1, 10] if not pd.isna(df.iloc[1, 10]) else "TBC"
    quote_number = df.iloc[2, 10] if not pd.isna(df.iloc[2, 10]) else "TBC"
    quoted_by = df.iloc[3, 10] if not pd.isna(df.iloc[3, 10]) else "TBC"
    site_address = df.iloc[10, 1] if not pd.isna(df.iloc[10, 1]) else "Site Address"

    # Extract scope from SOR descriptions starting from row 15
    scope_lines = []
    for i in range(15, len(df)):
        desc = df.iloc[i, 2]
        if isinstance(desc, str) and desc.strip():
            scope_lines.append(f"- {desc.strip()}")
    scope_text = "\n".join(scope_lines) if scope_lines else "Scope of works"

    # Load template
    template_path = os.path.join("templates", "cpp_template.docx")
    doc = Document(template_path)

    # Replace static text
    replacements = {
        "{{SiteAddress}}": site_address,
        "{{ScopeOfWorks}}": scope_text,
        "{{Client}}": "LiveWest",
        "{{JobNumber}}": job_number
    }

    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)

    # Save temporary file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return send_file(tmp.name, as_attachment=True, download_name="CPP_Filled.docx")
