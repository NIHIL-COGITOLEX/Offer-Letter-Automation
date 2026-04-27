from flask import Flask, request, render_template, Response, jsonify
from flask_cors import CORS
from docx import Document
from datetime import datetime
from werkzeug.utils import secure_filename

import os
import io
import tempfile
import subprocess
import platform

app = Flask(__name__, template_folder="templates")
CORS(app)

# =====================================================
# ROLE -> DOCX TEMPLATE
# =====================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES = {
    "telecaller": os.path.join(BASE_DIR, "templates_docx", "telecaller.docx"),
    "team_leader": os.path.join(BASE_DIR, "templates_docx", "team_leader.docx"),
    "backend": os.path.join(BASE_DIR, "templates_docx", "backend.docx"),
    "hr": os.path.join(BASE_DIR, "templates_docx", "hr.docx"),
    "data_analyst": os.path.join(BASE_DIR, "templates_docx", "data_analyst.docx"),
}

# =====================================================
# BRANCH ADDRESS MAP
# =====================================================
BRANCHES = {
    "vashi": "3rd Floor, Vashi Plaza, Alfa TZA LLP, D Wing-512, Plot No. 80/81, Sector 17, Navi Mumbai, Maharashtra 400703",
    "thane": "Alfa Tza LLP B-102 Rajdarshan CHS Ltd, Thane - 400602",
    "virar": "Virar Branch Address Here"
}

# =====================================================
# FORMAT DATE
# =====================================================
def format_date(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.strftime("%d %B %Y")


# =====================================================
# REPLACE TEXT INSIDE DOCX
# =====================================================
def replace_text(doc, values):
    # Paragraphs
    for para in doc.paragraphs:
        for key, val in values.items():
            if key in para.text:
                para.text = para.text.replace(key, val)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in values.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)


# =====================================================
# DOCX -> PDF USING LIBREOFFICE
# Works on Render / Linux / Cloud
# =====================================================
def convert_to_pdf(docx_path, output_dir):

    if platform.system() == "Windows":
        libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    else:
        libreoffice_path = "soffice"

    command = [
        libreoffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        docx_path
    ]

    subprocess.run(command, check=True)

    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    return os.path.join(output_dir, pdf_name)


# =====================================================
# HOME PAGE
# =====================================================
@app.route("/")
def home():
    return render_template("index.html")


# =====================================================
# GENERATE PDF
# =====================================================
@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json()

        required_fields = [
            "name",
            "phone",
            "address",
            "role",
            "branch",
            "salary",
            "joining"
        ]

        for field in required_fields:
            if not data.get(field):
                return jsonify({"error": f"Missing field: {field}"}), 400

        role = data["role"]
        branch = data["branch"]

        template_path = TEMPLATES.get(role)

        if not template_path:
            return jsonify({"error": "Invalid role selected"}), 400

        if not os.path.exists(template_path):
            return jsonify({"error": "Template file not found"}), 500

        # Load template
        doc = Document(template_path)

        joining_date = format_date(data["joining"])
        today_date = datetime.now().strftime("%d %B %Y")

        values = {
            "{{name}}": data["name"],
            "{{phone}}": data["phone"],
            "{{address}}": data["address"],
            "{{branch_address}}": BRANCHES.get(branch, ""),
            "{{salary}}": data["salary"],
            "{{joining}}": joining_date,
            "{{date}}": today_date
        }

        replace_text(doc, values)

        # temp folder because cloud hosting hates permanent storage
        with tempfile.TemporaryDirectory() as temp_dir:

            safe_name = secure_filename(data["name"])
            docx_path = os.path.join(temp_dir, f"{safe_name}.docx")

            # Save edited DOCX internally
            doc.save(docx_path)

            # Convert to PDF
            pdf_path = convert_to_pdf(docx_path, temp_dir)

            # Return PDF inline for browser preview
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            response = Response(pdf_bytes, mimetype="application/pdf")
            response.headers["Content-Disposition"] = (
                f'inline; filename="{safe_name}_offer_letter.pdf"'
            )

            return response

    except subprocess.CalledProcessError:
        return jsonify({
            "error": "PDF conversion failed. LibreOffice not installed on server."
        }), 500

    except Exception as e:
        return jsonify({
            "error": str(e)
        }), 500


# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
