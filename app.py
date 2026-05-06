from flask import Flask, request, send_file, jsonify, render_template
from docx import Document
from datetime import datetime
from werkzeug.utils import secure_filename
from urllib.parse import quote

import os
import tempfile
import subprocess

app = Flask(__name__, template_folder="templates")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES = {
    "telecaller": os.path.join(BASE_DIR, "templates_docx", "telecaller.docx"),
    "team_leader": os.path.join(BASE_DIR, "templates_docx", "team_leader.docx"),
    "backend": os.path.join(BASE_DIR, "templates_docx", "backend.docx"),
    "hr": os.path.join(BASE_DIR, "templates_docx", "hr.docx"),
    "data_analyst": os.path.join(BASE_DIR, "templates_docx", "data_analyst.docx"),
}

BRANCHES = {
    "vashi": "Vashi Address",
    "thane": "Thane Address",
    "virar": "Virar Address"
}

# =============================
# HELPERS
# =============================
def format_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")

def format_salary(val):
    return f"{int(val):,}"

def replace_text(doc, values):
    for para in doc.paragraphs:
        for k, v in values.items():
            if k in para.text:
                para.text = para.text.replace(k, v)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for k, v in values.items():
                    if k in cell.text:
                        cell.text = cell.text.replace(k, v)

def convert_to_pdf(docx_path, output_dir):
    result = subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ], capture_output=True, text=True)

    if result.returncode != 0:
        raise Exception(result.stderr)

    return os.path.join(
        output_dir,
        os.path.basename(docx_path).replace(".docx", ".pdf")
    )

# =============================
# ROUTES
# =============================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json()

        required = [
            "name","employee_code","phone",
            "address","role","branch",
            "salary","joining"
        ]

        for f in required:
            if not data.get(f):
                return jsonify({"error": f"Missing {f}"}), 400

        template_path = TEMPLATES.get(data["role"])
        if not template_path:
            return jsonify({"error": "Invalid role"}), 400

        doc = Document(template_path)

        replace_text(doc, {
            "{{name}}": data["name"],
            "{{employee_code}}": data["employee_code"],
            "{{phone}}": data["phone"],
            "{{address}}": data["address"],
            "{{branch_address}}": BRANCHES.get(data["branch"], ""),
            "{{salary}}": format_salary(data["salary"]),
            "{{joining}}": format_date(data["joining"]),
            "{{date}}": datetime.now().strftime("%d %B %Y")
        })

        temp_dir = tempfile.mkdtemp()
        filename = secure_filename(data["name"])

        docx_path = os.path.join(temp_dir, f"{filename}.docx")
        doc.save(docx_path)

        pdf_path = convert_to_pdf(docx_path, temp_dir)

        # MAILTO FIXED
        subject = quote(f"Issuance of Offer Letter – {data['branch'].capitalize()} Branch")

        body = quote(f"""Dear {data['name']},

Please find your Offer Letter attached.

Regards,
HR Team""")

        mailto = f"mailto:?subject={subject}&body={body}"

        response = send_file(
            pdf_path,
            as_attachment=True,
            download_name=f"{filename}.pdf",
            mimetype="application/pdf"
        )

        response.headers["X-Mailto"] = mailto

        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =============================
# RUN
# =============================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
