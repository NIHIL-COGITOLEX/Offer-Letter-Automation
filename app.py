from flask import Flask, request, render_template, Response, jsonify
from flask_cors import CORS
from docx import Document
from datetime import datetime
from werkzeug.utils import secure_filename

import os
import tempfile
import subprocess
import platform

app = Flask(__name__, template_folder="templates")
CORS(app)

# =====================================================
# BASE DIR
# =====================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =====================================================
# ROLE -> DOCX TEMPLATE
# =====================================================
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
# NUMBER TO WORDS
# Because humans love writing numbers twice.
# =====================================================
ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six",
    "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
    "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"
]

TENS = [
    "", "", "Twenty", "Thirty", "Forty",
    "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
]


def two_digit_words(n):
    if n < 20:
        return ONES[n]
    return TENS[n // 10] + (" " + ONES[n % 10] if n % 10 else "")


def three_digit_words(n):
    word = ""

    if n >= 100:
        word += ONES[n // 100] + " Hundred"
        n %= 100
        if n:
            word += " "

    if n:
        word += two_digit_words(n)

    return word.strip()


def number_to_words_indian(n):
    if n == 0:
        return "Zero"

    parts = []

    crore = n // 10000000
    n %= 10000000

    lakh = n // 100000
    n %= 100000

    thousand = n // 1000
    n %= 1000

    hundred = n

    if crore:
        parts.append(two_digit_words(crore) + " Crore")

    if lakh:
        parts.append(two_digit_words(lakh) + " Lakh")

    if thousand:
        parts.append(two_digit_words(thousand) + " Thousand")

    if hundred:
        parts.append(three_digit_words(hundred))

    return " ".join(parts).strip()


def format_salary(value):
    amount = int(str(value).replace(",", "").strip())

    formatted_number = f"{amount:,}"
    words = number_to_words_indian(amount).lower()

    return f"{formatted_number} ({words})"


# =====================================================
# REPLACE TEXT INSIDE DOCX
# =====================================================
def replace_text(doc, values):

    for para in doc.paragraphs:
        for key, val in values.items():
            if key in para.text:
                para.text = para.text.replace(key, val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in values.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)


# =====================================================
# DOCX -> PDF
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
# HOME
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
            "employee_code",
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

        # Load DOCX Template
        doc = Document(template_path)

        joining_date = format_date(data["joining"])
        today_date = datetime.now().strftime("%d %B %Y")
        salary_text = format_salary(data["salary"])

        values = {
            "{{name}}": data["name"],
            "{{employee_code}}": data["employee_code"],
            "{{phone}}": data["phone"],
            "{{address}}": data["address"],
            "{{branch_address}}": BRANCHES.get(branch, ""),
            "{{salary}}": salary_text,
            "{{joining}}": joining_date,
            "{{date}}": today_date
        }

        replace_text(doc, values)

        with tempfile.TemporaryDirectory() as temp_dir:

            safe_name = secure_filename(data["name"])
            docx_path = os.path.join(temp_dir, f"{safe_name}.docx")

            doc.save(docx_path)

            pdf_path = convert_to_pdf(docx_path, temp_dir)

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            response = Response(pdf_bytes, mimetype="application/pdf")

            response.headers["Content-Disposition"] = (
                f'inline; filename="{safe_name}_offer_letter.pdf"'
            )

            return response

    except subprocess.CalledProcessError:
        return jsonify({
            "error": "PDF conversion failed. LibreOffice threw a tantrum."
        }), 500

    except ValueError:
        return jsonify({
            "error": "Invalid salary amount."
        }), 400

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
