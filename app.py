from flask import (
    Flask,
    request,
    render_template,
    jsonify,
    session,
    redirect,
    url_for
)

from flask_cors import CORS
from docx import Document
from datetime import datetime
from werkzeug.utils import secure_filename
from authlib.integrations.flask_client import OAuth
from flask_mail import Mail, Message

import os
import tempfile
import subprocess
import platform

# =====================================================
# APP CONFIG
# =====================================================
app = Flask(__name__, template_folder="templates")
CORS(app)

app.secret_key = os.environ.get("SECRET_KEY")

# =====================================================
# GOOGLE OAUTH
# =====================================================
oauth = OAuth(app)

google = oauth.register(
    name="google",
    client_id=os.environ.get("GOOGLE_CLIENT_ID"),
    client_secret=os.environ.get("GOOGLE_CLIENT_SECRET"),
    server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
    client_kwargs={
        "scope": "openid email profile"
    }
)

# =====================================================
# MAIL CONFIG
# =====================================================
app.config["MAIL_SERVER"] = "smtp.gmail.com"
app.config["MAIL_PORT"] = 587
app.config["MAIL_USE_TLS"] = True
app.config["MAIL_USERNAME"] = os.environ.get("MAIL_USERNAME")
app.config["MAIL_PASSWORD"] = os.environ.get("MAIL_PASSWORD")
app.config["MAIL_DEFAULT_SENDER"] = os.environ.get("MAIL_USERNAME")

mail = Mail(app)

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

    pdf_name = os.path.splitext(
        os.path.basename(docx_path)
    )[0] + ".pdf"

    return os.path.join(output_dir, pdf_name)

# =====================================================
# LOGIN
# =====================================================
@app.route("/login")
def login():

    redirect_uri = url_for(
        "authorize",
        _external=True
    )

    return google.authorize_redirect(redirect_uri)

# =====================================================
# AUTHORIZE
# =====================================================
@app.route("/authorize")
def authorize():

    token = google.authorize_access_token()
    user = token.get("userinfo")

    if not user:
        return "Google login failed"

    email = user.get("email")

    # =========================================
    # ALLOWED EMAILS
    # =========================================
    allowed_emails = [
        "hr@alfatza.com"
    ]

    if email not in allowed_emails:
        return "Unauthorized Access"

    session["user"] = user

    return redirect("/")

# =====================================================
# LOGOUT
# =====================================================
@app.route("/logout")
def logout():

    session.clear()

    return redirect("/")

# =====================================================
# HOME
# =====================================================
@app.route("/")
def home():

    if "user" not in session:
        return redirect("/login")

    return render_template(
        "index.html",
        user=session["user"]
    )

# =====================================================
# GENERATE + SEND EMAIL
# =====================================================
@app.route("/generate", methods=["POST"])
def generate():

    try:

        if "user" not in session:
            return jsonify({
                "error": "Unauthorized"
            }), 401

        data = request.get_json()

        required_fields = [
            "name",
            "employee_code",
            "phone",
            "email",
            "address",
            "role",
            "branch",
            "salary",
            "joining"
        ]

        for field in required_fields:

            if not data.get(field):
                return jsonify({
                    "error": f"Missing field: {field}"
                }), 400

        role = data["role"]
        branch = data["branch"]

        template_path = TEMPLATES.get(role)

        if not template_path:
            return jsonify({
                "error": "Invalid role selected"
            }), 400

        if not os.path.exists(template_path):
            return jsonify({
                "error": "Template file not found"
            }), 500

        # =====================================================
        # LOAD DOCX
        # =====================================================
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

            docx_path = os.path.join(
                temp_dir,
                f"{safe_name}.docx"
            )

            doc.save(docx_path)

            pdf_path = convert_to_pdf(
                docx_path,
                temp_dir
            )

            # =====================================================
            # EMAIL CONTENT
            # =====================================================
            branch_name = branch.capitalize()

            subject = (
                f"Issuance of Offer Letter – "
                f"{branch_name} Branch"
            )

            body = f"""
Dear {data['name']},

Please find attached your formal Offer/Appointment Letter for your position at our {branch_name} Branch.

As you have already joined and are continuing your employment with us, this letter serves as the official documentation of your role, compensation, and terms of employment.

Kindly sign and return a copy for our records.

We look forward to your continued contribution and growth with the organization.

Warm regards,
Rashid Ali
H.R
ALFA TZA LLP
"""

            msg = Message(
                subject=subject,
                recipients=[data["email"]],
                body=body
            )

            with open(pdf_path, "rb") as f:

                msg.attach(
                    f"{safe_name}_offer_letter.pdf",
                    "application/pdf",
                    f.read()
                )

            mail.send(msg)

            return jsonify({
                "success": True,
                "message": "Offer Letter generated and emailed successfully."
            })

    
    except subprocess.CalledProcessError:

        return jsonify({
            "error": "PDF conversion failed."
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

    app.run(
        host="0.0.0.0",
        port=port
    )
