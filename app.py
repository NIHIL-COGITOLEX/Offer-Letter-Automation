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
from authlib.integrations.base_client.errors import MismatchingStateError

import os
import tempfile
import subprocess
import platform
import requests
import base64

# =====================================================
# APP CONFIG
# =====================================================
app = Flask(__name__, template_folder="templates")
CORS(app)

app.secret_key = os.environ.get("SECRET_KEY", "fallback_secret")

# 🔥 FIX FOR RENDER (IMPORTANT)
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="None"
)

# =====================================================
# GOOGLE OAUTH
# =====================================================
oauth = OAuth(app)

google = oauth.register(
    name="google",
    client_id=os.environ.get("GOOGLE_CLIENT_ID"),
    client_secret=os.environ.get("GOOGLE_CLIENT_SECRET"),
    server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
    client_kwargs={"scope": "openid email profile"}
)

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
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")

# =====================================================
# SALARY FORMAT
# =====================================================
def format_salary(value):
    amount = int(str(value).replace(",", "").strip())
    return f"{amount:,}"

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
    libreoffice_path = "soffice"  # Docker compatible

    subprocess.run([
        libreoffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        docx_path
    ], check=True)

    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    return os.path.join(output_dir, pdf_name)

# =====================================================
# LOGIN
# =====================================================
@app.route("/login")
def login():
    return google.authorize_redirect(url_for("authorize", _external=True))

# =====================================================
# AUTHORIZE (FIXED)
# =====================================================
@app.route("/authorize")
def authorize():
    try:
        token = google.authorize_access_token()
        user = token.get("userinfo")

        if not user:
            return "Google login failed"

        if user.get("email") not in ["hr@alfatza.com"]:
            return "Unauthorized Access"

        session["user"] = user
        return redirect("/")

    except MismatchingStateError:
        return redirect("/login")

# =====================================================
# HOME
# =====================================================
@app.route("/")
def home():
    if "user" not in session:
        return redirect("/login")
    return render_template("index.html", user=session["user"])

# =====================================================
# BREVO EMAIL FUNCTION (FIXED + DEBUG)
# =====================================================
def send_email_brevo(to_email, name, pdf_path, subject, body):

    print("Sending email to:", to_email)

    api_key = os.environ.get("BREVO_API_KEY")
    print("BREVO KEY:", api_key)

    with open(pdf_path, "rb") as f:
        encoded_file = base64.b64encode(f.read()).decode()

    response = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={
            "accept": "application/json",
            "api-key": api_key,
            "content-type": "application/json"
        },
        json={
            "sender": {
                "name": "ALFA TZA HR",
                "email": "hr@alfatza.com"
            },
            "to": [{"email": to_email, "name": name}],
            "subject": subject,
            "htmlContent": body.replace("\n", "<br>"),
            "attachment": [{
                "content": encoded_file,
                "name": f"{name}_offer_letter.pdf"
            }]
        }
    )

    print("BREVO STATUS:", response.status_code)
    print("BREVO RESPONSE:", response.text)

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

# =====================================================
# GENERATE + SEND EMAIL
# =====================================================
@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "user" not in session:
            return jsonify({"error": "Unauthorized"}), 401

        data = request.get_json()

        required = ["name","employee_code","phone","email","address","role","branch","salary","joining"]

        for field in required:
            if not data.get(field):
                return jsonify({"error": f"Missing field: {field}"}), 400

        template_path = TEMPLATES.get(data["role"])
        if not template_path or not os.path.exists(template_path):
            return jsonify({"error": "Template not found"}), 500

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

        with tempfile.TemporaryDirectory() as temp_dir:

            safe_name = secure_filename(data["name"])
            docx_path = os.path.join(temp_dir, f"{safe_name}.docx")

            doc.save(docx_path)
            pdf_path = convert_to_pdf(docx_path, temp_dir)

            branch_name = data["branch"].capitalize()

            subject = f"Issuance of Offer Letter – {branch_name} Branch"

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

            send_email_brevo(
                to_email=data["email"],
                name=data["name"],
                pdf_path=pdf_path,
                subject=subject,
                body=body
            )

        return jsonify({"success": True})

    except Exception as e:
        print("ERROR:", str(e))
        return jsonify({"error": str(e)}), 500

# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
