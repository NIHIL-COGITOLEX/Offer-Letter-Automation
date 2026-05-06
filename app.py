from flask import (
    Flask,
    request,
    render_template,
    jsonify,
    session,
    redirect,
    url_for,
    send_file
)

from flask_cors import CORS
from docx import Document
from datetime import datetime
from werkzeug.utils import secure_filename
from authlib.integrations.flask_client import OAuth

import os
import tempfile
import subprocess
import platform
import urllib.parse

# =====================================================
# APP CONFIG
# =====================================================
app = Flask(__name__, template_folder="templates")
CORS(app)

app.secret_key = os.environ.get("SECRET_KEY", "supersecret")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

ALLOWED_EMAILS = ["hr@alfatza.com"]

# =====================================================
# TEMPLATES
# =====================================================
TEMPLATES = {
    "telecaller": os.path.join(BASE_DIR, "templates_docx", "telecaller.docx"),
    "team_leader": os.path.join(BASE_DIR, "templates_docx", "team_leader.docx"),
    "backend": os.path.join(BASE_DIR, "templates_docx", "backend.docx"),
    "hr": os.path.join(BASE_DIR, "templates_docx", "hr.docx"),
    "data_analyst": os.path.join(BASE_DIR, "templates_docx", "data_analyst.docx"),
}

BRANCHES = {
    "vashi": "Vashi Branch Address",
    "thane": "Thane Branch Address",
    "virar": "Virar Branch Address"
}

# =====================================================
# HELPERS
# =====================================================
def format_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")


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
    libreoffice = "soffice"
    if platform.system() == "Windows":
        libreoffice = r"C:\Program Files\LibreOffice\program\soffice.exe"

    subprocess.run([
        libreoffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        docx_path
    ], check=True)

    return os.path.join(
        output_dir,
        os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    )


def build_gmail_link(to, subject, body):
    base = "https://mail.google.com/mail/?view=cm&fs=1"
    params = {
        "to": to,
        "su": subject,
        "body": body
    }
    return base + "&" + urllib.parse.urlencode(params)


# =====================================================
# AUTH ROUTES
# =====================================================
@app.route("/login")
def login():
    return google.authorize_redirect(url_for("authorize", _external=True))


@app.route("/authorize")
def authorize():
    token = google.authorize_access_token()
    user = token.get("userinfo")

    if not user:
        return "Login failed"

    email = user.get("email")

    if email not in ALLOWED_EMAILS:
        return "Unauthorized"

    session["user"] = user
    return redirect("/")


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

    return render_template("index.html", user=session["user"])


# =====================================================
# GENERATE PDF + RETURN DOWNLOAD + GMAIL LINK
# =====================================================
@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "user" not in session:
            return jsonify({"error": "Unauthorized"}), 401

        data = request.get_json()

        required = [
            "name", "employee_code", "phone",
            "email", "address", "role",
            "branch", "salary", "joining"
        ]

        for field in required:
            if not data.get(field):
                return jsonify({"error": f"Missing {field}"}), 400

        template_path = TEMPLATES.get(data["role"])
        if not template_path or not os.path.exists(template_path):
            return jsonify({"error": "Template not found"}), 500

        doc = Document(template_path)

        values = {
            "{{name}}": data["name"],
            "{{employee_code}}": data["employee_code"],
            "{{phone}}": data["phone"],
            "{{address}}": data["address"],
            "{{branch_address}}": BRANCHES.get(data["branch"], ""),
            "{{salary}}": data["salary"],
            "{{joining}}": format_date(data["joining"]),
            "{{date}}": datetime.now().strftime("%d %B %Y")
        }

        replace_text(doc, values)

        temp_dir = tempfile.mkdtemp()
        safe_name = secure_filename(data["name"])

        docx_path = os.path.join(temp_dir, f"{safe_name}.docx")
        doc.save(docx_path)

        pdf_path = convert_to_pdf(docx_path, temp_dir)

        # ============================
        # GMAIL PREFILL
        # ============================
        branch_name = data["branch"].capitalize()

        subject = f"Issuance of Offer Letter – {branch_name} Branch"

        body = f"""Dear {data['name']},

Please find attached your Offer Letter.

Kindly sign and return a copy.

Regards,
HR Team
ALFA TZA LLP
"""

        gmail_link = build_gmail_link(
            data["email"],
            subject,
            body
        )

        return jsonify({
            "success": True,
            "download_url": f"/download/{safe_name}",
            "gmail_link": gmail_link
        })

    except subprocess.CalledProcessError:
        return jsonify({"error": "PDF conversion failed"}), 500

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# =====================================================
# DOWNLOAD ROUTE
# =====================================================
@app.route("/download/<filename>")
def download(filename):
    temp_dir = tempfile.gettempdir()

    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            if file.startswith(filename) and file.endswith(".pdf"):
                return send_file(os.path.join(root, file), as_attachment=True)

    return "File not found", 404


# =====================================================
# RUN
# =====================================================
if __name__ == "__main__":
    app.run(debug=True)
