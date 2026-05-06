from flask import (
    Flask, request, render_template,
    jsonify, session, redirect,
    url_for, send_file
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

# =====================================================
# APP CONFIG
# =====================================================
app = Flask(__name__, template_folder="templates")
CORS(app)

app.secret_key = os.environ.get("SECRET_KEY", "fallback_secret")

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

# =====================================================
# HELPERS
# =====================================================
def format_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")

def format_salary(value):
    return f"{int(value):,}"

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
    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ], check=True)

    return os.path.join(
        output_dir,
        os.path.basename(docx_path).replace(".docx", ".pdf")
    )

# =====================================================
# AUTH
# =====================================================
@app.route("/login")
def login():
    return google.authorize_redirect(url_for("authorize", _external=True))

@app.route("/authorize")
def authorize():
    try:
        token = google.authorize_access_token()
        user = token.get("userinfo")

        if user.get("email") != "hr@alfatza.com":
            return "Unauthorized"

        session["user"] = user
        return redirect("/")

    except MismatchingStateError:
        return redirect("/login")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

@app.route("/")
def home():
    if "user" not in session:
        return redirect("/login")
    return render_template("index.html", user=session["user"])

# =====================================================
# GENERATE
# =====================================================
@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "user" not in session:
            return jsonify({"error": "Unauthorized"}), 401

        data = request.get_json()

        required = ["name","employee_code","phone","email","address","role","branch","salary","joining"]

        for f in required:
            if not data.get(f):
                return jsonify({"error": f"Missing {f}"}), 400

        template_path = TEMPLATES.get(data["role"])
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

        # 🔥 STORE PATH IN SESSION
        session["pdf_path"] = pdf_path

        branch_name = data["branch"].capitalize()

        subject = f"Issuance of Offer Letter – {branch_name} Branch"

        body = f"""Dear {data['name']},

Please find attached your Offer Letter.

Regards,
HR Team"""

        # 🔥 MAIL LINK
        mailto_link = f"mailto:{data['email']}?subject={subject}&body={body}"

        return jsonify({
            "success": True,
            "download_url": "/download",
            "mailto": mailto_link
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =====================================================
# DOWNLOAD PDF
# =====================================================
@app.route("/download")
def download():
    path = session.get("pdf_path")

    if not path or not os.path.exists(path):
        return "File not found", 404

    return send_file(path, as_attachment=True)

# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
