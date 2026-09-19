import os
import sys
import json
import io
import argparse
import datetime
from datetime import timedelta
import zipfile
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from werkzeug.security import check_password_hash, generate_password_hash
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

import pandas as pd
import db
from services.excel_reader import (
    read_weekly_observation_excel,
    read_names_locations_excel,
    read_weekly_activity_excel,
    format_date_for_jsl
)
from services.mapping import process_and_map_observations, split_into_pages
from services.or_generator import generate_observation_register_doc
from services.jsl_generator import generate_jsl_log_excel
from services.weekly_report_generator import generate_weekly_activity_report
from services.signage_generator import generate_composite_signage

# Base paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")
SIGNAGE_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "signages")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SAMPLE_DIR, exist_ok=True)
os.makedirs(SIGNAGE_OUTPUT_DIR, exist_ok=True)

# Initialize SQLite database
db.init_db()

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = 'heisco-venkata-404510-super-secret-2024'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)
app.config['SESSION_PERMANENT'] = True
app.config.update(
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_HTTPONLY=True,
)

@app.route('/public/<path:filename>')
def serve_public(filename):
    return send_file(os.path.join(BASE_DIR, 'public', filename))

@app.route('/assets/<path:filename>')
def serve_assets(filename):
    if os.path.exists(os.path.join(BASE_DIR, 'public', 'assets', filename)):
        return send_file(os.path.join(BASE_DIR, 'public', 'assets', filename))
    return send_file(os.path.join(BASE_DIR, 'static', 'assets', filename))

@app.context_processor
def inject_user_context():
    username = session.get("username") or session.get("user") or "admin"
    role = session.get("role", "admin").upper()
    
    if username == "admin":
        display_name = "VENKATA NAGA PRASAD"
        employee_id = "404510"
        role = "ADMIN"
    else:
        display_name = session.get("display_name") or username.title()
        employee_id = session.get("employee_id") or "—"
        
    return {
        "current_user_name": display_name,
        "current_employee_id": employee_id,
        "current_user_role": role
    }

# =============================================================
# AUTHENTICATION & ROLE-BASED ACCESS DECORATORS
# =============================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Support fallback url param for iframe/preview environments
        url_user = request.args.get("user")
        if url_user:
            u_obj = db.get_user_by_username(url_user)
            if u_obj:
                session.permanent = True
                session["user"] = u_obj["username"]
                session["username"] = u_obj["username"]
                session["user_id"] = u_obj["id"]
                session["role"] = u_obj["role"]
                dn = u_obj["display_name"] if "display_name" in u_obj.keys() and u_obj["display_name"] else ("VENKATA NAGA PRASAD" if u_obj["username"] == "admin" else u_obj["username"].title())
                eid = u_obj["employee_id"] if "employee_id" in u_obj.keys() and u_obj["employee_id"] else ("404510" if u_obj["username"] == "admin" else "—")
                session["display_name"] = dn
                session["employee_id"] = eid
                
        if "user" not in session and "user_id" not in session and "username" not in session:
            flash("Please sign in to access Central HSE Operations.", "error")
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        url_user = request.args.get("user")
        if url_user:
            u_obj = db.get_user_by_username(url_user)
            if u_obj:
                session.permanent = True
                session["user"] = u_obj["username"]
                session["username"] = u_obj["username"]
                session["user_id"] = u_obj["id"]
                session["role"] = u_obj["role"]
                dn = u_obj["display_name"] if "display_name" in u_obj.keys() and u_obj["display_name"] else ("VENKATA NAGA PRASAD" if u_obj["username"] == "admin" else u_obj["username"].title())
                eid = u_obj["employee_id"] if "employee_id" in u_obj.keys() and u_obj["employee_id"] else ("404510" if u_obj["username"] == "admin" else "—")
                session["display_name"] = dn
                session["employee_id"] = eid

        if "user" not in session and "user_id" not in session and "username" not in session:
            flash("Please sign in to access this page.", "error")
            return redirect(url_for("login_page"))
        if session.get("role") != "admin":
            flash("Access Denied: Administrator privilege required.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated_function

# Temporary storage file for active session state
STATE_FILE = os.path.join(UPLOAD_DIR, "current_session.json")

def save_current_state(data: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)

def load_current_state() -> dict:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def ensure_sample_files_exist():
    """Ensure sample data files are generated and available for quick demonstration."""
    weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
    names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
    activity_sample = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")
    
    if not os.path.exists(weekly_sample) or not os.path.exists(names_sample) or not os.path.exists(activity_sample):
        from sample_data.generate_samples import generate_sample_files
        generate_sample_files()

# =============================================================
# AUTHENTICATION & USER MANAGEMENT ROUTES
# =============================================================

@app.route("/check-users")
def check_users():
    import sqlite3
    conn = sqlite3.connect("users.db")
    cur = conn.cursor()
    cur.execute("SELECT username, role FROM users")
    users = cur.fetchall()
    conn.close()
    return str(users)

@app.route("/login", methods=["GET", "POST"])
def login_page():
    """Login page with SQLite authentication."""
    if "user" in session or "user_id" in session or "username" in session:
        return redirect(url_for("dashboard"))
        
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        
        user = db.get_user_by_username(username)
        print(f"[LOGIN ATTEMPT] Username: '{username}', Found in DB: {user is not None}")
        
        is_valid = False
        if user:
            stored_pw = user["password"]
            try:
                is_valid = check_password_hash(stored_pw, password)
            except Exception as e:
                print(f"[LOGIN ERROR] check_password_hash exception: {e}")
                is_valid = False
            if not is_valid and stored_pw == password:
                is_valid = True

        print(f"[LOGIN RESULT] Username: '{username}', Password Check Result: {is_valid}")
        
        if is_valid:
            session.permanent = True
            session["user"] = username
            session["username"] = user["username"]
            session["user_id"] = user["id"]
            session["role"] = user["role"]
            dn = "VENKATA NAGA PRASAD" if user["username"] == "admin" else (user["display_name"] if "display_name" in user.keys() and user["display_name"] else username.title())
            eid = "404510" if user["username"] == "admin" else (user["employee_id"] if "employee_id" in user.keys() and user["employee_id"] else "—")
            session["display_name"] = dn
            session["employee_id"] = eid
            flash(f"Welcome, {dn}!", "success")
            return redirect(url_for("dashboard", user=username))
        else:
            flash("Invalid credentials", "error")
            
    return render_template("login.html")

@app.route("/reset-admin-now")
def reset_admin_now():
    import sqlite3
    import os
    from werkzeug.security import generate_password_hash
    
    db_paths = ['users.db', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'users.db')]
    for path in set(db_paths):
        conn = sqlite3.connect(path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                display_name TEXT DEFAULT '',
                employee_id TEXT DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        try:
            cur.execute("ALTER TABLE users ADD COLUMN display_name TEXT DEFAULT ''")
        except Exception:
            pass
        try:
            cur.execute("ALTER TABLE users ADD COLUMN employee_id TEXT DEFAULT ''")
        except Exception:
            pass
        conn.commit()
        hashed = generate_password_hash('heisco123')
        cur.execute("DELETE FROM users WHERE username='admin'")
        cur.execute(
            "INSERT INTO users (username, password, role, display_name, employee_id) VALUES (?, ?, ?, ?, ?)",
            ('admin', hashed, 'admin', 'VENKATA NAGA PRASAD', '404510')
        )
        conn.commit()
        conn.close()
    return "DONE. Now you can login with admin / heisco123. Go to /login"

@app.route("/logout")
def logout():
    """Clears user session and redirects to login page."""
    session.clear()
    flash("You have been successfully signed out.", "success")
    return redirect(url_for("login_page"))

@app.route("/reset_password", methods=["GET", "POST"])
def reset_password():
    """Password reset page with database update and validation."""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        new_password = request.form.get("new_password", "").strip()
        confirm_password = request.form.get("confirm_password", "").strip()
        
        if not username:
            flash("Username is required.", "error")
            return render_template("reset_password.html")
            
        if new_password != confirm_password:
            flash("Passwords do not match. Please re-enter.", "error")
            return render_template("reset_password.html")
            
        if len(new_password) < 4:
            flash("Password must be at least 4 characters long.", "error")
            return render_template("reset_password.html")
            
        user = db.get_user_by_username(username)
        if not user:
            flash(f"User '{username}' was not found in users.db.", "error")
            return render_template("reset_password.html")
            
        if db.update_user_password(username, new_password):
            flash("Password updated successfully! You can now sign in.", "success")
            return redirect(url_for("login_page"))
        else:
            flash("Failed to update password in SQLite database.", "error")
            
    return render_template("reset_password.html")

@app.route("/register", methods=["GET", "POST"])
@admin_required
def register():
    """Admin-only User Registration & Management page."""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "officer").strip()
        
        success, msg = db.create_new_user(username, password, role)
        flash(msg, "success" if success else "error")
        return redirect(url_for("register"))
        
    users = db.get_all_users()
    return render_template("register.html", users=users)

@app.route("/delete_user/<int:user_id>", methods=["POST"])
@admin_required
def delete_user(user_id):
    """Admin-only route to delete a user from SQLite database."""
    if db.delete_user_by_id(user_id):
        flash("User removed from database successfully.", "success")
    else:
        flash("Cannot delete primary administrator account.", "error")
    return redirect(url_for("register"))

@app.route("/charts")
@login_required
def charts_page():
    """HSE Performance & KPI Chart Generator UI."""
    return render_template("charts.html")

# =============================================================
# FRONT-END OPERATIONAL ROUTES
# =============================================================

@app.route("/")
@app.route("/dashboard")
@login_required
def dashboard():
    """Main Dashboard Page with HEISCO – ENPPI – Aramco Corporate Theme."""
    return render_template("dashboard.html")

@app.route("/signage", methods=["GET", "POST"])
@login_required
def signage_page():
    """PPE Signage Generator UI."""
    signage_name = request.form.get("signage_name", "PPE FREE ZONE").strip() if request.method == "POST" else request.args.get("signage_name", "PPE FREE ZONE")
    if not signage_name:
        signage_name = "PPE FREE ZONE"
    return render_template("signage.html", signage_name=signage_name)

@app.route("/observations", methods=["GET", "POST"])
@login_required
def observations_page():
    """Observations Report Generator Page."""
    state = load_current_state()
    return render_template("observations.html", state=state)

@app.route("/weekly_report", methods=["GET", "POST"])
@login_required
def weekly_report_page():
    """Weekly HSE Activity Report Generator Page."""
    state = load_current_state()
    return render_template("weekly_report.html", state=state)

@app.route("/jsl_log", methods=["GET", "POST"])
@login_required
def jsl_log_page():
    """JSL Log Sheet Generator Page."""
    state = load_current_state()
    return render_template("jsl_log.html", state=state)

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload_page():
    """Alias for observations upload page."""
    return redirect(url_for("observations_page"))

# =============================================================
# 1-CLICK SAMPLE DEMONSTRATION LOADER
# =============================================================

@app.route("/use-sample", methods=["POST"])
@login_required
def use_sample():
    """Immediately load pre-configured sample files and go to preview."""
    ensure_sample_files_exist()
    weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
    names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
    activity_sample = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")

    try:
        raw_obs = read_weekly_observation_excel(weekly_sample)
        names_map = read_names_locations_excel(names_sample)
        processed = process_and_map_observations(raw_obs, names_map)
        activity_data = read_weekly_activity_excel(activity_sample)

        save_current_state({
            "processed": processed,
            "names_mapping": names_map,
            "weekly_filename": "Weekly_Observation_Sample.xlsx",
            "names_filename": "Names_Locations_Sample.xlsx",
            "activity_filename": "Weekly_HSE_Activity_Sample.xlsx",
            "activity_data": activity_data
        })

        return redirect(url_for("preview_page"))
    except Exception as e:
        return render_template("observations.html", error=f"Failed to load sample data: {str(e)}")

@app.route("/preview")
@login_required
def preview_page():
    state = load_current_state()
    processed = state.get("processed", [])
    activity_data = state.get("activity_data", {})
    
    if not processed:
        # Load sample data if state is empty so user never sees a broken blank page
        ensure_sample_files_exist()
        weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
        names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
        activity_sample = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")
        raw_obs = read_weekly_observation_excel(weekly_sample)
        names_map = read_names_locations_excel(names_sample)
        processed = process_and_map_observations(raw_obs, names_map)
        activity_data = read_weekly_activity_excel(activity_sample)
        state = {
            "processed": processed,
            "names_mapping": names_map,
            "weekly_filename": "Weekly_Observation_Sample.xlsx",
            "names_filename": "Names_Locations_Sample.xlsx",
            "activity_filename": "Weekly_HSE_Activity_Sample.xlsx",
            "activity_data": activity_data
        }
        save_current_state(state)

    total_records = len(processed)
    total_pages = (total_records + 2) // 3
    
    locations = sorted(list({item.get("area", "") for item in processed if item.get("area")}))
    dates = [item.get("observation_date", "") for item in processed if item.get("observation_date")]
    if dates:
        date_range = f"{dates[0]} to {dates[-1]}" if len(dates) > 1 else dates[0]
        default_reviewed_date = "05/09/2026"
    else:
        date_range = "N/A"
        default_reviewed_date = datetime.date.today().strftime("%d/%m/%Y")

    period_start = activity_data.get("period_start", "05-09-2026") if activity_data else "05-09-2026"
    period_end = activity_data.get("period_end", "10-09-2026") if activity_data else "10-09-2026"
    week_no = activity_data.get("week_no", "36") if activity_data else "36"

    preview_records = processed[:25]

    return render_template(
        "preview.html",
        preview_records=preview_records,
        total_records=total_records,
        total_pages=total_pages,
        unique_locations=locations,
        date_range=date_range,
        default_reviewed_date=default_reviewed_date,
        period_start=period_start,
        period_end=period_end,
        week_no=week_no,
        activity_filename=state.get("activity_filename", "Weekly HSE Activity File Loaded"),
        activity_data=activity_data
    )

# =============================================================
# REQUIRED BACKEND FUNCTIONS
# =============================================================

@app.route("/generate_signage", methods=["POST"])
@login_required
def generate_signage():
    """
    PPE Signage Generator (PPT)
    Backend uses exact paths:
    enppi_logo = "static/logos/enppi.png"
    heisco_logo = "static/logos/heisco.png"
    aramco_logo = "static/logos/aramco.png"
    pictogram_path = "static/pictograms/pictogram.png"
    """
    signage_name = request.form.get("signage_name", "PPE FREE ZONE").strip()
    if not signage_name:
        signage_name = "PPE FREE ZONE"

    enppi_logo = "static/logos/enppi.png"
    heisco_logo = "static/logos/heisco.png"
    aramco_logo = "static/logos/aramco.png"
    pictogram_path = "static/pictograms/pictogram.png"

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank slide

    # 1. TOP COMMON HEADER (header_final.png)
    header_final = "public/assets/header_final.png"
    if not os.path.exists(header_final):
        header_final = "static/assets/header_final.png"
    if not os.path.exists(header_final):
        header_final = "static/header_final.png"

    if os.path.exists(header_final):
        # Add header_final.png at the very top as common header across width
        slide.shapes.add_picture(header_final, Inches(0.5), Inches(0.2), Inches(9.0), Inches(0.95))
    else:
        if os.path.exists(enppi_logo):
            slide.shapes.add_picture(enppi_logo, Inches(0.5), Inches(0.3), Inches(2.5), Inches(0.8))
        if os.path.exists(heisco_logo):
            slide.shapes.add_picture(heisco_logo, Inches(3.2), Inches(0.3), Inches(2.5), Inches(0.8))
        if os.path.exists(aramco_logo):
            slide.shapes.add_picture(aramco_logo, Inches(6.5), Inches(0.3), Inches(2.5), Inches(0.8))

    # 2. HEADER BAR (Solid Yellow Caution Header)
    header = slide.shapes.add_shape(1, Inches(0.5), Inches(1.3), Inches(9), Inches(1))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(255, 215, 0)  # Yellow
    tf = header.text_frame
    p = tf.add_paragraph()
    p.text = "CAUTION"
    p.font.bold = True
    p.font.size = Pt(72)
    p.font.color.rgb = RGBColor(0, 0, 0)

    # 3. PICTOGRAM
    norm = signage_name.upper()
    chosen_pictogram = pictogram_path
    req_pic = request.form.get("pictogram")
    if req_pic and os.path.exists(req_pic):
        chosen_pictogram = req_pic
    elif ("SLIP" in norm or "TRIP" in norm or "FALL" in norm) and os.path.exists("icons/slip_trip_fall.png"):
        chosen_pictogram = "icons/slip_trip_fall.png"
    elif ("HARD HAT" in norm or "HELMET" in norm) and os.path.exists("icons/hard_hat.png"):
        chosen_pictogram = "icons/hard_hat.png"
    elif ("GLASSES" in norm or "EYE" in norm) and os.path.exists("icons/safety_glasses.png"):
        chosen_pictogram = "icons/safety_glasses.png"
    elif ("SMOKING" in norm) and os.path.exists("icons/no_smoking.png"):
        chosen_pictogram = "icons/no_smoking.png"

    if os.path.exists(chosen_pictogram):
        slide.shapes.add_picture(chosen_pictogram, Inches(0.5), Inches(2.5), Inches(3.5), Inches(3.5))

    # 4. MULTILINGUAL TEXT (4-language translations)
    if "SLIP" in norm or "TRIP" in norm or "FALL" in norm:
        translations = [
            signage_name.upper(),
            "خطر الانزلاق والتعثر والسقوط",
            "फिसलने, ठोकर लगने और गिरने का खतरा",
            "پھسلنے ، ٹھوکر لگنے اور گرنے کا خطرہ"
        ]
    elif "HARD HAT" in norm:
        translations = [
            signage_name.upper(),
            "يجب ارتداء خوذة السلامة",
            "सुरक्षा हेलमेट आवश्यक है",
            "حفاظتی ہیلمٹ درکار ہے"
        ]
    elif "SAFETY GLASSES" in norm:
        translations = [
            signage_name.upper(),
            "يجب ارتداء نظارات السلامة",
            "सुरक्षा चश्मा आवश्यक है",
            "حفاظتی چشمہ درکار ہے"
        ]
    elif "SMOKING" in norm:
        translations = [
            signage_name.upper(),
            "ممنوع التدخين",
            "धूम्रपान निषेध",
            "تمباکو نوشی منع ہے"
        ]
    elif "MANDATORY" in norm:
        translations = [
            signage_name.upper(),
            "منطقة ارتداء معدات الوقاية الشخصية إلزامي",
            "अनिवार्य पीपीई क्षेत्र",
            "لازمی پی پی ای زون"
        ]
    else:
        translations = [
            signage_name.upper(),
            "منطقة خالية من معدات الحماية الشخصية",
            "पीपीई मुक्त क्षेत्र",
            "پی پی ای فری زون"
        ]

    text_box = slide.shapes.add_textbox(Inches(4.5), Inches(2.5), Inches(5), Inches(4))
    tf2 = text_box.text_frame
    tf2.word_wrap = True

    for line in translations:
        p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(36)
        p.font.bold = True

    # Save PPTX to Memory
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    clean_filename = "".join(c for c in signage_name if c.isalnum() or c in (' ', '_', '-')).strip()
    return send_file(
        output,
        as_attachment=True,
        download_name=f"{clean_filename or 'HSE_Signage'}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

@app.route("/generate_observations_report", methods=["POST"])
@login_required
def generate_observations_report():
    """
    Observations Report Automation (Word .docx)
    - Auto-sort by date
    - Auto-generate Ref No (433001/HSE-OR/XX/2026)
    - Auto-generate CC Name (DEBOTTLENECK PRODUCTION FACILITIES ABQAIQ BI NO. 10-10303)
    - Auto-map Safety Officer / Supervisor / Engineer
    - Auto-format HEISCO–ENPPI–Aramco style
    - Export Word (.docx) with complete exception handling
    """
    import traceback

    # 1. Verify and create runtime directories (Render / Linux resilience)
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(SAMPLE_DIR, exist_ok=True)
    except Exception as dir_err:
        tb = traceback.format_exc()
        print(f"[ERROR] Directory initialization failed: {dir_err}\n{tb}")
        return render_template(
            "observations.html",
            error=f"Storage directory initialization error: {str(dir_err)}",
            detailed_error=tb,
            exception_type=type(dir_err).__name__
        ), 200

    try:
        weekly_file = request.files.get("weekly_file")
        names_file = request.files.get("names_file")

        # 2. Ensure baseline sample files exist as fallbacks
        try:
            ensure_sample_files_exist()
        except Exception as sample_err:
            print(f"[WARN] ensure_sample_files_exist error: {sample_err}")

        weekly_path = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
        names_path = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")

        # 3. Handle Weekly Observations File Upload
        if weekly_file and weekly_file.filename and weekly_file.filename.strip():
            filename_lower = weekly_file.filename.lower()
            if not any(filename_lower.endswith(ext) for ext in ['.xlsx', '.xls', '.xlsm']):
                return render_template(
                    "observations.html",
                    error=f"Invalid file format: '{weekly_file.filename}'. Please upload an Excel file (.xlsx or .xls).",
                    detailed_error=f"File extension not permitted. Expected Excel workbook, received: {weekly_file.filename}",
                    exception_type="InvalidFileFormatError"
                ), 200

            uploaded_weekly = os.path.join(UPLOAD_DIR, "uploaded_weekly.xlsx")
            try:
                weekly_file.save(uploaded_weekly)
                weekly_path = uploaded_weekly
            except Exception as save_err:
                tb = traceback.format_exc()
                print(f"[ERROR] Failed to save weekly file: {save_err}\n{tb}")
                return render_template(
                    "observations.html",
                    error=f"Could not save uploaded observations file: {str(save_err)}",
                    detailed_error=tb,
                    exception_type=type(save_err).__name__
                ), 200

        # Verify weekly file exists on disk
        if not os.path.exists(weekly_path):
            return render_template(
                "observations.html",
                error="Weekly observations spreadsheet not found. Please upload your Excel observations file.",
                detailed_error=f"File not found at expected path: '{weekly_path}'.",
                exception_type="FileNotFoundError"
            ), 200

        # 4. Handle Names Mapping File Upload (Optional)
        if names_file and names_file.filename and names_file.filename.strip():
            uploaded_names = os.path.join(UPLOAD_DIR, "uploaded_names.xlsx")
            try:
                names_file.save(uploaded_names)
                names_path = uploaded_names
            except Exception as names_save_err:
                print(f"[WARN] Could not save uploaded names file: {names_save_err}")
                # Non-fatal: falls back to sample or built-in Abqaiq roster

        # 5. Read Excel Files with try/except
        try:
            raw_obs = read_weekly_observation_excel(weekly_path)
        except Exception as read_err:
            tb = traceback.format_exc()
            print(f"[ERROR] read_weekly_observation_excel failed: {read_err}\n{tb}")
            return render_template(
                "observations.html",
                error=f"Error reading Weekly Observations Excel file: {str(read_err)}",
                detailed_error=tb,
                exception_type=type(read_err).__name__
            ), 200

        if not raw_obs or len(raw_obs) == 0:
            return render_template(
                "observations.html",
                error="Could not find any observation records in the Excel file. Please ensure the sheet has rows with Location/Area, Finding/Observation, and Date.",
                detailed_error=f"read_weekly_observation_excel returned 0 records for '{weekly_path}'. Check if columns exist or if data is located on an unmapped sheet.",
                exception_type="EmptyObservationsError"
            ), 200

        # Read Names Mapping (safe, non-blocking)
        names_map = {}
        if os.path.exists(names_path):
            try:
                names_map = read_names_locations_excel(names_path)
            except Exception as names_err:
                print(f"[WARN] Error reading names mapping from '{names_path}': {names_err}")
                names_map = {}

        # 6. Process & Map Observations
        try:
            processed = process_and_map_observations(raw_obs, names_map)
        except Exception as map_err:
            tb = traceback.format_exc()
            print(f"[ERROR] process_and_map_observations failed: {map_err}\n{tb}")
            return render_template(
                "observations.html",
                error=f"Error mapping observations data: {str(map_err)}",
                detailed_error=tb,
                exception_type=type(map_err).__name__
            ), 200

        reviewed_by_name = request.form.get("reviewed_by_name", "AHMED GHALWASH").strip() or "AHMED GHALWASH"
        reviewed_by_date = request.form.get("reviewed_by_date", "").strip() or datetime.date.today().strftime("%d/%m/%Y")

        try:
            pages_data = split_into_pages(processed, reviewed_by_date=reviewed_by_date, reviewed_by_name=reviewed_by_name)
        except Exception as page_err:
            tb = traceback.format_exc()
            print(f"[ERROR] split_into_pages failed: {page_err}\n{tb}")
            return render_template(
                "observations.html",
                error=f"Error formatting pages for observations register: {str(page_err)}",
                detailed_error=tb,
                exception_type=type(page_err).__name__
            ), 200

        if not pages_data:
            return render_template(
                "observations.html",
                error="Pagination resulted in 0 pages. No valid records available to format.",
                detailed_error="split_into_pages produced an empty pages_data list.",
                exception_type="EmptyPagesError"
            ), 200

        # 7. Word Document Generation (.docx)
        output_word_path = os.path.join(OUTPUT_DIR, "HSE_Observation_Register.docx")
        try:
            generate_observation_register_doc(pages_data, output_word_path)
        except Exception as doc_err:
            tb = traceback.format_exc()
            print(f"[ERROR] generate_observation_register_doc failed: {doc_err}\n{tb}")
            return render_template(
                "observations.html",
                error=f"Error generating Word document (.docx): {str(doc_err)}",
                detailed_error=tb,
                exception_type=type(doc_err).__name__
            ), 200

        if not os.path.exists(output_word_path):
            return render_template(
                "observations.html",
                error="Word document generation completed, but output file was not found on disk.",
                detailed_error=f"Expected document at '{output_word_path}'.",
                exception_type="OutputFileMissingError"
            ), 200

        # 8. Save State to Session (Safe)
        try:
            state = load_current_state()
            state["processed"] = processed
            state["names_mapping"] = names_map
            state["reviewed_by_name"] = reviewed_by_name
            state["reviewed_by_date"] = reviewed_by_date
            save_current_state(state)
        except Exception as state_err:
            print(f"[WARN] Failed to update session state: {state_err}")

        # 9. Send Generated File as Download
        return send_file(
            output_word_path,
            as_attachment=True,
            download_name="HSE_Observation_Register.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as unhandled_err:
        tb = traceback.format_exc()
        print(f"[CRITICAL ERROR] Unhandled exception in generate_observations_report: {unhandled_err}\n{tb}")
        return render_template(
            "observations.html",
            error=f"Unexpected error generating observations report: {str(unhandled_err)}",
            detailed_error=tb,
            exception_type=type(unhandled_err).__name__
        ), 200

@app.route("/generate_weekly_report", methods=["POST"])
@login_required
def generate_weekly_report():
    """
    Weekly HSE Activity Report (Word .docx)
    - Auto-fill template
    - Auto-insert logos
    - Auto-format tables
    - Export Word
    """
    enppi_logo = "static/logos/enppi.png"
    heisco_logo = "static/logos/heisco.png"
    aramco_logo = "static/logos/aramco.png"
    pictogram_path = "static/pictograms/pictogram.png"

    activity_file = request.files.get("activity_file")
    ensure_sample_files_exist()
    activity_path = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")

    if activity_file and activity_file.filename:
        activity_path = os.path.join(UPLOAD_DIR, "uploaded_activity.xlsx")
        activity_file.save(activity_path)

    activity_data = read_weekly_activity_excel(activity_path) if os.path.exists(activity_path) else {}

    state = load_current_state()
    processed = state.get("processed", [])
    names_mapping = state.get("names_mapping", {})

    if not processed:
        weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
        names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
        raw_obs = read_weekly_observation_excel(weekly_sample)
        names_mapping = read_names_locations_excel(names_sample)
        processed = process_and_map_observations(raw_obs, names_mapping)

    week_no = request.form.get("week_no", "").strip() or activity_data.get("week_no", "36") or "36"
    period_start = request.form.get("period_start", "").strip() or activity_data.get("period_start", "05-09-2026") or "05-09-2026"
    period_end = request.form.get("period_end", "").strip() or activity_data.get("period_end", "10-09-2026") or "10-09-2026"
    prepared_by_name = request.form.get("prepared_by_name", "Obaid Khan").strip() or "Obaid Khan"
    prepared_by_title = request.form.get("prepared_by_title", "HSE Manager").strip() or "HSE Manager"

    output_weekly_path = os.path.join(OUTPUT_DIR, "Weekly_HSE_Activity_Report.docx")
    generate_weekly_activity_report(
        processed_observations=processed,
        names_mapping=names_mapping,
        output_path=output_weekly_path,
        week_no=week_no,
        period_start=period_start,
        period_end=period_end,
        activity_file_data=activity_data,
        prepared_by_name=prepared_by_name,
        prepared_by_title=prepared_by_title,
        prepared_by_date=period_start
    )

    return send_file(
        output_weekly_path,
        as_attachment=True,
        download_name="Weekly_HSE_Activity_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route("/generate_jsl_log", methods=["POST"])
@login_required
def generate_jsl_log():
    """
    JSL Log Sheet Generator (Excel .xlsx)
    - Auto-format log sheet
    - Auto-insert logos
    - Export Excel
    """
    enppi_logo = "static/logos/enppi.png"
    heisco_logo = "static/logos/heisco.png"
    aramco_logo = "static/logos/aramco.png"
    pictogram_path = "static/pictograms/pictogram.png"

    jsl_file = request.files.get("jsl_file")
    state = load_current_state()
    processed = []

    if jsl_file and jsl_file.filename:
        uploaded_jsl_path = os.path.join(UPLOAD_DIR, "uploaded_jsl.xlsx")
        jsl_file.save(uploaded_jsl_path)
        raw_obs = read_weekly_observation_excel(uploaded_jsl_path)
        ensure_sample_files_exist()
        names_map = read_names_locations_excel(os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx"))
        processed = process_and_map_observations(raw_obs, names_map)
    else:
        processed = state.get("processed", [])
        if not processed:
            ensure_sample_files_exist()
            weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
            names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
            raw_obs = read_weekly_observation_excel(weekly_sample)
            names_map = read_names_locations_excel(names_sample)
            processed = process_and_map_observations(raw_obs, names_map)

    project_name = request.form.get("project_name", "ABQ-DBN Project").strip() or "ABQ-DBN Project"
    project_number = request.form.get("project_number", "BI-10-10303").strip() or "BI-10-10303"
    last_update_date = processed[-1].get("observation_date", "") if processed else ""

    output_excel_path = os.path.join(OUTPUT_DIR, "HSE_Action_Tracking_Register_JSL.xlsx")
    generate_jsl_log_excel(
        processed_observations=processed,
        output_path=output_excel_path,
        project_name=project_name,
        project_number=project_number,
        last_update_date=last_update_date
    )

    return send_file(
        output_excel_path,
        as_attachment=True,
        download_name="HSE_Action_Tracking_Register_JSL.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =============================================================
# COMBINED BATCH GENERATOR & DOWNLOADS
# =============================================================

@app.route("/generate", methods=["POST"])
@login_required
def generate_all_reports():
    """Generates Observation Register, JSL Log, and Weekly Activity Report all together."""
    state = load_current_state()
    processed = state.get("processed", [])
    names_mapping = state.get("names_mapping", {})
    activity_data = state.get("activity_data", {})

    if not processed:
        ensure_sample_files_exist()
        weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
        names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
        activity_sample = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")
        raw_obs = read_weekly_observation_excel(weekly_sample)
        names_mapping = read_names_locations_excel(names_sample)
        processed = process_and_map_observations(raw_obs, names_mapping)
        activity_data = read_weekly_activity_excel(activity_sample)

    reviewed_by_date = request.form.get("reviewed_by_date", "").strip() or datetime.date.today().strftime("%d/%m/%Y")
    reviewed_by_name = request.form.get("reviewed_by_name", "AHMED GHALWASH").strip() or "AHMED GHALWASH"
    week_no = request.form.get("week_no", "").strip() or (activity_data.get("week_no") if activity_data else "36") or "36"
    period_start = request.form.get("period_start", "").strip() or (activity_data.get("period_start") if activity_data else "05-09-2026") or "05-09-2026"
    period_end = request.form.get("period_end", "").strip() or (activity_data.get("period_end") if activity_data else "10-09-2026") or "10-09-2026"

    pages_data = split_into_pages(processed, reviewed_by_date=reviewed_by_date, reviewed_by_name=reviewed_by_name)

    output_word_path = os.path.join(OUTPUT_DIR, "HSE_Observation_Register.docx")
    output_excel_path = os.path.join(OUTPUT_DIR, "HSE_Action_Tracking_Register_JSL.xlsx")
    output_weekly_path = os.path.join(OUTPUT_DIR, "Weekly_HSE_Activity_Report.docx")
    output_zip_path = os.path.join(OUTPUT_DIR, "All_HSE_Reports.zip")

    generate_observation_register_doc(pages_data, output_word_path)

    generate_jsl_log_excel(
        processed_observations=processed,
        output_path=output_excel_path,
        project_name="ABQ-DBN Project",
        project_number="BI-10-10303",
        last_update_date=processed[-1].get("observation_date", "") if processed else ""
    )

    generate_weekly_activity_report(
        processed_observations=processed,
        names_mapping=names_mapping,
        output_path=output_weekly_path,
        week_no=week_no,
        period_start=period_start,
        period_end=period_end,
        activity_file_data=activity_data,
        prepared_by_name="Obaid Khan",
        prepared_by_title="HSE Manager",
        prepared_by_date=period_start
    )

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(output_word_path, arcname="1_HSE_Observation_Register.docx")
        zipf.write(output_excel_path, arcname="2_HSE_Action_Tracking_Register_JSL.xlsx")
        zipf.write(output_weekly_path, arcname="3_Weekly_HSE_Activity_Report.docx")

    state["reviewed_by_date"] = reviewed_by_date
    state["reviewed_by_name"] = reviewed_by_name
    state["week_no"] = week_no
    state["period_start"] = period_start
    state["period_end"] = period_end
    state["total_pages"] = len(pages_data)
    state["total_records"] = len(processed)
    save_current_state(state)

    return redirect(url_for("download_page"))

@app.route("/download")
@login_required
def download_page():
    state = load_current_state()
    return render_template(
        "download.html",
        total_pages=state.get("total_pages", 12),
        total_records=state.get("total_records", 35),
        reviewed_by_date=state.get("reviewed_by_date", "05/09/2026"),
        reviewed_by_name=state.get("reviewed_by_name", "AHMED GHALWASH"),
        week_no=state.get("week_no", "36")
    )

@app.route("/download/word")
@login_required
def download_word():
    output_word_path = os.path.join(OUTPUT_DIR, "HSE_Observation_Register.docx")
    if not os.path.exists(output_word_path):
        use_sample()
        generate_all_reports()
    return send_file(
        output_word_path,
        as_attachment=True,
        download_name="HSE_Observation_Register.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route("/download/excel")
@login_required
def download_excel():
    output_excel_path = os.path.join(OUTPUT_DIR, "HSE_Action_Tracking_Register_JSL.xlsx")
    if not os.path.exists(output_excel_path):
        use_sample()
        generate_all_reports()
    return send_file(
        output_excel_path,
        as_attachment=True,
        download_name="HSE_Action_Tracking_Register_JSL.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/download/weekly")
@login_required
def download_weekly():
    output_weekly_path = os.path.join(OUTPUT_DIR, "Weekly_HSE_Activity_Report.docx")
    if not os.path.exists(output_weekly_path):
        use_sample()
        generate_all_reports()
    return send_file(
        output_weekly_path,
        as_attachment=True,
        download_name="Weekly_HSE_Activity_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route("/download/all")
@login_required
def download_all():
    output_zip_path = os.path.join(OUTPUT_DIR, "All_HSE_Reports.zip")
    if not os.path.exists(output_zip_path):
        use_sample()
        generate_all_reports()
    return send_file(
        output_zip_path,
        as_attachment=True,
        download_name="All_HSE_Reports.zip",
        mimetype="application/zip"
    )

@app.route("/sample/weekly")
def download_sample_weekly():
    ensure_sample_files_exist()
    path = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
    return send_file(
        path,
        as_attachment=True,
        download_name="Weekly_Observation_Sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/sample/names")
def download_sample_names():
    ensure_sample_files_exist()
    path = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
    return send_file(
        path,
        as_attachment=True,
        download_name="Names_Locations_Sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/sample/activity")
def download_sample_activity():
    ensure_sample_files_exist()
    path = os.path.join(SAMPLE_DIR, "Weekly_HSE_Activity_Sample.xlsx")
    return send_file(
        path,
        as_attachment=True,
        download_name="Weekly_HSE_Activity_Sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="HEISCO-ENPPI-Aramco HSE Automation Web App")
    parser.add_argument("-p", "--port", type=int, default=None, help="Port to listen on")
    parser.add_argument("-H", "--host", type=str, default=None, help="Host to bind to")
    args, unknown = parser.parse_known_args()

    port = args.port or int(os.environ.get("PORT", 3000))
    host = args.host or os.environ.get("HOST", "0.0.0.0")

    print(f"Starting Central HSE Operations Application on http://{host}:{port}")
    app.run(host=host, port=port, debug=False)
