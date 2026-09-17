import os
import sys
import json
import argparse
import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session

from services.excel_reader import read_weekly_observation_excel, read_names_locations_excel
from services.mapping import process_and_map_observations, split_into_pages
from services.or_generator import generate_observation_register_doc
from services.jsl_generator import generate_jsl_log_excel

# Base paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SAMPLE_DIR, exist_ok=True)

app = Flask(__name__, template_folder="templates")
app.secret_key = "hse-observation-register-secret-key-abqaiq"

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

@app.route("/")
def index():
    return redirect(url_for("upload_page"))

@app.route("/upload", methods=["GET", "POST"])
def upload_page():
    if request.method == "POST":
        weekly_file = request.files.get("weekly_file")
        names_file = request.files.get("names_file")

        if not weekly_file or not weekly_file.filename:
            return render_template("upload.html", error="Please select the Weekly Observation Excel file.")
        if not names_file or not names_file.filename:
            return render_template("upload.html", error="Please select the Names + Locations Excel file.")

        try:
            weekly_path = os.path.join(UPLOAD_DIR, "uploaded_weekly.xlsx")
            names_path = os.path.join(UPLOAD_DIR, "uploaded_names.xlsx")
            
            weekly_file.save(weekly_path)
            names_file.save(names_path)

            raw_obs = read_weekly_observation_excel(weekly_path)
            names_map = read_names_locations_excel(names_path)

            if not raw_obs:
                return render_template("upload.html", error="Could not extract any observation records from the weekly file.")

            processed = process_and_map_observations(raw_obs, names_map)
            
            save_current_state({
                "processed": processed,
                "weekly_filename": weekly_file.filename,
                "names_filename": names_file.filename
            })

            return redirect(url_for("preview_page"))

        except Exception as e:
            return render_template("upload.html", error=f"Processing failed: {str(e)}")

    return render_template("upload.html")

@app.route("/use-sample", methods=["POST"])
def use_sample():
    """Immediately load pre-configured sample files and go to preview."""
    weekly_sample = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
    names_sample = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
    
    if not os.path.exists(weekly_sample) or not os.path.exists(names_sample):
        from sample_data.generate_samples import generate_sample_files
        generate_sample_files()

    try:
        raw_obs = read_weekly_observation_excel(weekly_sample)
        names_map = read_names_locations_excel(names_sample)
        processed = process_and_map_observations(raw_obs, names_map)

        save_current_state({
            "processed": processed,
            "weekly_filename": "Weekly_Observation_Sample.xlsx",
            "names_filename": "Names_Locations_Sample.xlsx"
        })

        return redirect(url_for("preview_page"))
    except Exception as e:
        return render_template("upload.html", error=f"Failed to load sample data: {str(e)}")

@app.route("/preview")
def preview_page():
    state = load_current_state()
    processed = state.get("processed", [])
    
    if not processed:
        return redirect(url_for("upload_page"))

    total_records = len(processed)
    total_pages = (total_records + 2) // 3
    
    # Extract unique locations
    locations = sorted(list({item.get("area", "") for item in processed if item.get("area")}))
    
    # Calculate date range
    dates = [item.get("observation_date", "") for item in processed if item.get("observation_date")]
    if dates:
        date_range = f"{dates[0]} to {dates[-1]}" if len(dates) > 1 else dates[0]
        # Propose reviewed by date 5-7 days after the last observation date, or default to 05/09/2026
        default_reviewed_date = "05/09/2026"
    else:
        date_range = "N/A"
        default_reviewed_date = datetime.date.today().strftime("%d/%m/%Y")

    # Showing first 20 rows
    preview_records = processed[:20]

    return render_template(
        "preview.html",
        preview_records=preview_records,
        total_records=total_records,
        total_pages=total_pages,
        unique_locations=locations,
        date_range=date_range,
        default_reviewed_date=default_reviewed_date
    )

@app.route("/generate", methods=["POST"])
def generate_reports():
    state = load_current_state()
    processed = state.get("processed", [])
    
    if not processed:
        return redirect(url_for("upload_page"))

    reviewed_by_date = request.form.get("reviewed_by_date", "").strip()
    reviewed_by_name = request.form.get("reviewed_by_name", "AHMED GHALWASH").strip() or "AHMED GHALWASH"

    if not reviewed_by_date:
        reviewed_by_date = datetime.date.today().strftime("%d/%m/%Y")

    # 1. Split into 3 observations per page
    pages_data = split_into_pages(processed, reviewed_by_date=reviewed_by_date, reviewed_by_name=reviewed_by_name)

    # 2. Output file paths
    output_word_path = os.path.join(OUTPUT_DIR, "HSE_Observation_Register.docx")
    output_excel_path = os.path.join(OUTPUT_DIR, "HSE_Action_Tracking_Register_JSL.xlsx")

    # 3. Generate Word Document
    generate_observation_register_doc(pages_data, output_word_path)

    # 4. Generate JSL Log Excel
    generate_jsl_log_excel(
        processed_observations=processed,
        output_path=output_excel_path,
        project_name="ABQ-DBN Project",
        project_number="BI-10-10303",
        last_update_date=processed[-1].get("observation_date", "") if processed else ""
    )

    # Save generation details
    state["reviewed_by_date"] = reviewed_by_date
    state["reviewed_by_name"] = reviewed_by_name
    state["total_pages"] = len(pages_data)
    state["total_records"] = len(processed)
    save_current_state(state)

    return redirect(url_for("download_page"))

@app.route("/download")
def download_page():
    state = load_current_state()
    if not state.get("processed"):
        return redirect(url_for("upload_page"))

    return render_template(
        "download.html",
        total_pages=state.get("total_pages", 0),
        total_records=state.get("total_records", 0),
        reviewed_by_date=state.get("reviewed_by_date", ""),
        reviewed_by_name=state.get("reviewed_by_name", "AHMED GHALWASH")
    )

@app.route("/download/word")
def download_word():
    output_word_path = os.path.join(OUTPUT_DIR, "HSE_Observation_Register.docx")
    if not os.path.exists(output_word_path):
        return redirect(url_for("preview_page"))
    return send_file(
        output_word_path,
        as_attachment=True,
        download_name="HSE_Observation_Register.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route("/download/excel")
def download_excel():
    output_excel_path = os.path.join(OUTPUT_DIR, "HSE_Action_Tracking_Register_JSL.xlsx")
    if not os.path.exists(output_excel_path):
        return redirect(url_for("preview_page"))
    return send_file(
        output_excel_path,
        as_attachment=True,
        download_name="HSE_Action_Tracking_Register_JSL.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/sample/weekly")
def download_sample_weekly():
    path = os.path.join(SAMPLE_DIR, "Weekly_Observation_Sample.xlsx")
    if not os.path.exists(path):
        from sample_data.generate_samples import generate_sample_files
        generate_sample_files()
    return send_file(
        path,
        as_attachment=True,
        download_name="Weekly_Observation_Sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/sample/names")
def download_sample_names():
    path = os.path.join(SAMPLE_DIR, "Names_Locations_Sample.xlsx")
    if not os.path.exists(path):
        from sample_data.generate_samples import generate_sample_files
        generate_sample_files()
    return send_file(
        path,
        as_attachment=True,
        download_name="Names_Locations_Sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="HSE Observation Register Generator")
    parser.add_argument("--port", type=int, default=None, help="Port to listen on")
    parser.add_argument("--host", type=str, default=None, help="Host to bind to")
    args, unknown = parser.parse_known_args()

    port = args.port or int(os.environ.get("PORT", 3000))
    host = args.host or os.environ.get("HOST", "0.0.0.0")

    print(f"Starting HSE Flask Application on http://{host}:{port}")
    app.run(host=host, port=port, debug=False)
