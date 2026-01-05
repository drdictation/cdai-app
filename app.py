import os
import zipfile
import pandas as pd
import fitz  # PyMuPDF
import logging
import shutil
from flask import Flask, request, render_template, after_this_request, send_file, send_from_directory, abort
from werkzeug.utils import secure_filename

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
COMPLETED_FOLDER = 'completed'
STATIC_FOLDER = 'static'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['COMPLETED_FOLDER'] = COMPLETED_FOLDER

# Setup basic logging
logging.basicConfig(level=logging.INFO)

# --- Helper Functions ---

def split_pdf(input_pdf_path, output_dir):
    app.logger.info(f"Splitting PDF: {input_pdf_path}")
    pdf_document = fitz.open(input_pdf_path)
    split_paths = []
    for page_num in range(len(pdf_document)):
        split_pdf_path = os.path.join(output_dir, f"page_{page_num + 1}.pdf")
        split_paths.append(split_pdf_path)
        new_pdf = fitz.open()
        new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        new_pdf.save(split_pdf_path)
        new_pdf.close()
    pdf_document.close()
    return split_paths

def fill_page(pdf_path, data, coordinates):
    pdf_document = fitz.open(pdf_path)
    page = pdf_document.load_page(0)
    for field, coord in coordinates.items():
        if field in data and pd.notna(data[field]):
            value = str(data[field])
            # coord is expected to be (x, y)
            if isinstance(coord[0], tuple): # handle list of coords if needed
                 for c in coord:
                     x, y = c
                     # Removed draw_rect to stop obscuring PDF lines
                     # Rotation 0 for horizontal normal reading
                     page.insert_text((x, y+5), value, fontsize=12, rotate=0, color=(0, 0, 0))
            else:
                x, y = coord
                # Removed draw_rect to stop obscuring PDF lines
                page.insert_text((x, y+5), value, fontsize=12, rotate=0, color=(0, 0, 0))

    temp_pdf_path = pdf_path.replace(".pdf", "_filled.pdf")
    pdf_document.save(temp_pdf_path)
    pdf_document.close()
    return temp_pdf_path

def merge_pdfs(output_pdf_path, pdf_paths):
    merged_pdf = fitz.open()
    for pdf_path in pdf_paths:
        if os.path.exists(pdf_path):
            pdf_document = fitz.open(pdf_path)
            merged_pdf.insert_pdf(pdf_document)
            pdf_document.close()
    merged_pdf.save(output_pdf_path)
    merged_pdf.close()

# --- Main Application Logic ---

@app.route('/', methods=['GET'])
def main_page():
    return render_template('index.html')

@app.route('/download/<path:filename>', methods=['GET'])
def download_example(filename):
    # Serve example files from the static/examples directory bundled with the app
    example_dir = os.path.join(app.static_folder, 'examples')
    file_path = os.path.join(example_dir, filename)
    if not os.path.isfile(file_path):
        app.logger.warning(f"Example download missing: {file_path}")
        abort(404)
    return send_from_directory(example_dir, filename, as_attachment=True)

@app.route('/process', methods=['POST'])
def process_files():
    app.logger.info("POST request received. Starting PDF generation process.")

    if 'data_file' not in request.files:
        return "Missing data file", 400
    
    data_file = request.files['data_file']
    # template_files = request.files.getlist('template_files') # Removed requirement for upload
    
    # Use Hardcoded Local Template
    LOCAL_TEMPLATE_PATH = os.path.abspath("BLANK_CDAI_APP_DEC-25.pdf")
    if not os.path.exists(LOCAL_TEMPLATE_PATH):
        return "Server Error: Default PDF template not found.", 500

    if data_file.filename == '':
        return "No selected file", 400

    run_id = secure_filename(pd.Timestamp.now().strftime('%Y%m%d_%H%M%S_%f'))
    session_upload_folder = os.path.join(app.config['UPLOAD_FOLDER'], run_id)
    session_completed_folder = os.path.join(app.config['COMPLETED_FOLDER'], run_id)
    os.makedirs(session_upload_folder, exist_ok=True)
    os.makedirs(session_completed_folder, exist_ok=True)

    data_filename = secure_filename(data_file.filename)
    data_path = os.path.join(session_upload_folder, data_filename)
    data_file.save(data_path)

    # Copy template to session folder
    template_filename = "template.pdf"
    session_template_path = os.path.join(session_upload_folder, template_filename)
    shutil.copy(LOCAL_TEMPLATE_PATH, session_template_path)

    ext = os.path.splitext(data_filename)[1].lower()
    df = pd.read_csv(data_path) if ext == '.csv' else pd.read_excel(data_path)
    
    df.columns = [c.strip().lower().replace(" ", "") for c in df.columns]
    df.rename(columns={
        'birthdate': 'dob', 'medicarenumber': 'mc', 'infusionlocation': 'location',
        'height': 'ht', 'weight': 'wt',
    }, inplace=True)

    processing_log = []
    generated_files = []

    processing_log.append("--- CDAI PDF Processing Report ---")
    processing_log.append(f"Data file: {data_filename}")
    processing_log.append(f"Using Default Template: {LOCAL_TEMPLATE_PATH}")
    processing_log.append("\n---\n")

    for index, row in df.iterrows():
        data = row.to_dict()
        patient_name = f"{data.get('firstname', '')} {data.get('lastname', 'N/A')}"
        row_identifier = f"Row {index + 2} (Patient: {patient_name})"

        # Skipped Gastro check since we use one template
        
        patient_lastname = data.get('lastname', 'unknown_lastname')
        patient_firstname = data.get('firstname', 'unknown_firstname')
        filename = f"{patient_lastname}_{patient_firstname}_CDAI.pdf"
        out_pdf = os.path.join(session_completed_folder, filename)

        try:
            pages = split_pdf(session_template_path, session_upload_folder)
            filled_pages = []
            
            # UPDATED COORDINATES (X, Y)
            coords_p1 = {
                'mc': (75, 250), 
                'lastname': (75, 330), 
                'firstname': (75, 360),
                'dob': (75, 400), 
                'wt': (75, 430), 
                'ht': (75, 475),
                # 'location': (585.36, 347.04), # User has not updated this
            }
            # coords_p2 = {'last': (731.52, 86.16)} # Removed
            coords_p3 = {'wt': (325, 525)}

            if len(pages) > 0: filled_pages.append(fill_page(pages[0], data, coords_p1))
            if len(pages) > 1: filled_pages.append(pages[1]) # No changes to P2
            if len(pages) > 2: filled_pages.append(fill_page(pages[2], data, coords_p3))
            if len(pages) > 3: # Handle extra pages if any
                 filled_pages.extend(pages[3:])

            merge_pdfs(out_pdf, filled_pages)
            
            for p in pages + filled_pages:
                if os.path.exists(p): os.remove(p)
            
            generated_files.append(out_pdf)
            log_entry = f"SUCCESS: {row_identifier} -> Generated {filename}"
            processing_log.append(log_entry)

        except Exception as e:
            log_entry = f"ERROR: {row_identifier} - An unexpected error occurred: {e}"
            app.logger.error(log_entry, exc_info=True)
            processing_log.append(log_entry)
            continue
    
    processing_log.append("\n---\nReport finished.")
    report_content = "\n".join(processing_log)
    report_path = os.path.join(session_completed_folder, "processing_report.txt")
    with open(report_path, "w") as f:
        f.write(report_content)

    if not generated_files and len(processing_log) <= 4:
         return "Processing finished, but the data file was empty or no actions were taken.", 400

    zip_filename = f"CDAI_processed_{run_id}.zip"
    zip_path = os.path.join(session_completed_folder, zip_filename)
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for f in generated_files:
            zf.write(f, arcname=os.path.basename(f))
        zf.write(report_path, arcname="processing_report.txt")
    
    @after_this_request
    def cleanup(response):
        app.logger.info(f"Cleaning up temporary files for run_id: {run_id}")
        try:
            shutil.rmtree(session_upload_folder)
            shutil.rmtree(session_completed_folder)
        except Exception as e:
            app.logger.error(f"Error during cleanup: {e}")
        return response

    app.logger.info(f"Sending zip file: {zip_filename} from path {zip_path}")
    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['COMPLETED_FOLDER'], exist_ok=True)
    app.run(debug=True)
