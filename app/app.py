import os
import uuid
import shutil
import logging
import json
from datetime import datetime
from flask import Flask, request, render_template, send_file, jsonify, redirect, url_for
from werkzeug.utils import secure_filename
import threading
import time
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.geminiOCR.pdf_to_json import main as pdf_to_json_main
from src.geminiOCR.json_to_excel import main as json_to_excel_main

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'herrohowyoudoin')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSING_FOLDER'] = 'processing'
app.config['OUTPUT_FOLDER'] = 'output'

for folder in [app.config['UPLOAD_FOLDER'], app.config['PROCESSING_FOLDER'], app.config['OUTPUT_FOLDER']]:
    os.makedirs(folder, exist_ok=True)

job_status = {}

class ProcessingJob:
    def __init__(self, job_id: str, filename: str):
        self.job_id = job_id
        self.filename = filename
        self.status = "uploaded"
        self.progress = 0
        self.message = "File uploaded successfully"
        self.start_time = datetime.now()
        self.json_path = None
        self.excel_path = None
        self.error = None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

def cleanup_old_files():
    current_time = time.time()
    for folder in [app.config['UPLOAD_FOLDER'], app.config['PROCESSING_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path):
                if current_time - os.path.getmtime(file_path) > 3600:
                    try:
                        os.remove(file_path)
                        logger.info(f"Cleaned up old file: {file_path}")
                    except Exception as e:
                        logger.error(f"Error cleaning up {file_path}: {e}")

def process_pdf_async(job_id: str, pdf_path: str):
    job = job_status[job_id]
    
    try:
        job.status = "processing"
        job.progress = 10
        job.message = "Converting PDF to images..."
        
        job_processing_dir = os.path.join(app.config['PROCESSING_FOLDER'], job_id)
        os.makedirs(job_processing_dir, exist_ok=True)
        
        json_output = os.path.join(os.path.dirname(app.root_path), 'output', "results.json")
        excel_output = os.path.join(os.path.dirname(app.root_path), 'output', "results.xlsx")
        
        job.progress = 30
        job.message = "Extracting structured data from PDF..."
        pdf_to_json_main(pdf_path, job_processing_dir, json_output)
        job.json_path = json_output

        job.progress = 70
        job.message = "Converting data to Excel format..."
        json_to_excel_main(json_output, excel_output)
        job.excel_path = excel_output
    
        job.status = "completed"
        job.progress = 100
        job.message = "Processing completed successfully!"
        
        logger.info(f"Job {job_id} completed successfully")
        
    except Exception as e:
        job.status = "error"
        job.error = str(e)
        job.message = f"Error during processing: {str(e)}"
        logger.error(f"Job {job_id} failed: {e}")
    
    finally:
        try:
            if os.path.exists(job_processing_dir):
                shutil.rmtree(job_processing_dir)
        except Exception as e:
            logger.error(f"Error cleaning up processing directory for job {job_id}: {e}")

def process_pdf_direct(pdf_path: str):
    temp_id = str(uuid.uuid4())
    temp_processing_dir = os.path.join(app.config['PROCESSING_FOLDER'], temp_id)
    
    try:
        os.makedirs(temp_processing_dir, exist_ok=True)
        
        json_output = os.path.join(temp_processing_dir, f"results_{temp_id}.json")
        
        logger.info(f"Starting direct PDF processing for {pdf_path}")
        
        pdf_to_json_main(pdf_path, temp_processing_dir, json_output)
        
        if os.path.exists(json_output):
            with open(json_output, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            logger.info("Direct PDF processing completed successfully")
            return json_data
        else:
            raise Exception("JSON output file not created")
            
    except Exception as e:
        logger.error(f"Direct PDF processing failed: {e}")
        raise e
    
    finally:
        try:
            if os.path.exists(temp_processing_dir):
                shutil.rmtree(temp_processing_dir)
        except Exception as e:
            logger.error(f"Error cleaning up temp directory {temp_processing_dir}: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert-pdf', methods=['POST'])
def convert_pdf_direct():
    try:
        if 'file' not in request.files:
            logger.error("No file in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Only PDF files are allowed'}), 400
        
        temp_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        temp_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{temp_id}_{filename}")
        
        logger.info(f"Saving uploaded file to {temp_file_path}")
        file.save(temp_file_path)
        
        try:
            # Process PDF and get JSON data
            json_data = process_pdf_direct(temp_file_path)
            
            # Return the JSON data directly
            response = jsonify({
                'status': 'success',
                'message': 'PDF converted successfully',
                'data': json_data
            })
            
            logger.info(f"Successfully processed PDF: {filename}")
            return response
            
        finally:
            # Clean up uploaded file
            try:
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)
            except Exception as e:
                logger.error(f"Error cleaning up uploaded file {temp_file_path}: {e}")    
    except Exception as e:
        logger.error(f"Error in convert_pdf_direct: {e}")
        return jsonify({
            'status': 'error',
            'error': str(e),
            'message': 'Failed to process PDF'
        }), 500

@app.route('/convert', methods=['POST'])
def upload_file():
    """Handle file upload for async processing (existing functionality)"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Only PDF files are allowed'}), 400
        
        job_id = str(uuid.uuid4())
        
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_{filename}")
        file.save(file_path)
        
        job = ProcessingJob(job_id, filename)
        job_status[job_id] = job
        
        processing_thread = threading.Thread(
            target=process_pdf_async,
            args=(job_id, file_path)
        )
        processing_thread.start()
        
        return jsonify({
            'job_id': job_id,
            'message': 'File uploaded successfully. Processing started.',
            'status_url': url_for('get_status', job_id=job_id)
        })
        
    except Exception as e:
        logger.error(f"Upload error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/status/<job_id>')
def get_status(job_id):
    if job_id not in job_status:
        return jsonify({'error': 'Job not found'}), 404
    
    job = job_status[job_id]
    
    response_data = {
        'job_id': job_id,
        'status': job.status,
        'progress': job.progress,
        'message': job.message,
        'filename': job.filename
    }
    
    if job.status == 'completed':
        response_data['download_url'] = url_for('download_file', job_id=job_id)
    elif job.status == 'error':
        response_data['error'] = job.error
    
    return jsonify(response_data)

@app.route('/download/<job_id>')
def download_file(job_id):
    if job_id not in job_status:
        return jsonify({'error': 'Job not found'}), 404
    
    job = job_status[job_id]
    
    if job.status != 'completed' or not job.excel_path:
        return jsonify({'error': 'File not ready for download'}), 400
    
    if not os.path.exists(job.excel_path):
        return jsonify({'error': 'Output file not found'}), 404
    
    try:
        return send_file(
            job.excel_path,
            as_attachment=True,
            download_name=f"{os.path.splitext(job.filename)[0]}_converted.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.error(f"Download error for job {job_id}: {e}")
        return jsonify({'error': 'Error downloading file'}), 500

@app.route('/progress/<job_id>')
def progress_page(job_id):
    if job_id not in job_status:
        return redirect(url_for('index'))
    
    return render_template('progress.html', job_id=job_id)

@app.route('/health')
def health_check():
    """Health check endpoint for VBA to test connectivity"""
    return jsonify({
        'status': 'healthy',
        'message': 'PDF Converter API is running',
        'timestamp': datetime.now().isoformat()
    })

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 413

@app.errorhandler(500)
def internal_error(e):
    return jsonify({'error': 'Internal server error. Please try again.'}), 500

def periodic_cleanup():
    while True:
        try:
            cleanup_old_files()
            
            current_time = datetime.now()
            jobs_to_remove = []
            for job_id, job in job_status.items():
                if (current_time - job.start_time).total_seconds() > 7200:  # 2 hours
                    jobs_to_remove.append(job_id)
            
            for job_id in jobs_to_remove:
                del job_status[job_id]
                logger.info(f"Cleaned up old job status: {job_id}")
                
        except Exception as e:
            logger.error(f"Error in periodic cleanup: {e}")
        
        time.sleep(1800)

if __name__ == '__main__':
    cleanup_thread = threading.Thread(target=periodic_cleanup, daemon=True)
    cleanup_thread.start()
    
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    app.run(host='0.0.0.0', port=port, debug=debug)