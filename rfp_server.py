"""
RFP Dashboard Server
====================
Flask server that provides:
  1. Dashboard UI at http://localhost:5000
  2. File upload endpoint (Excel, PDF, Word, JSON, CSV, TXT)
  3. One-click RFP generation (no terminal commands needed)

Usage:
    python rfp_server.py
    Then open http://localhost:5000 in your browser.
"""

import json
import os
import sys
import tempfile
import traceback
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, send_file
from werkzeug.utils import secure_filename

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_rfp import generate_rfp_content, fill_template, read_rfi_file

app = Flask(__name__, static_folder='.', static_url_path='')

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf', 'docx', 'doc', 'json', 'csv', 'txt'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return send_from_directory('.', 'rfp_dashboard.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Upload a file and extract its content as RFI insights."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': f'Unsupported file type. Allowed: {", ".join(ALLOWED_EXTENSIONS)}'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_DIR, filename)
    file.save(filepath)

    try:
        rfi_data = read_rfi_file(filepath)
        return jsonify({
            'success': True,
            'filename': filename,
            'data': rfi_data,
            'message': f'Successfully read {filename}'
        })
    except Exception as e:
        return jsonify({'error': f'Error reading file: {str(e)}'}), 500


@app.route('/generate', methods=['POST'])
def generate():
    """Generate RFP from provided data (form fields or uploaded file data)."""
    try:
        rfi_insights = request.get_json()
        if not rfi_insights:
            return jsonify({'error': 'No data provided'}), 400

        # Generate RFP content via AI
        rfp_content = generate_rfp_content(rfi_insights)

        # Determine output filename
        company = rfi_insights.get('company', 'Client').replace(' ', '_')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        output_filename = f"{company}_Temp_Labor_RFP_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        # Save intermediate JSON
        json_path = os.path.join(OUTPUT_DIR, f"{company}_rfp_content_{timestamp}.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(rfp_content, f, indent=2, ensure_ascii=False)

        # Fill template
        fill_template(rfp_content, rfi_insights, output_path)

        return jsonify({
            'success': True,
            'output_file': output_filename,
            'output_path': output_path,
            'stats': {
                'requirements': len(rfp_content.get('requirements', [])),
                'questions': len(rfp_content.get('questions', [])),
                'slas': len(rfp_content.get('slas', [])),
                'roles': len(rfp_content.get('roles', [])),
                'markups': len(rfp_content.get('markups', [])),
                'conversions': len(rfp_content.get('conversions', []))
            }
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download(filename):
    """Download a generated RFP file."""
    filepath = os.path.join(OUTPUT_DIR, secure_filename(filename))
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404


if __name__ == '__main__':
    print("=" * 60)
    print("  RFP DASHBOARD SERVER")
    print("=" * 60)
    print()
    print("  Open in browser: http://localhost:5000")
    print()
    print("  Features:")
    print("    - Upload Excel/PDF/Word files directly")
    print("    - One-click RFP generation (no terminal needed)")
    print("    - Download generated Excel files")
    print()
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
