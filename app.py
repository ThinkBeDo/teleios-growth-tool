from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import os
import io
from werkzeug.utils import secure_filename
from utils.excel_processor import process_county_files
import tempfile
import zipfile

app = Flask(__name__)
app.secret_key = 'teleios-growth-tool-secret-key'

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Check if main workbook is provided
        if 'main_workbook' not in request.files:
            flash('No main workbook file selected')
            return redirect(request.url)
        
        main_file = request.files['main_workbook']
        if main_file.filename == '':
            flash('No main workbook file selected')
            return redirect(request.url)
        
        # Check if county files are provided
        county_files = request.files.getlist('county_files')
        if not county_files or all(f.filename == '' for f in county_files):
            flash('No county files selected')
            return redirect(request.url)
        
        # Validate file types
        if not allowed_file(main_file.filename):
            flash('Main workbook must be an Excel file (.xlsx or .xls)')
            return redirect(request.url)
        
        valid_county_files = []
        for file in county_files:
            if file.filename != '' and allowed_file(file.filename):
                valid_county_files.append(file)
        
        if not valid_county_files:
            flash('No valid county Excel files found')
            return redirect(request.url)
        
        # Process the files
        result_file = process_county_files(main_file, valid_county_files)
        
        if result_file:
            # Return the processed file
            return send_file(
                result_file,
                as_attachment=True,
                download_name='updated_growth_tool.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            flash('Error processing files. Please check your data and try again.')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'Error processing files: {str(e)}')
        return redirect(url_for('index'))

@app.route('/health')
def health_check():
    return {'status': 'healthy', 'service': 'Teleios Growth Tool'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
