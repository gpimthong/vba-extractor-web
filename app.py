import os
import zipfile
import pandas as pd
from flask import Flask, render_template, request, send_file
from oletools.olevba import VBA_Parser
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file", 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    zip_path = os.path.join(OUTPUT_FOLDER, f"{filename}_export.zip")

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # 1. Extract VBA
        vba_parser = VBA_Parser(filepath)
        if vba_parser.detect_vba():
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                zipf.writestr(f"vba/{vba_filename}.vb", vba_code)
        vba_parser.close()

        # 2. Extract Sheets to CSV
        try:
            excel_data = pd.read_excel(filepath, sheet_name=None)
            for sheet_name, df in excel_data.items():
                csv_data = df.to_csv(index=False)
                zipf.writestr(f"sheets/{sheet_name}.csv", csv_data)
        except Exception as e:
            print(f"Sheet error: {e}")

    return send_file(zip_path, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
