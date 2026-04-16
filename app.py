import os
import io
from flask import Flask, request, send_file, jsonify, render_template
from generate_photosheet import generate_photosheet

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    if 'log' not in request.files:
        return jsonify({'error': 'No photo log file uploaded.'}), 400

    log_file = request.files['log']
    photo_files = request.files.getlist('photos')

    if not log_file.filename:
        return jsonify({'error': 'Photo log file is empty.'}), 400

    log_bytes = log_file.read()
    log_filename = log_file.filename.lower()

    # Build photo map: lowercase basename -> bytes
    photo_map = {}
    for pf in photo_files:
        if pf.filename:
            photo_map[os.path.basename(pf.filename).lower()] = pf.read()

    try:
        docx_bytes = generate_photosheet(log_bytes, log_filename, photo_map)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='photo_sheet.docx'
    )


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
