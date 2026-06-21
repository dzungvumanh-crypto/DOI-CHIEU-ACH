"""
web_app.py — Web UI cho DOI-CHIEU-ACH
Chay:   python web_app.py
Truy cap tu LAN:  http://<IP_MAY_CHU>:8080
"""
import os
import uuid
import threading

from flask import Flask, request, jsonify, render_template, send_file
from flask_socketio import SocketIO

from main import main_from_dir

app = Flask(__name__)
app.config['SECRET_KEY'] = 'doi_chieu_ach_secret'
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB

socketio = SocketIO(
    app,
    async_mode='threading',
    cors_allowed_origins='*',
    ping_timeout=300,
    ping_interval=25,
)

UPLOAD_DIR = os.path.abspath('./uploads')
OUTPUT_DIR = os.path.abspath('./output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    job_id  = str(uuid.uuid4())[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'Khong co file nao duoc gui len'}), 400

    saved = 0
    for f in files:
        filename = os.path.basename(f.filename.replace('\\', '/'))
        if filename and not filename.startswith('~$'):
            f.save(os.path.join(job_dir, filename))
            saved += 1

    if saved == 0:
        return jsonify({'error': 'Khong luu duoc file nao. Thu chon lai folder/file.'}), 400
    print(f'[UPLOAD] job={job_id}  saved={saved} files  dir={job_dir}')

    ngay = request.form.get('ngay_doi_chieu', '').strip()

    thread = threading.Thread(
        target=_run_processing,
        args=(job_id, job_dir, ngay or None),
        daemon=True,
    )
    thread.start()

    return jsonify({'job_id': job_id, 'message': 'Dang xu ly...'})


def _run_processing(job_id: str, input_dir: str, ngay: str):
    def emit_log(msg: str):
        socketio.emit('log', {'job_id': job_id, 'msg': msg})

    try:
        emit_log(f'[{job_id}] Bat dau xu ly...')
        output_path = main_from_dir(
            input_dir=input_dir,
            output_dir=OUTPUT_DIR,
            ngay=ngay,
            log_callback=emit_log,
        )

        # Thu thap tat ca file ket qua (xlsx + CSV)
        base = os.path.basename(output_path).replace('.xlsx', '')
        result_files = [{'name': os.path.basename(output_path),
                         'url': f'/download/{os.path.basename(output_path)}'}]
        for fname in os.listdir(OUTPUT_DIR):
            if fname.endswith('.csv') and base.replace('doi_chieu_', '') in fname:
                result_files.append({'name': fname, 'url': f'/download/{fname}'})

        socketio.emit('done', {'job_id': job_id, 'files': result_files})
    except Exception as e:
        socketio.emit('error', {'job_id': job_id, 'msg': str(e)})


@app.route('/download/<filename>')
def download(filename):
    filename = os.path.basename(filename)
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        print(f'[404] {path}  |  OUTPUT_DIR={OUTPUT_DIR}')
        files_in_dir = os.listdir(OUTPUT_DIR) if os.path.isdir(OUTPUT_DIR) else []
        return (f'File khong ton tai: {filename}\n'
                f'Thu muc output: {OUTPUT_DIR}\n'
                f'Cac file hien co: {files_in_dir}'), 404
    return send_file(path, as_attachment=True)


if __name__ == '__main__':
    import socket
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = '127.0.0.1'
    print('\n' + '=' * 50)
    print(f'  Web UI chay tai: http://{ip}:8080')
    print(f'  Tu may khac trong LAN, truy cap dia chi tren.')
    print('=' * 50 + '\n')
    socketio.run(app, host='0.0.0.0', port=8080, debug=False, allow_unsafe_werkzeug=True)
