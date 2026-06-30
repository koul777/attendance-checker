import os
import socket
import sys
import tempfile
import threading
import webbrowser

from flask import Flask, jsonify, render_template, request, send_file

from analyzer import analyze_file, generate_marked_workbook


APP_VERSION = '2026-06-29-samples-1-6'
DEFAULT_PORT = 5000


def get_base_path():
    """PyInstaller 패키징 시 리소스 경로 처리"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


base_path = get_base_path()
app = Flask(
    __name__,
    template_folder=os.path.join(base_path, 'templates'),
    static_folder=os.path.join(base_path, 'static'),
)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

UPLOAD_FOLDER = tempfile.mkdtemp()
UPLOADED_FILE = os.path.join(UPLOAD_FOLDER, 'uploaded.xlsx')


@app.route('/')
def index():
    return render_template('index.html', app_version=APP_VERSION)


@app.route('/health')
def health():
    return jsonify({'ok': True, 'app_version': APP_VERSION})


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': '파일을 선택해주세요.'}), 400

    if not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': '엑셀 파일(.xlsx)만 업로드 가능합니다.'}), 400

    file.save(UPLOADED_FILE)

    try:
        result = analyze_file(UPLOADED_FILE)
        result['app_version'] = APP_VERSION
        result['summary']['app_version'] = APP_VERSION
        return jsonify(result)
    except Exception as exc:
        return jsonify({'error': f'분석 중 오류가 발생했습니다: {exc}'}), 500


@app.route('/download', methods=['POST'])
def download():
    if not os.path.exists(UPLOADED_FILE):
        return jsonify({'error': '먼저 엑셀 파일을 업로드해주세요.'}), 400

    output_path = os.path.join(UPLOAD_FOLDER, 'marked_attendance.xlsx')
    try:
        generate_marked_workbook(UPLOADED_FILE, output_path)
    except Exception as exc:
        return jsonify({'error': f'엑셀 표시 중 오류가 발생했습니다: {exc}'}), 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name='복무관리_검증결과.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


def open_browser(port):
    webbrowser.open(f'http://localhost:{port}')


def port_available(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex(('127.0.0.1', port)) != 0


def select_port(preferred_port):
    for port in range(preferred_port, preferred_port + 20):
        if port_available(port):
            return port
    raise RuntimeError(f'사용 가능한 포트를 찾지 못했습니다: {preferred_port}-{preferred_port + 19}')


if __name__ == '__main__':
    preferred_port = int(os.environ.get('PORT', str(DEFAULT_PORT)))
    port = select_port(preferred_port)
    if os.environ.get('ATTENDANCE_CHECKER_OPEN_BROWSER', '1') == '1':
        threading.Timer(1.0, open_browser, args=(port,)).start()
    app.run(host='0.0.0.0', port=port, debug=False)
