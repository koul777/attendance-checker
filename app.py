import atexit
import io
import os
import posixpath
import secrets
import shutil
import socket
import sys
import tempfile
import threading
import webbrowser
import zipfile

from flask import Flask, after_this_request, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from analyzer import analyze_file, generate_marked_workbook


APP_VERSION = '2026-06-29-samples-1-6'
DEFAULT_PORT = 5000
DEFAULT_MAX_UPLOAD_MB = 10
LOCAL_ADDRESSES = {'127.0.0.1', '::1'}
TOKEN_REQUIRED_PATHS = {'/upload', '/download'}
MAX_XLSX_ZIP_ENTRIES = 500
MAX_XLSX_UNCOMPRESSED_SIZE = 50 * 1024 * 1024
LOCAL_TOKEN = secrets.token_urlsafe(32)


def get_base_path():
    """PyInstaller 패키지 리소스 경로 처리"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


base_path = get_base_path()
app = Flask(
    __name__,
    template_folder=os.path.join(base_path, 'templates'),
    static_folder=os.path.join(base_path, 'static'),
)
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('ATTENDANCE_MAX_UPLOAD_MB', str(DEFAULT_MAX_UPLOAD_MB))) * 1024 * 1024

UPLOAD_FOLDER = tempfile.mkdtemp()
UPLOADED_FILE = os.path.join(UPLOAD_FOLDER, 'uploaded.xlsx')
MARKED_FILE = os.path.join(UPLOAD_FOLDER, 'marked_attendance.xlsx')
atexit.register(lambda: shutil.rmtree(UPLOAD_FOLDER, ignore_errors=True))


class InvalidXlsxError(ValueError):
    pass


def remove_file(path):
    try:
        os.remove(path)
    except FileNotFoundError:
        pass


def clear_previous_files():
    remove_file(UPLOADED_FILE)
    remove_file(MARKED_FILE)


def has_unsafe_zip_name(name):
    normalized = name.replace('\\', '/')
    return (
        normalized.startswith('/')
        or (len(normalized) >= 2 and normalized[1] == ':')
        or any(part == '..' for part in normalized.split('/'))
    )


def validate_xlsx_file(path):
    if not zipfile.is_zipfile(path):
        raise InvalidXlsxError('올바른 .xlsx 파일이 아닙니다. 엑셀 파일을 다시 확인해주세요.')

    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_XLSX_ZIP_ENTRIES:
                raise InvalidXlsxError('엑셀 파일 내부 항목이 너무 많습니다.')

            total_size = 0
            names = set()
            for info in infos:
                normalized = info.filename.replace('\\', '/')
                if has_unsafe_zip_name(info.filename):
                    raise InvalidXlsxError('엑셀 파일 내부 경로가 안전하지 않습니다.')
                names.add(posixpath.normpath(normalized))
                total_size += info.file_size
                if total_size > MAX_XLSX_UNCOMPRESSED_SIZE:
                    raise InvalidXlsxError('엑셀 파일 압축 해제 크기가 너무 큽니다.')

            if '[Content_Types].xml' not in names or 'xl/workbook.xml' not in names:
                raise InvalidXlsxError('올바른 .xlsx 파일 구조가 아닙니다.')
    except zipfile.BadZipFile as exc:
        raise InvalidXlsxError('올바른 .xlsx 파일이 아닙니다. 엑셀 파일을 다시 확인해주세요.') from exc


@app.before_request
def protect_local_requests():
    if request.remote_addr not in LOCAL_ADDRESSES:
        return jsonify({'error': '이 앱은 로컬 PC에서만 사용할 수 있습니다.'}), 403

    if request.method == 'POST' and request.path in TOKEN_REQUIRED_PATHS:
        token = request.headers.get('X-Local-Token', '')
        if not secrets.compare_digest(token, LOCAL_TOKEN):
            return jsonify({'error': '잘못된 로컬 요청입니다.'}), 403
    return None


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(_exc):
    return jsonify({'error': '업로드 파일 크기가 너무 큽니다.'}), 413


@app.route('/')
def index():
    return render_template('index.html', app_version=APP_VERSION, local_token=LOCAL_TOKEN)


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
        return jsonify({'error': '지원하는 파일 형식은 .xlsx입니다.'}), 400

    clear_previous_files()
    try:
        file.save(UPLOADED_FILE)
        validate_xlsx_file(UPLOADED_FILE)
        result = analyze_file(UPLOADED_FILE)
        result['app_version'] = APP_VERSION
        result['summary']['app_version'] = APP_VERSION
        return jsonify(result)
    except InvalidXlsxError as exc:
        remove_file(UPLOADED_FILE)
        return jsonify({'error': str(exc)}), 400
    except Exception:
        app.logger.exception('Failed to analyze uploaded workbook')
        remove_file(UPLOADED_FILE)
        return jsonify({'error': '분석 중 오류가 발생했습니다. 파일 형식과 내용을 확인해주세요.'}), 500


@app.route('/download', methods=['POST'])
def download():
    if not os.path.exists(UPLOADED_FILE):
        return jsonify({'error': '먼저 .xlsx 파일을 업로드해주세요.'}), 400

    output_path = MARKED_FILE
    try:
        generate_marked_workbook(UPLOADED_FILE, output_path)
        with open(output_path, 'rb') as generated:
            download_data = io.BytesIO(generated.read())
        download_data.seek(0)
    except Exception:
        app.logger.exception('Failed to generate marked workbook')
        remove_file(output_path)
        return jsonify({'error': '파일 생성 중 오류가 발생했습니다.'}), 500

    @after_this_request
    def cleanup_download(response):
        remove_file(output_path)
        return response

    return send_file(
        download_data,
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
    app.run(host='127.0.0.1', port=port, debug=False)
