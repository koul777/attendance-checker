# 저장소 작업 지침

## 프로젝트 구조와 모듈
이 저장소는 엑셀 복무/근태 자료에서 이상치를 찾는 Flask 기반 도구입니다.

- `app.py`: 업로드, 분석, 다운로드, 브라우저 실행을 담당하는 Flask 진입점입니다.
- `analyzer.py`: 엑셀 파싱, 시간 변환, 이상치 판정, 표시된 엑셀 생성 로직을 둡니다.
- `templates/index.html`: 파일 업로드와 결과 확인 화면입니다.
- `static/style.css`: Bootstrap 위에 얹는 사용자 정의 스타일입니다.
- `rules/`: 복무 규정 근거와 샘플 기대값을 관리합니다.
- `scripts/attendance_harness.py`: 샘플 엑셀 기준으로 판정 규칙을 검증하는 하네스입니다.
- `AttendanceChecker.spec`: PyInstaller 패키징 설정입니다.

## 개발, 실행, 검증 명령
- `setup.bat`: `flask`, `openpyxl` 실행 의존성을 설치합니다.
- `python app.py`: 로컬 서버를 실행합니다. 기본 포트는 `5000`이며 `PORT`로 바꿀 수 있습니다.
- `python scripts/attendance_harness.py`: 규정 하네스로 샘플 1~6의 기대 결과와 합성 회귀 케이스를 검증합니다.
- `python -m compileall app.py analyzer.py scripts`: 파이썬 문법을 빠르게 확인합니다.
- `pyinstaller AttendanceChecker.spec`: Windows 실행 파일을 빌드합니다.

## 코딩 스타일과 명명 규칙
Python은 4칸 들여쓰기를 사용합니다. 라우트 처리는 `app.py`에 얇게 두고, 복무 판정 로직은 `analyzer.py` 또는 하네스에서 검증된 규칙을 기준으로 옮깁니다. 함수와 변수는 `snake_case`, 상수는 `UPPER_SNAKE_CASE`를 사용합니다.

한글 UI 문구, 엑셀 헤더, 근무상황명은 UTF-8 기준으로 다룹니다. 깨진 문자열을 그대로 확장하지 말고, 실제 샘플 파일의 헤더와 상태명을 기준으로 수정합니다.

## 테스트와 하네스 지침
현재 정식 `tests/` 디렉터리는 없습니다. 복무 규칙 변경 전후에는 반드시 `python scripts/attendance_harness.py`를 실행해 `rules/attendance_harness.json`의 기대 행 번호와 비교합니다.

새 규칙을 추가할 때는 `rules/attendance_rules.md`에 규정 근거와 판정 기준을 먼저 적고, 샘플 행 번호를 `rules/attendance_harness.json`에 고정합니다. 현재 하네스 기준 문서는 `4-4-1. 복무규정.html`, `4-4-2. 유연근무제 운영세칙.html`, `4-3-2. 임금피크제 운영 세칙.html`입니다.

특히 시차근무, 근무시간선택제, 임금피크제 단축근무, 육아시간, 출장, 연가, 반차, 익일퇴근은 하네스 기대값으로 검증해야 합니다. 반차는 기준 출근 시각부터 점심시간을 제외한 4시간 근무 후 퇴근 가능으로 계산합니다.

## 커밋과 PR 지침
이 체크아웃에서는 Git 이력을 확인할 수 없습니다. 커밋 메시지는 짧은 명령형으로 작성합니다. 예: `analyzer: 반차 퇴근 기준 보정`.

PR에는 변경 요약, 실행한 하네스/수동 검증 결과, UI 변경 시 화면 캡처, 패키징 영향 여부를 포함합니다.
