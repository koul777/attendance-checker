# Attendance Checker

엑셀로 내려받은 복무/근태 자료에서 지참, 조퇴, 출퇴근 누락, 비근무일 출퇴근 기록 같은 검토 대상을 찾는 로컬 Flask 도구입니다. 파일은 서버에 업로드하지 않고 실행 중인 PC의 임시 폴더에서만 분석합니다.

## 주요 기능

- `.xlsx` 근태 파일 업로드 및 웹 화면 분석
- 정상근무, 시차출퇴근제, 반차, 전일 휴가, 출장, 육아시간, 익일퇴근 기준 반영
- 이상치, 보류, 검토 대상 요약
- 이상 행에 검증결과와 검증사유를 표시한 엑셀 다운로드
- 공개 테스트용 가상 근태 데이터 제공

## 비개발자용 다운로드

비개발자에게 배포할 때는 GitHub 저장소의 소스코드를 직접 받게 하기보다 GitHub Releases에 `MyAttendance.zip` 같은 실행파일 패키지를 올리는 방식을 권장합니다.

사용자는 다음 순서로 실행합니다.

1. GitHub Releases에서 최신 `MyAttendance.zip`을 다운로드합니다.
2. 압축을 풉니다.
3. `MyAttendance.exe`를 실행합니다.
4. 브라우저가 열리면 엑셀 파일을 업로드해 검증합니다.
5. 결과 화면에서 `Export`를 눌러 검증결과가 표시된 엑셀을 내려받습니다.

Windows 보안 경고가 뜨면 게시자가 등록되지 않은 개인 배포 실행파일이기 때문입니다. 조직 내부 배포용으로 쓸 경우 코드 서명 인증서를 적용하면 경고를 줄일 수 있습니다.

릴리스 패키지는 로컬에서 다음 명령으로 만들 수 있습니다.

```powershell
pyinstaller AttendanceChecker.spec
python scripts\create_release_package.py
```

생성된 `release/MyAttendance.zip`을 GitHub Release 첨부 파일로 업로드하세요. `dist/`와 `release/`는 빌드 산출물이므로 저장소 커밋 대상에서 제외합니다.

## 개발자용 빠른 시작

Windows에서 가장 간단히 실행하려면 다음 순서로 진행합니다.

```bat
setup.bat
python app.py
```

브라우저에서 `http://localhost:5000`을 열고 `sample_data/public_attendance_sample.xlsx`를 업로드하면 바로 동작을 확인할 수 있습니다. `python app.py` 실행 시 기본 포트는 `5000`이며, 다른 포트를 쓰려면 `PORT` 환경 변수를 지정합니다.

가상환경을 직접 쓰는 경우:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## 샘플 데이터

공개 저장소에 포함된 샘플은 모두 가상 데이터입니다.

- `sample_data/public_attendance_sample.xlsx`: 바로 업로드 가능한 공개용 샘플 엑셀
- `sample_data/public_attendance_sample_expected.json`: 샘플 분석 기대 결과
- `scripts/create_public_sample.py`: 샘플 엑셀 재생성 스크립트

샘플에는 정상근무, 미승인 지참, 미승인 조퇴, 승인 지참/조퇴, 전일 연가, 오전/오후 반차, 시차출퇴근제, 육아시간, 비근무일 출퇴근 기록, 퇴근 누락, 출퇴근 누락, 익일퇴근 사례가 들어 있습니다.

샘플을 다시 만들려면:

```powershell
python scripts\create_public_sample.py
```

## 입력 엑셀 형식

분석 대상 시트의 첫 20행 안에 아래 헤더가 있어야 합니다.

| 열 이름 | 설명 |
| --- | --- |
| `근태일` | 근무일 |
| `요일` | 요일 |
| `소속명` | 부서 또는 조직명 |
| `교직원` | 직원명 |
| `교직원번호` | 직원 식별번호 |
| `근무유형` | 정상근무, 시차출퇴근제 등 |
| `근무형태` | `정상근무(09:00~18:00)`, `시차A형(07:00~16:00)` 같은 기준 시간 |
| `출근시간` | 실제 출근 시각 |
| `퇴근시간` | 실제 퇴근 시각 |
| `익일여부` | 퇴근이 다음날이면 `1`, 아니면 `0` |
| `체크` | ERP 1차 체크값. 예: `A`, `B`, `A/B` |
| `근무상황` | 승인된 지참, 조퇴, 연가, 출장, 육아시간 등 |

시간은 `08:30`, `8시 30분`, `830`, `8.5`, 엑셀 시간 serial 등을 처리합니다. 근무상황 예시는 `지참(일반)(09:00~10:00)`, `조퇴(일반)(16:00~18:00)`, `연가(오전반차)`, `출장` 형식입니다.

## 검증 명령

```powershell
python -m compileall app.py analyzer.py scripts
python scripts\attendance_harness.py
```

`scripts/attendance_harness.py`는 `rules/attendance_harness.json`과 내부 합성 케이스를 기준으로 판정 로직을 확인합니다. 로컬에 규정 원문 샘플 파일이 없는 환경에서는 일부 외부 샘플 확인이 실패할 수 있습니다.

## 프로젝트 구조

```text
app.py                         Flask 진입점, 업로드/분석/다운로드 처리
analyzer.py                    엑셀 파싱, 시간 변환, 이상치 판정, 표시 엑셀 생성
templates/index.html           업로드 및 결과 화면
static/style.css               사용자 정의 스타일
rules/                         판정 기준 문서와 하네스 기대값
scripts/attendance_harness.py  판정 규칙 검증 하네스
scripts/create_public_sample.py 공개용 가상 데이터 생성
sample_data/                   공개 테스트용 가상 데이터
AttendanceChecker.spec         PyInstaller 패키징 설정
```

## 공개 저장소 사용 시 주의

- 실제 근태 파일, 교직원번호, 조직명, 규정 원문 원본 파일은 공개 저장소에 올리지 마세요.
- 이 저장소의 `sample_data`는 구조 검증을 위한 가상 데이터이며 실제 인물이나 조직과 무관합니다.
- 규정이나 단체협약이 다른 조직에서 쓰려면 `rules/attendance_rules.md`와 `analyzer.py`의 판정 기준을 먼저 확인해야 합니다.
- 라이선스 파일은 아직 포함되어 있지 않습니다. 공개 배포 전에 적용할 라이선스를 선택해 `LICENSE`를 추가하세요.
