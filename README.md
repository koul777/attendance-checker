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

## 다른 ERP와 기관 규정에 맞게 수정 개발하기

다른 기관에서 이 도구를 재사용할 때는 이 저장소를 Codex 작업 폴더로 지정하고, 기관별 ERP 샘플과 규정은 로컬 컨텍스트로만 제공해 수정합니다. 이 과정은 모델을 별도로 학습시키는 작업이 아니라, 마스킹한 예시 파일과 규정 문서를 근거로 `analyzer.py`, `rules/attendance_rules.md`, `scripts/attendance_harness.py`를 보완하는 개발 작업입니다.

### 1. 공개 시스템 샘플로 기본 동작 확인

```powershell
git clone <이 저장소 주소> attendance-checker
cd attendance-checker
setup.bat
python app.py
```

브라우저에서 `sample_data/public_attendance_sample.xlsx`를 업로드해 기본 동작을 확인합니다. 이 파일은 공개용 가상 데이터이며, 실제 인물, 조직, 사번, 근태 기록과 무관합니다.

공개 시스템 샘플을 다시 만들려면 다음 명령을 실행합니다.

```powershell
python scripts\create_public_sample.py
python scripts\attendance_harness.py
```

이후 Codex에서 작업 폴더를 이 프로젝트 루트(`attendance-checker`)로 지정합니다. 기관별 맞춤 개발을 요청할 때도 이 폴더를 기준으로 작업하게 해야 기존 Flask 앱, 분석 로직, 하네스, 공개 샘플을 함께 읽고 수정할 수 있습니다.

### 2. 기관 ERP 샘플 마스킹

ERP에서 실제로 내려받은 근태 Raw data를 복사한 뒤, 공개 가능한 개발 샘플로 바꿉니다. 최소 10-30행 정도라도 정상근무, 지참, 조퇴, 반차, 전일 휴가, 출장, 유연근무, 출퇴근 누락, 비근무일 기록처럼 기관에서 검증해야 할 사례를 포함하는 것이 좋습니다.

마스킹 기준은 다음과 같습니다.

- 직원명은 `직원001`, `직원002`처럼 가명으로 바꿉니다.
- 교직원번호, 사번, 개인 식별 ID는 `E0001`, `E0002` 같은 임의 코드로 바꿉니다.
- 소속명, 조직명, 상세 부서명은 `샘플부서A`, `샘플부서B`처럼 바꿉니다.
- 메모, 비고, 승인 사유에 개인 사유가 있으면 삭제하거나 `개인사유`처럼 일반화합니다.
- 날짜는 요일과 근무 패턴이 유지되도록 필요하면 가상 월로 이동합니다.
- 숨김 시트, 필터에 감춰진 행, 메모, 문서 속성에 원본 정보가 남아 있지 않은지 확인합니다.

마스킹한 파일은 `private_data/<기관명>/masked_attendance_sample.xlsx`처럼 저장하거나 저장소 밖의 로컬 폴더에 둡니다. `private_data/`, `local_data/`, `uploads/`, 실제 엑셀 파일, 엑셀 잠금 파일, 빌드 산출물은 `.gitignore` 대상이므로 커밋에 포함하지 않습니다.

ERP마다 컬럼명이 다를 수 있으므로 마스킹 샘플과 함께 컬럼 대응표를 만들어 두면 수정이 빠릅니다. 예를 들어 어떤 ERP는 `교직원번호` 대신 `사번`, `소속명` 대신 `부서`, `근태일` 대신 `일자`, `근무상황` 대신 `신청구분`이나 `휴가/근태명`을 쓸 수 있습니다.

```text
이 도구 기준 컬럼        기관 ERP 컬럼 예시
근태일                  일자, 근무일자, 기준일
요일                    요일
소속명                  부서, 조직, 소속부서
교직원                  성명, 직원명, 이름
교직원번호              사번, 직원번호, 개인번호
근무유형                근무제도, 근무구분
근무형태                근무시간, 근무스케줄, 출퇴근유형
출근시간                출근, 출근시각, 시작시간
퇴근시간                퇴근, 퇴근시각, 종료시간
익일여부                익일, 다음날퇴근, 야간근무여부
체크                    ERP체크, 이상구분, 점검결과
근무상황                근태명, 신청구분, 휴가근태, 복무상황
```

컬럼명이 다르면 엑셀을 이 도구의 표준 컬럼명으로 수작업 변경하기보다, `analyzer.py`에서 여러 헤더 별칭을 인식하도록 수정하는 편이 재사용에 유리합니다.

### 3. 규정 문서 준비

복무규정, 유연근무 규정, 단체협약, 임금피크제 규정, 육아시간 기준처럼 판정에 필요한 문서는 HTML 또는 텍스트로 변환해 Codex가 읽을 수 있게 준비합니다. 원본 PDF, HWP, DOCX, 내부 규정 원문은 로컬 작업 자료로만 둡니다.

권장 형태는 다음과 같습니다.

```text
private_data/<기관명>/rules/
  attendance_rule.html
  flexible_work_rule.html
  collective_agreement.html
```

규정 문서가 길면 전체 원문 대신 판정에 필요한 조항만 별도 텍스트로 추려도 됩니다. 다만 조항 번호, 적용 대상, 예외 조건, 시간 기준은 남겨야 하네스 기대값을 고정할 수 있습니다.

### 4. Codex에 수정 요청하기

Codex에는 샘플 파일과 규정 위치, 원하는 결과, 검증 명령을 함께 전달합니다. 예시는 다음과 같습니다.

```text
이 프로젝트 폴더를 기준으로 다른 기관 ERP에 맞게 수정해줘.

- 마스킹 샘플: private_data/abc/masked_attendance_sample.xlsx
- 규정 HTML: private_data/abc/rules/
- 컬럼 대응: 일자=근태일, 사번=교직원번호, 부서=소속명, 성명=교직원, 출근=출근시간, 퇴근=퇴근시간, 신청구분=근무상황
- 목표: 새 ERP 컬럼명과 근무상황명을 인식하고, 지참/조퇴/반차/출장/유연근무 판정 기준을 이 기관 규정에 맞게 보완
- 작업 범위: analyzer.py, rules/attendance_rules.md, scripts/attendance_harness.py, 필요 시 README.md
- 검증: python scripts\attendance_harness.py, python -m compileall app.py analyzer.py scripts

먼저 현재 코드와 마스킹 샘플의 실제 헤더를 확인한 뒤, 컬럼 별칭 매핑, 상태명 파싱, 규정별 판정 기준, 하네스 기대값을 함께 수정해줘.
공개 저장소에 실제 원본 파일이나 개인정보가 들어가지 않도록 확인해줘.
```

기관별 수정에서는 보통 아래 항목을 손봅니다.

- ERP 컬럼명이 다르면 `analyzer.py`의 헤더 인식 로직에 별칭 매핑을 추가합니다.
- 필수 컬럼이 여러 시트에 나뉘어 있거나 헤더가 20행 이후에 있으면 시트 선택과 헤더 탐색 범위를 조정합니다.
- 근무상황명이 다르면 지참, 조퇴, 연가, 병가, 출장, 육아시간, 반차 등 상태 파싱 규칙을 추가합니다.
- 유연근무 유형명이 다르면 `근무형태` 시간 범위 추출 또는 유형별 기본 시간을 보완합니다.
- 규정상 점심시간, 공동근무시간, 반차 계산, 익일퇴근, 임금피크 단축근무 기준이 다르면 판정 함수를 수정합니다.
- 새 규칙마다 `rules/attendance_rules.md`에 근거와 판정 기준을 요약합니다.
- 마스킹 샘플에서 기대되는 이상치, 보류, 검토 행 번호를 하네스에 고정합니다.

### 5. 검증과 커밋 전 점검

수정 후에는 다음 명령을 실행합니다.

```powershell
python scripts\attendance_harness.py
python -m compileall app.py analyzer.py scripts
git status --short
```

커밋하기 전에는 다음 항목을 확인합니다.

- `git status --short`에 `private_data/`, 실제 엑셀 파일, 규정 원문 파일, `build/`, `dist/`, `release/`, `__pycache__/`가 보이지 않아야 합니다.
- 공개 샘플은 실제 사람이 아닌 가상 데이터여야 합니다.
- `rules/attendance_rules.md`에는 원문 전체가 아니라 판정 기준 요약만 남깁니다.
- 기관 내부 배포용 실행 파일은 `pyinstaller AttendanceChecker.spec`와 `python scripts\create_release_package.py`로 별도 생성합니다.

여러 기관에서 계속 쓸 계획이면 기관별 변경을 한 번에 섞기보다, 기관별 브랜치나 별도 설정 파일 방식으로 분리하는 편이 유지보수에 유리합니다.

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

`scripts/attendance_harness.py`는 `sample_data/public_attendance_sample_expected.json`과 내부 합성 케이스를 기준으로 판정 로직을 확인합니다.

## 프로젝트 구조

```text
app.py                         Flask 진입점, 업로드/분석/다운로드 처리
analyzer.py                    엑셀 파싱, 시간 변환, 이상치 판정, 표시 엑셀 생성
templates/index.html           업로드 및 결과 화면
static/style.css               사용자 정의 스타일
rules/                         판정 기준 문서
scripts/attendance_harness.py  판정 규칙 검증 하네스
scripts/create_public_sample.py 공개용 가상 데이터 생성
sample_data/                   공개 테스트용 가상 데이터와 기대값
AttendanceChecker.spec         PyInstaller 패키징 설정
```
