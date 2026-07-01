import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_NAME = "근태 점검 지원 프로그램(My Attendace)"
EXE_NAME = f"{APP_NAME}.exe"
ZIP_NAME = f"{APP_NAME}.zip"
EXE_PATH = ROOT / "dist" / EXE_NAME
SAMPLE_PATH = ROOT / "sample_data" / "public_attendance_sample.xlsx"
RELEASE_DIR = ROOT / "release"
ZIP_PATH = RELEASE_DIR / ZIP_NAME

USER_GUIDE = f"""{APP_NAME} 사용방법

1. {EXE_NAME}를 실행합니다.
2. 브라우저가 자동으로 열리면 근태 엑셀(.xlsx)을 업로드합니다.
3. 검증 결과를 확인합니다.
4. Export 버튼을 눌러 검증결과가 표시된 엑셀을 내려받습니다.

참고
- 이 프로그램은 PC에서 로컬로 실행됩니다.
- 실제 근태 파일은 GitHub나 외부 서버에 업로드하지 마세요.
- public_attendance_sample.xlsx는 테스트용 가상 데이터입니다.
- Windows 보안 경고가 뜨면 '추가 정보'를 누른 뒤 실행할 수 있습니다.
"""


def main():
    if not EXE_PATH.exists():
        raise SystemExit(
            f"실행파일이 없습니다: {EXE_PATH}\n"
            "먼저 `pyinstaller AttendanceChecker.spec`를 실행하세요."
        )

    RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(EXE_PATH, EXE_NAME)
        if SAMPLE_PATH.exists():
            archive.write(SAMPLE_PATH, "public_attendance_sample.xlsx")
        archive.writestr("사용방법.txt", USER_GUIDE)

    print(f"created {ZIP_PATH}")


if __name__ == "__main__":
    main()
