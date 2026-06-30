@echo off
chcp 65001 >nul
echo ========================================
echo   근태 이상치 검출 시스템 시작
echo ========================================
echo.
echo 최신 실행 파일을 실행합니다.
echo 브라우저가 자동으로 열립니다.
echo 종료하려면 실행 파일 창을 닫으세요.
echo.
cd /d "%~dp0"
if exist "%~dp0dist\MyAttendance.exe" (
    "%~dp0dist\MyAttendance.exe"
) else (
    echo dist\MyAttendance.exe 파일이 없어 Python 개발 서버로 실행합니다.
    python app.py
)
