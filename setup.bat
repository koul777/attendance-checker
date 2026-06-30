@echo off
chcp 65001 >nul
echo ========================================
echo   근태 이상치 검출 시스템 - 초기 설정
echo ========================================
echo.
echo 필요 패키지를 설치합니다...
pip install -r requirements.txt
echo.
echo 설치가 완료되었습니다!
echo 이제 '근태검사.bat'을 실행하세요.
echo.
pause
