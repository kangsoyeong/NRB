@echo off
chcp 65001 > nul
echo ======================================
echo  아마란스 일별시재조회 자동 다운로드
echo  저장 위치: 바탕화면\강소영
echo ======================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 설치 후 다시 실행하세요.
    pause
    exit /b 1
)

echo [1단계] 필수 패키지 설치 중...
pip install selenium webdriver-manager -q
echo 완료.
echo.

echo [2단계] 일별시재조회 자동 다운로드 시작...
echo.
python "%~dp0amaranth_daily_download.py"

if errorlevel 1 (
    echo.
    echo [오류] 실패했습니다. logs 폴더의 스크린샷을 확인하세요.
) else (
    echo.
    echo [완료] 바탕화면\강소영 폴더를 확인하세요!
    explorer "%USERPROFILE%\Desktop\강소영"
)

pause
