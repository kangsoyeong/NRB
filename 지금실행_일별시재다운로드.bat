@echo off
chcp 65001 > nul
echo ======================================
echo  e-branch 일별시재조회 지금 실행
echo  저장 위치: 바탕화면\강소영
echo ======================================
echo.

:: Python 설치 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python을 설치한 후 다시 실행하세요.
    pause
    exit /b 1
)

:: 필수 패키지 설치
echo [1단계] 필수 패키지 설치 중...
pip install pyautogui pygetwindow pillow -q
echo 완료.
echo.

:: images 폴더 확인
set IMAGES_DIR=%~dp0images
if not exist "%IMAGES_DIR%" (
    echo [2단계] UI 이미지 캡처 필요
    echo images 폴더가 없습니다. capture_ui_images.py 를 먼저 실행해주세요.
    echo.
    echo 지금 캡처 도구를 실행할까요?
    set /p RUN_CAPTURE=  [Y/N] 입력:
    if /i "%RUN_CAPTURE%"=="Y" (
        python "%~dp0capture_ui_images.py"
    )
    echo.
)

:: 메인 스크립트 실행
echo [3단계] 일별시재조회 자동 다운로드 시작...
echo.
python "%~dp0ebranch_daily_download.py"

if errorlevel 1 (
    echo.
    echo [오류] 다운로드 중 오류가 발생했습니다.
    echo logs 폴더의 로그 파일을 확인하세요.
) else (
    echo.
    echo [완료] 바탕화면\강소영 폴더를 확인하세요!
    explorer "%USERPROFILE%\Desktop\강소영"
)

pause
