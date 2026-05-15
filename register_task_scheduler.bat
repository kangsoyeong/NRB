@echo off
:: e-branch 일별시재조회 자동 다운로드 - Windows 작업 스케줄러 등록
:: 관리자 권한으로 실행하세요

set TASK_NAME=eBranch_일별시재_자동다운로드
set PYTHON_PATH=python
set SCRIPT_PATH=%~dp0ebranch_daily_download.py
set RUN_TIME=09:00

echo [작업 스케줄러 등록 중...]
echo 작업명: %TASK_NAME%
echo 스크립트: %SCRIPT_PATH%
echo 실행시간: 매일 오전 %RUN_TIME%
echo.

schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "%PYTHON_PATH% \"%SCRIPT_PATH%\"" ^
  /sc DAILY ^
  /st %RUN_TIME% ^
  /ru "%USERDOMAIN%\%USERNAME%" ^
  /rl HIGHEST ^
  /f

if %errorlevel% equ 0 (
    echo.
    echo [완료] 작업 스케줄러에 등록되었습니다!
    echo 매일 오전 09:00 에 자동으로 실행됩니다.
) else (
    echo.
    echo [오류] 등록 실패. 관리자 권한으로 다시 실행하세요.
)

echo.
pause
