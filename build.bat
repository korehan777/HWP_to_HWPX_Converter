@echo off
chcp 65001 > nul
echo =====================================
echo HWP to HWPX Converter - 빌드 스크립트
echo =====================================
echo.

REM 가상환경이 있으면 활성화
if exist .venv\Scripts\activate.bat (
    echo 가상환경 활성화 중...
    call .venv\Scripts\activate.bat
)

echo 빌드 시작...
echo.

REM PyInstaller로 빌드
pyinstaller --onefile ^
    --windowed ^
    --name="HWP_to_HWPX_Converter" ^
    --add-data "converter.py;." ^
    --add-data "utils.py;." ^
    --clean ^
    main.py

echo.
echo =====================================
echo 빌드 완료!
echo 실행 파일 위치: dist\HWP_to_HWPX_Converter.exe
echo =====================================
echo.

pause
