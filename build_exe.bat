@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  전기공사 공량산출  —  단일 EXE 빌드
echo ============================================
echo 작업 폴더: %CD%
echo.

if not exist "main.py" (
  echo [오류] main.py 가 이 폴더에 없습니다.
  echo 프로젝트 폴더에 있는 build_exe.bat 을 실행하세요.
  echo 현재 위치: %CD%
  pause
  exit /b 1
)

set "PY="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if not defined PY (
  python -c "import sys" >nul 2>&1
  if not errorlevel 1 set "PY=python"
)
if not defined PY (
  echo [오류] Python 을 찾지 못했습니다.
  echo 1. https://www.python.org/downloads/ 에서 Python 3.10 이상 설치
  echo 2. 설치 화면에서 "Add python.exe to PATH" 체크
  echo 3. 명령 프롬프트를 닫았다가 다시 연 뒤 이 파일을 실행
  pause
  exit /b 1
)

echo [1] 필수 라이브러리 설치  ^(명령: %PY% -m pip^)
%PY% -m pip install --upgrade pip
if errorlevel 1 goto :fail
%PY% -m pip install pyinstaller pyqt6 pandas openpyxl
if errorlevel 1 goto :fail

echo.
echo [2] 콘솔 창 없는 단일 EXE 빌드
echo     python -m PyInstaller 를 사용합니다. ^(PATH에 pyinstaller.exe 가 없어도 됩니다^)
%PY% -m PyInstaller --noconsole --onefile --clean ^
  --name GongryangCalc ^
  --icon assets/app.ico ^
  --version-file file_version_info.txt ^
  --collect-all PyQt6 ^
  --add-data "data;data" ^
  --add-data "assets;assets" ^
  --hidden-import openpyxl ^
  --hidden-import pandas ^
  --hidden-import app ^
  --hidden-import app.excel_io ^
  --hidden-import app.paths ^
  --hidden-import app.merge_parse ^
  --hidden-import app.estimate_parse ^
  --hidden-import app.pumsam ^
  --hidden-import app.pumsam_aliases ^
  --hidden-import app.pumsam_text ^
  --hidden-import app.electric_pumsam_data ^
  --hidden-import app.wages ^
  --hidden-import app.official_wages ^
  --hidden-import app.version ^
  --hidden-import app.discipline ^
  --hidden-import app.items ^
  --hidden-import app.ilwidae ^
  --hidden-import app.pipeline ^
  --hidden-import app.main_window ^
  --hidden-import app.drop_zone ^
  main.py
if errorlevel 1 goto :fail

echo.
echo 완료: %CD%\dist\GongryangCalc.exe
echo 이 파일을 파이썬이 없는 사무용 PC에 복사해 실행하면 됩니다.
echo 결과는 창에서 고른 폴더에 저장됩니다. 기본 후보는 C:\전기공사_공량산출_결과 입니다.
pause
exit /b 0

:fail
echo.
echo [실패] 위 오류를 확인하세요.
pause
exit /b 1
