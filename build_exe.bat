@echo off
chcp 65001 >nul
setlocal

echo [1] 필수 라이브러리 설치
python -m pip install --upgrade pip
python -m pip install pyinstaller pyqt6 pandas openpyxl

echo.
echo [2] 콘솔 없는 단일 EXE 빌드
python -m PyInstaller --noconsole --onefile --clean ^
  --name "전기공사_공량산출" ^
  --collect-all PyQt6 ^
  --hidden-import openpyxl ^
  --hidden-import pandas ^
  --hidden-import app ^
  --hidden-import app.excel_io ^
  --hidden-import app.paths ^
  --hidden-import app.merge_parse ^
  --hidden-import app.main_window ^
  --hidden-import app.drop_zone ^
  main.py

echo.
echo 완료: dist\전기공사_공량산출.exe
echo 이 파일을 파이썬이 없는 사무용 PC에 복사해 실행하면 됩니다.
echo 결과는 C:\전기공사_공량산출_결과 에 저장됩니다.
pause
