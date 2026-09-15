@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  GitHub 저장소에 올리기
echo  주소: https://github.com/ska22110597-afk/jeonki-quantity
echo ============================================
echo 지금 폴더: %CD%
echo.

git --version >nul 2>&1
if errorlevel 1 (
  echo [안내] 이 PC에 Git 이 없습니다.
  echo 1. https://git-scm.com/download/win  에서 Git 설치
  echo 2. 설치 후 이 파일을 다시 더블클릭
  echo.
  echo 더 쉬운 방법: https://desktop.github.com  에서 GitHub Desktop 설치
  pause
  exit /b 1
)

if not exist "main.py" (
  echo [오류] main.py 가 없습니다. C:\공량산출 프로그램 폴더에서 실행하세요.
  pause
  exit /b 1
)

if not exist ".git" (
  echo [1] 이 폴더를 Git 저장소로 만듭니다.
  git init
  if errorlevel 1 goto :fail
)

echo [2] 파일 준비
git add .
git status

echo [3] 저장
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "전기공사 공량산출"
) else (
  echo 새로 올릴 변경이 없거나, 이미 커밋되어 있습니다.
  git commit -m "전기공사 공량산출" >nul 2>&1
)

git branch -M main

git remote get-url github >nul 2>&1
if errorlevel 1 (
  echo [4] GitHub 주소 연결
  git remote add github https://github.com/ska22110597-afk/jeonki-quantity.git
)

echo [5] GitHub로 올리기
echo 로그인 창이 나오면 GitHub 계정으로 로그인하세요.
git push -u github main
if errorlevel 1 goto :fail

echo.
echo 성공. 브라우저에서 확인하세요:
echo https://github.com/ska22110597-afk/jeonki-quantity
echo 그다음 Actions 탭에서 Build Windows EXE 를 실행하면 됩니다.
pause
exit /b 0

:fail
echo.
echo 실패했습니다. 아래를 확인하세요.
echo - GitHub에 로그인되어 있는지
echo - 저장소 주소가 맞는지
echo - 인터넷이 되는지
pause
exit /b 1
