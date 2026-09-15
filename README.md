# 전기공사 공량 산출

전기공사 견적·산출용 **PyQt6 데스크톱 앱**입니다.  
단가대비표를 놓으면 일위대가·내역서·공량산출서까지 수식으로 연결된 새 파일을 만듭니다.  
원본은 읽기만 합니다.

## 안전 규칙

- 드래그 앤 드롭한 **원본 엑셀은 읽기만** 합니다. 수정·덮어쓰기를 하지 않습니다.
- 병합 셀은 원본에서 해제하지 않고, 메모리에서 상단/좌측 대표 값만 채웁니다.
- 기본 저장 경로는 OneDrive를 피하기 위해 **`C:\전기공사_공량산출_결과`** 입니다. 창에서 **폴더 찾기**로 바꿀 수 있습니다.
- PyInstaller exe로 묶여도 실행 파일 임시 폴더(`sys._MEIPASS`)에는 저장하지 않습니다.
- 결과 파일명은 항상 `공량산출_결과_YYYYMMDD_HHMMSS.xlsx` 형태로 **새 파일**만 만듭니다.

Windows가 아닌 환경에서는 같은 이름의 홈 폴더(`~/전기공사_공량산출_결과`)가 기본값입니다. 테스트 시에는 `JEONKI_RESULT_DIR`로 경로를 바꿀 수 있습니다.

## 프로그램 창

드롭 칸은 세 개입니다.

| 칸 | 역할 |
| --- | --- |
| 단가대비표 | 물량 품목의 자재 단가. 놓으면 이후 단계를 모두 작성 |
| 일위대가 | 이미 만든 일위대가가 있을 때 (선택) |
| 내역서 | 이미 만든 내역서만 놓고 공량산출서를 만들 때 |

## 결과 엑셀 구성

서식은 샘플처럼 **바탕 흰색, 행 높이 20, 자동 줄 바꿈 없음**입니다.

| 시트 | 역할 |
| --- | --- |
| 단가대비표 | 드롭한 자재 단가표 |
| 일위대가 | 호표. 전선관이면 전선관부속품비·잡재료비·공구손료 + 표준품셈 인부 |
| 내역서 | 일위대가 단가 × 단가대비표 수량 |
| 공량산출서 | 내역서 수량으로 몇 인이 필요한지 산출 |
| 품셈표 | 표준품셈 데이터베이스 복사 |
| 노임단가 | 직종별 노임 |

일위대가 규칙:

- 전선관: 자재 + 전선관부속품비(15%) + 잡재료비(2%) + 품셈 인부 + 공구손료(3%)
- 케이블/전선: 자재 + 잡재료비(2%) + 품셈 인부 + 공구손료(3%)
- 그 외: 자재 + 품셈 인부 + 공구손료(3%)
- 표준품셈에 내선전공·보통인부처럼 인부가 여러 명이면 호표에 **모두** 넣습니다.

## 표준품셈 · 노임단가 파일

프로그램 안에 `data/표준품셈.xlsx`, `data/노임단가.xlsx` 씨앗 파일이 들어 있습니다. exe로 묶을 때도 같이 실립니다.

처음 산출하면 저장 폴더에 복사본이 생깁니다.

`C:\전기공사_공량산출_결과\데이터베이스\표준품셈.xlsx`  
`C:\전기공사_공량산출_결과\데이터베이스\노임단가.xlsx`

이 복사본을 고치면 다음 산출부터 반영됩니다. 대한전기협회 표준품셈을 정리해 행을 계속 추가하면 됩니다. 명칭·규격이 같고 인부만 다른 행(또는 아래 칸만 인부가 있는 행)은 같은 호표에 인부가 여러 명으로 들어갑니다.

## 실행 방법 (개발 PC)

Python 3.10 이상이 필요합니다.

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

## 사용 순서

1. **단가대비표**를 첫 칸에 놓습니다. (클릭 선택도 가능)
2. 이미 있는 일위대가·내역서가 있으면 해당 칸에 놓습니다.
3. **폴더 찾기**로 결과 저장 위치를 정합니다.
4. 저장 확인란을 체크합니다.
5. **산출 및 저장**을 누릅니다.

노임단가는 `데이터베이스\노임단가.xlsx` 에서 직종별로 고칩니다. 내선전공 샘플 값은 예시이므로 최신 노임으로 바꿔 주세요.

## 사무용 PC용 단일 .exe 빌드

파이썬이 없는 사무 컴퓨터에서는 아래처럼 **개발 PC에서 exe를 만든 뒤** `dist\GongryangCalc.exe` 만 복사하면 됩니다.

### 1) 필수 라이브러리 설치

개발 PC 터미널(명령 프롬프트 또는 PowerShell):

```bat
python -m pip install pyinstaller pyqt6 pandas openpyxl
```

또는 프로젝트 폴더에서:

```bat
python -m pip install -r requirements.txt
```

### 2) 콘솔 창 없는 단일 파일 빌드

**반드시 `main.py`가 있는 프로젝트 폴더에서** 실행하세요.

Windows에서는 `pyinstaller` 명령이 PATH에 없는 경우가 많습니다.  
그래서 **`python -m PyInstaller`** 또는 **`py -m PyInstaller`** 를 씁니다.

```bat
cd /d 프로젝트폴더경로
python -m pip install pyinstaller pyqt6 pandas openpyxl
python -m PyInstaller --noconsole --onefile --clean --name GongryangCalc --collect-all PyQt6 --add-data "data;data" --hidden-import openpyxl --hidden-import pandas --hidden-import app --hidden-import app.excel_io --hidden-import app.paths --hidden-import app.merge_parse --hidden-import app.estimate_parse --hidden-import app.pumsam --hidden-import app.wages --hidden-import app.items --hidden-import app.ilwidae --hidden-import app.pipeline --hidden-import app.main_window --hidden-import app.drop_zone main.py
```

`python` 이 인식되지 않으면 `python` 대신 `py -3` 을 쓰세요.

또는 프로젝트 폴더의 `build_exe.bat` 을 더블클릭하세요.

- `--noconsole` : 검은 CMD 창 없이 GUI만 실행
- `--onefile` : `dist\GongryangCalc.exe` 한 개로 묶음

빌드가 끝나면 `dist\GongryangCalc.exe` 를 USB 등으로 복사해 사용합니다. 대상 PC에 파이썬을 설치할 필요는 없습니다.

### 다른 Windows PC가 없을 때 (클라우드)

exe는 **Windows에서만** 만들 수 있습니다.

**GitHub에 이 프로젝트를 올린 뒤**

1. 저장소에서 **Actions** 탭을 엽니다.
2. 왼쪽에서 **Build Windows EXE** 를 고릅니다.
3. **Run workflow** → **Run workflow** 를 누릅니다.
4. 초록 체크가 끝나면 해당 실행 화면에서 **Artifacts** 의 `jeonki-quantity-exe` 를 받습니다.
5. 압축을 풀면 `GongryangCalc.exe` 가 있습니다. 이걸 사무컴으로 복사하세요.

### 3) exe에서 저장 폴더

기본 후보는 `C:\전기공사_공량산출_결과` 입니다. 창의 **폴더 찾기**로 다른 로컬 폴더를 고를 수 있습니다. 실행 파일 임시 경로에는 저장하지 않습니다.

## 테스트

```bash
python -m pytest
```
