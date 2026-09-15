# 전기공사 공량 산출

전기공사 견적·산출용 **PyQt6 데스크톱 앱**입니다.  
내역서 엑셀을 끌어다 놓으면 원본은 읽기만 하고, 내역서·품셈표·공량산출서가 수식으로 연결된 새 파일을 만듭니다.

## 안전 규칙

- 드래그 앤 드롭한 **원본 엑셀은 읽기만** 합니다. 수정·덮어쓰기를 하지 않습니다.
- 병합 셀은 원본에서 해제하지 않고, 메모리에서 상단/좌측 대표 값만 채웁니다.
- 기본 저장 경로는 OneDrive를 피하기 위해 **`C:\전기공사_공량산출_결과`** 입니다. 창에서 **폴더 찾기**로 바꿀 수 있습니다.
- PyInstaller exe로 묶여도 실행 파일 임시 폴더(`sys._MEIPASS`)에는 저장하지 않습니다.
- 결과 파일명은 항상 `공량산출_결과_YYYYMMDD_HHMMSS.xlsx` 형태로 **새 파일**만 만듭니다.

Windows가 아닌 환경에서는 같은 이름의 홈 폴더(`~/전기공사_공량산출_결과`)가 기본값입니다. 테스트 시에는 `JEONKI_RESULT_DIR`로 경로를 바꿀 수 있습니다.

## 결과 엑셀 구성

결과는 시트 3개입니다. 서식은 샘플처럼 **바탕 흰색, 행 높이 20, 자동 줄 바꿈 없음**입니다.

| 시트 | 역할 |
| --- | --- |
| 내역서 | 드롭한 내역서. 원본 행 번호를 유지하고 병합을 그대로 둡니다 |
| 품셈표 | 품목별 노무명칭·품셈·할증%·근거. 저장 폴더의 `품셈표_데이터베이스.xlsx` 에 계속 쌓임 |
| 공량산출서 | 같은 행의 내역서 수량과 품셈표를 수식으로 연결 |

공량산출서 수식(샘플과 동일, 행 번호 일치):

- A 검색키 = `CONCATENATE(B,C)`
- G 산출수량 = 같은 행 내역서 수량 (`='내역서'!D행`)
- F 할증 = 기본 0 (엑셀에서 5%면 `0.05` 또는 `5%`)
- E 결정수량 = `TRUNC(F*G+G, 0)`
- H 노무 명칭 / I 품셈 / J 할증% / L 품셈근거 = 품셈표 VLOOKUP
- K 공량 = `내역서 수량 × 품셈 × (할증%/100)`

품셈표에 없는 품목은 H/I/J가 비고, 공량도 비어 있습니다. 품셈표에 행을 추가하거나 품셈표 엑셀을 드롭하면 다음 산출부터 붙습니다.

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

1. 내역서 `.xlsx` / `.xlsm` 파일을 드롭 존에 놓습니다. (클릭해서 선택도 가능)
2. 필요하면 품셈표 엑셀을 두 번째 칸에 놓아 데이터베이스에 합칩니다.
3. **폴더 찾기**로 결과 저장 위치를 정합니다. 칸에 경로를 직접 붙여 넣어도 됩니다.
4. 저장 확인란을 체크합니다.
5. **공량 산출 및 저장**을 누릅니다.

## 앞으로 붙일 기능 (단가대비표 → 일위대가 → 내역서 → 공량산출서)

지금은 **내역서가 이미 있을 때 공량산출서를 만드는** 단계입니다.  
최종 순서는 아래처럼 한 장씩 붙이는 편이 안전합니다. 품셈표는 처음부터 끝까지 쓰는 **마스터 데이터베이스**로 둡니다.

1. **품셈표 DB** — 지금처럼 계속 최신화해서 쌓기 (노무 품셈·할증·근거)
2. **단가대비표** — 자재 단가 입력/관리 (다음 단계)
3. **일위대가** — 단가대비표 재료비 + 품셈표 노무비로 호표 블록 작성
4. **내역서** — 일위대가 단가 × 수량으로 자동 작성
5. **공량산출서** — 지금 만든 수식 연결 유지

한 번에 네 장을 자동으로 만들면, 품셈·단가·호표 규칙이 어긋났을 때 어디서 틀렸는지 찾기 어렵습니다.  
품셈표를 먼저 쌓고, 그다음 단가대비표 → 일위대가 → 내역서 순으로 붙이면 각 단계 엑셀을 눈으로 확인할 수 있습니다.

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
python -m PyInstaller --noconsole --onefile --clean --name GongryangCalc --collect-all PyQt6 --hidden-import openpyxl --hidden-import pandas --hidden-import app --hidden-import app.excel_io --hidden-import app.paths --hidden-import app.merge_parse --hidden-import app.estimate_parse --hidden-import app.pumsam --hidden-import app.main_window --hidden-import app.drop_zone main.py
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
