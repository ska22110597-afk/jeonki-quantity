# 전기공사 공량 산출

전기공사 견적·산출용 **PyQt6 데스크톱 앱**입니다.  
단가대비표 엑셀을 끌어다 놓으면 원본은 읽기만 하고, 병합 셀을 채운 뒤 공량산출표 수식이 들어 있는 새 파일을 C드라이브에 저장합니다.

## 안전 규칙

- 드래그 앤 드롭한 **원본 엑셀은 읽기만** 합니다. 수정·덮어쓰기를 하지 않습니다.
- 병합 셀은 원본에서 해제하지 않고, 메모리에서 상단/좌측 대표 값만 채웁니다.
- 기본 저장 경로는 OneDrive를 피하기 위해 **`C:\전기공사_공량산출_결과`** 입니다.
- PyInstaller exe로 묶여도 이 경로는 `sys._MEIPASS` 임시 폴더를 쓰지 않는 **C드라이브 절대 경로**입니다.
- 결과 파일명은 항상  
  `단가대비_공량산출_결과_YYYYMMDD_HHMMSS.xlsx`  
  형태로 **새 파일**만 만듭니다.

Windows가 아닌 환경에서는 같은 이름의 홈 폴더(`~/전기공사_공량산출_결과`)로 저장합니다. 테스트 시에는 `JEONKI_RESULT_DIR`로 경로를 바꿀 수 있습니다.

## 결과 엑셀 구성

결과는 시트 3개입니다. 원본 내역서는 건드리지 않습니다.

| 시트 | 역할 |
| --- | --- |
| 내역서 | 드롭한 내역서. 병합 셀을 채운 뒤 명칭/규격/단위/수량을 정리 |
| 품셈표 | 품목별 노무명칭·품셈·할증%·근거. `C:\전기공사_공량산출_결과\품셈표_데이터베이스.xlsx` 에 계속 쌓임 |
| 공량산출서 | 수식으로 두 시트를 연결 |

공량산출서 수식:

- E 결정수량 = 내역서 수량 셀 (`='내역서'!D행`)
- F 수량할증 % = 기본 0. 품목별로 직접 입력 (5 를 넣으면 5%)
- G 산출수량 = `=E행*(1+F행/100)`
- H 노무 명칭 / I 품셈 / J 노무 할증 % / L 품셈근거 = 품셈표에서 명칭+규격 VLOOKUP
- K 공량 = `산출수량 × 품셈 × (노무할증/100)`

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

1. 단가대비표 `.xlsx` / `.xlsm` 파일을 드롭 존에 놓습니다. (클릭해서 선택도 가능)
2. 원본 경로와 C드라이브 저장 경로를 확인합니다.
3. 저장 경로 확인란을 체크합니다.
4. **공량 산출 및 C드라이브 저장**을 누릅니다.

## 사무용 PC용 단일 .exe 빌드

파이썬이 없는 사무 컴퓨터에서는 아래처럼 **개발 PC에서 exe를 만든 뒤** `dist\전기공사_공량산출.exe` 만 복사하면 됩니다.

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
`C:\Users\ADMIN>` 같은 홈 폴더에서 실행하면 안 됩니다.

Windows에서는 `pyinstaller` 명령이 PATH에 없는 경우가 많습니다.  
그래서 **`python -m PyInstaller`** 또는 **`py -m PyInstaller`** 를 씁니다.

```bat
cd /d 프로젝트폴더경로
python -m pip install pyinstaller pyqt6 pandas openpyxl
python -m PyInstaller --noconsole --onefile --clean --name "전기공사_공량산출" --collect-all PyQt6 --hidden-import openpyxl --hidden-import pandas --hidden-import app --hidden-import app.excel_io --hidden-import app.paths --hidden-import app.merge_parse --hidden-import app.main_window --hidden-import app.drop_zone main.py
```

`python` 이 인식되지 않으면 `python` 대신 `py -3` 을 쓰세요.

```bat
py -3 -m pip install pyinstaller pyqt6 pandas openpyxl
py -3 -m PyInstaller --noconsole --onefile --clean --name "전기공사_공량산출" --collect-all PyQt6 --hidden-import openpyxl --hidden-import pandas --hidden-import app --hidden-import app.excel_io --hidden-import app.paths --hidden-import app.merge_parse --hidden-import app.main_window --hidden-import app.drop_zone main.py
```

또는 프로젝트 폴더의 `build_exe.bat` 을 더블클릭하세요. 이 파일이 Python을 찾고 `python -m PyInstaller`로 빌드합니다.

- `--noconsole` : 검은 CMD 창 없이 GUI만 실행
- `--onefile` : `dist\전기공사_공량산출.exe` 한 개로 묶음

`'pyinstaller'은(는) 내부 또는 외부 명령...` 이 나오면 `pyinstaller` 단독 명령 대신 위처럼 `python -m PyInstaller` 를 쓰면 됩니다.

빌드가 끝나면 `dist\전기공사_공량산출.exe` 를 USB 등으로 복사해 사용합니다. 대상 PC에 파이썬을 설치할 필요는 없습니다.

### 다른 Windows PC가 없을 때 (클라우드)

exe는 **Windows에서만** 만들 수 있습니다.

| 환경 | Windows exe 가능? |
| --- | --- |
| 구글 콜랩, Cloud Shell, 리눅스 VM | 불가 |
| GitHub Actions (`windows-latest`) | 가능 (추천, 보통 무료) |
| 구글클라우드 **Windows** 가상 PC + 원격데스크톱 | 가능 (과금·카드 등록) |

**GitHub에 이 프로젝트를 올린 뒤**

1. 저장소에서 **Actions** 탭을 엽니다.
2. 왼쪽에서 **Build Windows EXE** 를 고릅니다.
3. **Run workflow** → **Run workflow** 를 누릅니다.
4. 초록 체크가 끝나면 해당 실행 화면에서 **Artifacts** 의 `jeonki-quantity-exe` 를 받습니다.
5. 압축을 풀면 `전기공사_공량산출.exe` 가 있습니다. 이걸 사무컴으로 복사하세요.

구글클라우드를 쓰려면 Compute Engine에서 **Windows Server** VM을 만들고, 원격데스크톱으로 접속한 뒤 이 폴더에서 `build_exe.bat` 을 실행하면 됩니다. 리눅스 VM이나 콜랩은 안 됩니다.

### 3) exe에서도 C드라이브 저장이 되는 이유

`--onefile` exe는 실행 시 임시 폴더(`sys._MEIPASS`)에 풀립니다. 이 프로그램은 그 임시 경로를 저장 위치로 쓰지 않고, 항상 `C:\전기공사_공량산출_결과` 절대 경로를 만든 뒤 쓰기 권한을 확인합니다. 폴더 생성·쓰기가 막히면 한글 오류 메시지로 알려 줍니다.

## 테스트

```bash
python -m pytest
```
