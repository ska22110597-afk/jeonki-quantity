# 전기공사 공량 산출

전기공사 견적·산출용 **PyQt6 데스크톱 앱 1단계 뼈대**입니다.  
단가대비표 엑셀을 끌어다 놓으면 원본은 읽기만 하고, 결과는 C드라이브 로컬 폴더에 타임스탬프 파일로만 저장합니다.

## 안전 규칙

- 드래그 앤 드롭한 **원본 엑셀은 Read-Only**로만 엽니다. 수정·덮어쓰기를 하지 않습니다.
- 기본 저장 경로는 OneDrive를 피하기 위해 **`C:\전기공사_공량산출_결과`** 입니다.
- 결과 파일명은 항상  
  `단가대비_공량산출_결과_YYYYMMDD_HHMMSS.xlsx`  
  형태로 **새 파일**만 만듭니다.

Windows가 아닌 환경에서는 같은 이름의 홈 폴더(`~/전기공사_공량산출_결과`)로 저장합니다. 테스트 시에는 `JEONKI_RESULT_DIR`로 경로를 바꿀 수 있습니다.

## 결과 엑셀 구성

| 시트 | 이름 | 내용 |
| --- | --- | --- |
| 1 | 단가대비표 | 원본 첫 시트를 빈 행·열만 정리해 복사 |
| 2 | 공량산출표 | 번호, 공종, 품명, 규격, 단위, 산출근거, 수량, 비고 기본 골격 |

## 실행 방법

Python 3.10 이상이 필요합니다.

```bash
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

## 테스트

```bash
python -m pytest
```
