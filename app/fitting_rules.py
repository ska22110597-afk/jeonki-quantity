"""전선관부속품비 제외 규칙 (규칙 4).

전선관 본체에만 부속품비(전선관의 15%)를 붙인다. 아래 낱말이 품명·규격에 들어 있는 부속 자재와,
사용자가 데이터베이스 폴더의 `부속품비_제외목록.txt` 에 적은 낱말은 부속품비 행을 넣지 않는다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

EXCLUDE_FILE_NAME = "부속품비_제외목록.txt"

DEFAULT_FITTING_EXCLUDES: tuple[str, ...] = (
    "부품",
    "부속",
    "커넥터",
    "커플링",
    "노말밴드",
    "엘보",
    "행거",
    "클램프",
    "새들",
    "로크너트",
    "부싱",
    "박스",
    # 표준품셈 3-1-2(통신)/5-2(전기) 부속품률 표에 없는 관 — 부속품은 실소요량 별도 계상
    "나사없는",
    "나사 없는",
    "광케블내외관",
    "파상형",
    "파형관",
    "ELP",
    "이중벽",
    "흄관",
    "도관",
    "관로",
)

_SEED_TEXT = (
    "# 전선관부속품비(15%)를 넣지 않을 품명·규격 낱말을 한 줄에 하나씩 적습니다.\n"
    "# 품명이나 규격에 이 낱말이 들어 있으면 부속품비 행을 만들지 않습니다. # 로 시작하는 줄은 설명입니다.\n"
    "# 예) 광케블내외관일체형  /  파상형  /  관로구\n"
)


def load_fitting_excludes(directory: Path | None) -> tuple[str, ...]:
    """기본 낱말 + 사용자가 적은 낱말. 파일이 없으면 안내용 파일을 만든다."""
    words = list(DEFAULT_FITTING_EXCLUDES)
    if directory is None:
        return tuple(words)
    path = Path(directory) / EXCLUDE_FILE_NAME
    try:
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(_SEED_TEXT, encoding="utf-8")
            return tuple(words)
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and line not in words:
                words.append(line)
    except OSError:
        pass
    return tuple(words)


def fitting_excluded(name: Any, spec: Any, excludes: tuple[str, ...] = DEFAULT_FITTING_EXCLUDES) -> bool:
    text = f"{name or ''} {spec or ''}".replace(" ", "")
    return any(word.replace(" ", "") in text for word in excludes if word.strip())
