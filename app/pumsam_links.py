"""사용자가 직접 정한 「이 단가대비표 품목 = 이 품셈」 연결. 데이터베이스 폴더의 JSON 에 쌓아 다음부터 자동 적용한다."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app.estimate_parse import lookup_key

LINKS_FILE_NAME = "품셈_연결목록.json"
_active: dict[str, list[str]] = {}


def links_path(db_dir: Path | None) -> Path:
    from app.paths import user_database_dir

    folder = user_database_dir(db_dir)
    folder.mkdir(parents=True, exist_ok=True)
    return folder / LINKS_FILE_NAME


def load_links(db_dir: Path | None) -> dict[str, list[str]]:
    try:
        data = json.loads(links_path(db_dir).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return {str(k): list(v) for k, v in data.items() if isinstance(v, list) and len(v) == 2}


def save_link(db_dir: Path | None, name: Any, spec: Any, book_name: str, book_spec: str) -> None:
    links = load_links(db_dir)
    links[lookup_key(name, spec)] = [str(book_name), str(book_spec)]
    links_path(db_dir).write_text(json.dumps(links, ensure_ascii=False, indent=1), encoding="utf-8")


def activate(db_dir: Path | None) -> None:
    global _active
    _active = load_links(db_dir)


def linked_target(name: Any, spec: Any) -> tuple[str, str] | None:
    value = _active.get(lookup_key(name, spec))
    return (value[0], value[1]) if value else None
