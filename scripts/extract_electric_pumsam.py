#!/usr/bin/env python3
"""전기 표준품셈 PDF 추출. 공통 추출기를 호출한다."""

from pathlib import Path

from scripts.extract_pumsam_book import extract

PDF = Path("/tmp/pumsam/2025_elec.pdf")
OUT = Path("/workspace/data/전기품셈_원표.json")


def main() -> None:
    from scripts.extract_pumsam_book import main as shared_main
    import sys

    sys.argv = [
        "extract_pumsam_book.py",
        "--discipline",
        "전기",
        "--pdf",
        str(PDF),
        "--out",
        str(OUT),
    ]
    shared_main()


if __name__ == "__main__":
    main()
