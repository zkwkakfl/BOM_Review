"""
정식 앱 진입점 — 인자 없이 실행 시 GUI가 열립니다.
python main.py  /  python -m bom_review  /  BOM_Review.exe
"""

from bom_review.cli import run

if __name__ == "__main__":
    run()
