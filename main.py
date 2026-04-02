"""
PyInstaller 등으로 패키징할 때 사용하는 루트 진입점.
개발 시: python main.py | python -m bom_review
"""

from bom_review.cli import run

if __name__ == "__main__":
    run()
