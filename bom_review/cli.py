"""콘솔 진입점 — GUI 추가 전까지 self-check 및 도움말."""

from __future__ import annotations

import argparse
import sys

from bom_review._version import __version__
from bom_review.matching import bom_vs_source_findings, duplicate_reference_findings


def _pause_if_frozen_without_args() -> None:
    """
    exe 더블클릭 시 인자가 없으면 도움말 출력 후 창이 바로 닫히므로,
    PyInstaller(frozen)이고 argv가 실행 파일만일 때 Enter 대기.
    """
    if not getattr(sys, "frozen", False):
        return
    if len(sys.argv) > 1:
        return
    try:
        input("\n종료하려면 Enter 키를 누르세요 . . .")
    except EOFError:
        pass


def cmd_self_check() -> int:
    """패키지·매칭 로직 스모크 테스트."""
    r = bom_vs_source_findings(["A"], ["A", "B"])
    assert not r.has_errors
    assert len(r.source_only_info) == 1
    dup = duplicate_reference_findings(["X", "X"])
    assert len(dup) == 1
    print("self-check: OK")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="BOM_Review",
        description="BOM·원본좌표·메탈좌표 검토 도구 (콘솔 진입점)",
    )
    p.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    sub = p.add_subparsers(dest="command", metavar="COMMAND")

    sc = sub.add_parser("self-check", help="모듈 로드 및 매칭 함수 스모크 테스트")
    sc.set_defaults(_handler=cmd_self_check)

    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    handler = getattr(args, "_handler", None)
    if handler is not None:
        return int(handler())
    parser.print_help()
    return 0


def run() -> None:
    code = main()
    _pause_if_frozen_without_args()
    raise SystemExit(code)


if __name__ == "__main__":
    run()
