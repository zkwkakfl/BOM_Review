"""콘솔 진입점 — GUI 추가 전까지 self-check 및 도움말."""

from __future__ import annotations

import argparse
import sys

from bom_review._version import __version__
from bom_review.matching import bom_vs_source_findings, duplicate_reference_findings


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
    sys.exit(main())


if __name__ == "__main__":
    run()
