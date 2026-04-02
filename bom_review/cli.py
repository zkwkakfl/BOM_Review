"""앱 진입점 — 기본은 정식 GUI, 부가로 demo / self-check."""

from __future__ import annotations

import argparse
import sys

from bom_review._version import __version__
from bom_review.matching import bom_vs_source_findings, duplicate_reference_findings


def _pause_if_frozen_subcommand() -> None:
    """
    exe에서 서브커맨드(demo 등) 실행 후 콘솔이 곧 닫히지 않도록 Enter 대기.
    인자 없음(더블클릭) → GUI만 띄우므로 여기서는 대기하지 않음.
    """
    if not getattr(sys, "frozen", False):
        return
    if len(sys.argv) <= 1:
        return
    if hasattr(sys.stdout, "isatty") and not sys.stdout.isatty():
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


def cmd_gui() -> int:
    """폴더·파일·열 매핑 후 실제 데이터 검토 (Tkinter)."""
    from bom_review.gui import run_gui

    run_gui()
    return 0


def cmd_demo() -> int:
    """샘플 데이터로 매칭·중복 검사 동작을 콘솔에 출력 (동작 확인용)."""
    print(f"BOM_Review {__version__} — 매칭 데모\n")

    bom_refs = ["R1", "R2", "GHOST"]
    src_refs = ["R1", "R2", "TP1", "FID1"]
    report = bom_vs_source_findings(bom_refs, src_refs)

    print("■ BOM Reference (펼친 뒤 예시):", ", ".join(bom_refs))
    print("■ 원본 Reference (예시):", ", ".join(src_refs))
    print()

    print("[오류] BOM에만 있고 원본에 없음")
    if report.bom_only_errors:
        for f in report.bom_only_errors:
            print(f"  - {f.reference}: {f.message} ({f.code})")
    else:
        print("  (없음)")
    print()

    print("[참고] 원본에만 있음 — 정책상 오류 아님")
    if report.source_only_info:
        for f in report.source_only_info:
            print(f"  - {f.reference}: {f.message} ({f.code})")
    else:
        print("  (없음)")
    print()

    dup_list = ["U1", "U2", "U1"]
    dups = duplicate_reference_findings(dup_list)
    print("[오류] 좌표명 중복 (파일 전체 유일 위반, 예시)")
    for f in dups:
        print(f"  - {f.reference}: {f.message} / {f.detail}")
    print()

    print("— 단위 테스트: 프로젝트 루트에서  python -m pytest tests -v")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="BOM_Review",
        description="BOM·원본좌표 검토 도구 — 기본 실행은 정식 화면(GUI)입니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "실행 예:\n"
            "  python -m bom_review            정식 GUI (인자 없음과 동일)\n"
            "  python -m bom_review gui\n"
            "  python main.py                    위와 동일\n"
            "  BOM_Review.exe                    빌드 exe — 정식 GUI\n"
            "\n"
            "개발·점검용:\n"
            "  python -m bom_review demo         콘솔 샘플 출력\n"
            "  python -m bom_review self-check   스모크 테스트\n"
            "  python -m pytest tests -v"
        ),
    )
    p.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    sub = p.add_subparsers(dest="command", metavar="COMMAND")

    gu = sub.add_parser("gui", help="정식 GUI (기본과 동일)")
    gu.set_defaults(_handler=cmd_gui)

    dm = sub.add_parser("demo", help="샘플 BOM/원본으로 매칭 결과 출력 (동작 확인)")
    dm.set_defaults(_handler=cmd_demo)

    sc = sub.add_parser("self-check", help="모듈 로드 및 매칭 함수 스모크 테스트")
    sc.set_defaults(_handler=cmd_self_check)

    return p


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv_rest = sys.argv[1:]
    else:
        argv_rest = argv

    # 정식 앱: 인자 없으면 항상 GUI
    if len(argv_rest) == 0:
        return cmd_gui()

    parser = build_parser()
    args = parser.parse_args(argv_rest if argv is not None else None)
    handler = getattr(args, "_handler", None)
    if handler is not None:
        return int(handler())
    parser.print_help()
    return 0


def run() -> None:
    code = main()
    _pause_if_frozen_subcommand()
    raise SystemExit(code)


if __name__ == "__main__":
    run()
