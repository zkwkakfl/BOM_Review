"""BOM·원본좌표 검토 패키지 — `python -m bom_review` 로 정식 GUI 실행."""

from bom_review._version import __version__

from bom_review.matching import (
    Finding,
    FindingKind,
    MatchReport,
    bom_vs_source_findings,
    duplicate_reference_findings,
)

__all__ = [
    "__version__",
    "Finding",
    "FindingKind",
    "MatchReport",
    "bom_vs_source_findings",
    "duplicate_reference_findings",
]
