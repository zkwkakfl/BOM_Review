"""
BOM ↔ 원본좌표 매칭 규칙.

정책(요구사항):
- 원본좌표(PCB 기준)가 진실이다.
- BOM에만 있고 원본에 없는 Reference → 오류(FindingKind.ERROR).
- 원본에만 있고 BOM에는 없는 Reference → 오류가 아님(보고는 참고용 FindingKind.INFO).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Iterable, Iterator


class FindingKind(Enum):
    """결과 시트 분류용. INFO는 오류가 아님."""

    ERROR = auto()  # 반드시 조치·확인 대상
    WARNING = auto()  # 확인 권장
    INFO = auto()  # 정책상 허용, 참고 (원본에만 있음 등)


@dataclass(frozen=True)
class Finding:
    kind: FindingKind
    code: str
    message: str
    reference: str | None = None
    detail: str | None = None


@dataclass
class MatchReport:
    """BOM ↔ 원본 매칭 요약."""

    bom_only_errors: list[Finding] = field(default_factory=list)
    """BOM에는 있으나 원본에 없음 → 오류."""

    source_only_info: list[Finding] = field(default_factory=list)
    """원본에만 있고 BOM에는 없음 → 오류 아님(참고)."""

    @property
    def has_errors(self) -> bool:
        return any(f.kind == FindingKind.ERROR for f in self.bom_only_errors)


def bom_vs_source_findings(
    bom_references: Iterable[str],
    source_references: Iterable[str],
) -> MatchReport:
    """
    BOM에서 펼쳐진 Reference 집합과 원본 Reference 집합을 비교한다.

    - bom - source → ERROR (BOM에 있는데 원본에 없음)
    - source - bom → INFO (원본에만 있음, 오류 아님)
    """
    bom_set = {r.strip() for r in bom_references if r and str(r).strip()}
    src_set = {r.strip() for r in source_references if r and str(r).strip()}

    report = MatchReport()

    for ref in sorted(bom_set - src_set):
        report.bom_only_errors.append(
            Finding(
                kind=FindingKind.ERROR,
                code="BOM_NOT_IN_SOURCE",
                message="BOM에 있으나 원본좌표에 없음",
                reference=ref,
            )
        )

    for ref in sorted(src_set - bom_set):
        report.source_only_info.append(
            Finding(
                kind=FindingKind.INFO,
                code="SOURCE_ONLY",
                message="원본에만 존재(BOM 미기재, 정책상 오류 아님)",
                reference=ref,
            )
        )

    return report


def duplicate_reference_findings(
    references: Iterable[str],
    *,
    scope_label: str = "파일 전체",
) -> list[Finding]:
    """
    Reference 문자열 목록에서 중복을 찾는다. (파일 전체 유일 정책)

    반환되는 Finding은 모두 ERROR.
    """
    seen: dict[str, int] = {}
    order: list[str] = []
    for r in references:
        key = (r or "").strip()
        if not key:
            continue
        if key not in seen:
            order.append(key)
        seen[key] = seen.get(key, 0) + 1

    out: list[Finding] = []
    for ref in order:
        if seen[ref] > 1:
            out.append(
                Finding(
                    kind=FindingKind.ERROR,
                    code="DUPLICATE_REFERENCE",
                    message=f"{scope_label}에서 좌표명(Reference) 중복",
                    reference=ref,
                    detail=f"출현 횟수: {seen[ref]}",
                )
            )
    return out


def iter_error_findings(report: MatchReport) -> Iterator[Finding]:
    """오류 시트용: ERROR 등급만 (원본에만 있음 INFO는 제외)."""
    yield from report.bom_only_errors


def iter_info_findings(report: MatchReport) -> Iterator[Finding]:
    """결과 요약 시 참고용: 원본에만 있음 등."""
    yield from report.source_only_info
