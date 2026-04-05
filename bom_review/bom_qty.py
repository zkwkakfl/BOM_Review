"""BOM 행별: 분리된 좌표명 개수 vs 수량 열."""

from __future__ import annotations

from typing import Any

from bom_review.bom_parse import tokenize_designators_loose
from bom_review.matching import Finding, FindingKind

__all__ = ["bom_quantity_mismatch_findings"]


def _parse_qty(v: Any) -> int | None:
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def bom_quantity_mismatch_findings(
    ref_cells: list[Any],
    qty_cells: list[Any],
    *,
    delimiter: str = ", ",
) -> list[Finding]:
    """
    같은 행 인덱스 기준으로 좌표명 분리 개수와 수량이 다르면 ERROR.
    수량 셀이 비어 있으면 해당 행은 건너뜀.

    좌표명은 쉼표·세미콜론·공백(탭 등)으로 나뉜 토큰을 모두 센다.
    (UI «구분자»와 무관하게 결과 파일 ', ' 정규화와 같은 규칙 — 셀에 «R1 R2»만 있어도 2개로 본다.)
    delimiter 인자는 하위 호환용으로만 받으며 사용하지 않는다.
    """
    n = min(len(ref_cells), len(qty_cells))
    out: list[Finding] = []
    for i in range(n):
        q = _parse_qty(qty_cells[i])
        if q is None:
            continue
        refs = tokenize_designators_loose(ref_cells[i])
        if len(refs) != q:
            out.append(
                Finding(
                    kind=FindingKind.ERROR,
                    code="BOM_QTY_MISMATCH",
                    message="BOM 행: 분리된 좌표명 개수와 수량 불일치",
                    reference=None,
                    detail=f"행 {i + 2} (데이터 기준): 좌표 {len(refs)}개, 수량 {q}",
                )
            )
    return out
