"""BOM 행별: 분리된 좌표명 개수 vs 수량 열."""

from __future__ import annotations

from typing import Any

from bom_review.bom_parse import split_designators
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
    delimiter: str,
) -> list[Finding]:
    """
    같은 행 인덱스 기준으로 좌표명 분리 개수와 수량이 다르면 ERROR.
    수량 셀이 비어 있으면 해당 행은 건너뜀.
    """
    n = min(len(ref_cells), len(qty_cells))
    out: list[Finding] = []
    for i in range(n):
        q = _parse_qty(qty_cells[i])
        if q is None:
            continue
        refs = split_designators(ref_cells[i], delimiter=delimiter)
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
