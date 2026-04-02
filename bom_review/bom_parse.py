"""BOM 좌표명 셀 → Reference 토큰 분리 (기본 구분자 ', ')."""

from __future__ import annotations

from typing import Any

__all__ = ["split_designators"]


def split_designators(cell: Any, delimiter: str = ", ") -> list[str]:
    """빈 값은 제외한 설계자(Reference) 문자열 목록."""
    if cell is None:
        return []
    s = str(cell).strip()
    if not s:
        return []
    parts = s.split(delimiter)
    return [p.strip() for p in parts if p.strip()]
