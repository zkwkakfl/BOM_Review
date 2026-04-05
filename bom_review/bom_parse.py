"""BOM 좌표명 셀 → Reference 토큰 분리 (기본 구분자 ', ')."""

from __future__ import annotations

import re
from typing import Any

__all__ = [
    "normalize_designators_to_comma_space",
    "split_designators",
    "tokenize_designators_loose",
]

# 쉼표·세미콜론·공백(탭/줄바꿈 포함)을 구분자로 간주해 토큰 분리
_LOOSE_DESIGNATOR_SPLIT = re.compile(r"[\s,;]+")


def tokenize_designators_loose(cell: Any) -> list[str]:
    """셀 값을 쉼표·공백·세미콜론 등으로 나눈 Reference 토큰 (검토용 복사본 정규화 등)."""
    if cell is None:
        return []
    s = str(cell).strip()
    if not s:
        return []
    return [p for p in _LOOSE_DESIGNATOR_SPLIT.split(s) if p]


def normalize_designators_to_comma_space(cell: Any) -> Any:
    """
    여러 구분 형태를 ', ' 로 통일한 문자열로 만든다.
    빈 셀·공백만 있는 셀은 빈 문자열, None은 None 유지.
    """
    if cell is None:
        return None
    raw = str(cell)
    if not raw.strip():
        return ""
    tokens = tokenize_designators_loose(raw)
    if not tokens:
        return ""
    return ", ".join(tokens)


def split_designators(cell: Any, delimiter: str = ", ") -> list[str]:
    """빈 값은 제외한 설계자(Reference) 문자열 목록."""
    if cell is None:
        return []
    s = str(cell).strip()
    if not s:
        return []
    parts = s.split(delimiter)
    return [p.strip() for p in parts if p.strip()]
