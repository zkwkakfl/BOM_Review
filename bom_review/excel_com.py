"""
Excel COM (Windows) — 통합문서를 띄우고 Selection 값을 2차원으로 읽는다.
pywin32 필요: pip install pywin32
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SelectionSourceMeta:
    """범위 선택 당시 원본 통합문서·시트·주소(감사·추적용)."""

    source_file: str
    source_sheet: str
    source_address: str

EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xlsb", ".xls"})


def is_excel_path(path: Path) -> bool:
    return path.suffix.lower() in EXCEL_SUFFIXES


def normalize_com_value(val: Any) -> list[list[Any]]:
    """Excel Range.Value → 균일한 2차원 리스트 (빈 셀은 None)."""
    if val is None:
        return []
    if not isinstance(val, tuple):
        return [[val]]
    # 한 행만 선택된 경우 COM이 (v1, v2, ...) 1차원 튜플로 줄 때가 많음
    if val and not any(isinstance(x, tuple) for x in val):
        return [list(val)]
    rows: list[list[Any]] = []
    for row in val:
        if isinstance(row, tuple):
            rows.append(list(row))
        else:
            rows.append([row])
    return rows


def _pad_rows(rows: list[list[Any]], width: int) -> list[list[Any]]:
    out: list[list[Any]] = []
    for r in rows:
        row = list(r)
        if len(row) < width:
            row.extend([None] * (width - len(row)))
        out.append(row[:width])
    return out


def read_selection_as_header_and_rows(xl: Any) -> tuple[list[str], list[list[Any]]] | None:
    """
    현재 Excel Selection → (헤더명 목록, 데이터 행들).
    첫 행을 헤더로 쓴다. 선택이 없거나 비어 있으면 None.
    """
    try:
        sel = xl.Selection
        val = sel.Value
    except Exception:  # noqa: BLE001
        return None
    rows = normalize_com_value(val)
    if not rows:
        return None
    w = max(len(r) for r in rows)
    rows = _pad_rows(rows, w)
    header_cells = rows[0]
    headers = [
        str(c).strip() if c is not None and str(c).strip() else f"열{i}"
        for i, c in enumerate(header_cells)
    ]
    data = rows[1:]
    if data:
        data = _pad_rows(data, len(headers))
    return headers, data


def read_selection_source_meta(xl: Any) -> SelectionSourceMeta:
    """
    현재 Selection 기준으로 원본 파일명·시트명·절대 참조 주소를 읽는다.
    COM 오류 시 placeholder로 채운다.
    """
    try:
        sel = xl.Selection
        ws = sel.Worksheet
        sheet_name = str(ws.Name)
        addr = str(sel.Address(RowAbsolute=True, ColumnAbsolute=True))
        parent_wb = ws.Parent
        try:
            full = str(parent_wb.FullName)
            src_name = Path(full).name
        except Exception:  # noqa: BLE001
            src_name = str(parent_wb.Name)
        return SelectionSourceMeta(
            source_file=src_name,
            source_sheet=sheet_name,
            source_address=addr,
        )
    except Exception:  # noqa: BLE001
        return SelectionSourceMeta(source_file="?", source_sheet="?", source_address="?")


def open_workbook_in_new_excel(path: Path) -> tuple[Any, Any]:
    """새 Excel 인스턴스를 띄우고 통합문서를 연다. (xl, wb)"""
    try:
        from win32com.client import DispatchEx
    except ImportError as e:
        raise RuntimeError(
            "Excel 연동에 pywin32 가 필요합니다. 다음을 실행하세요: pip install pywin32"
        ) from e
    path = path.resolve()
    xl = DispatchEx("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
    return xl, wb


def close_excel_quietly(xl: Any, rw: Any) -> None:
    """저장 없이 통합문서 닫기 및 Excel 종료."""
    try:
        if rw is not None:
            rw.Close(SaveChanges=False)
    except Exception:  # noqa: BLE001
        pass
    try:
        if xl is not None:
            xl.Quit()
    except Exception:  # noqa: BLE001
        pass
