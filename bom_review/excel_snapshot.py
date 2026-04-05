"""Excel 범위 선택 후 타임스탬프 결과 통합문서에 복사하고 Range_Set 시트를 유지한다."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from bom_review.bom_parse import normalize_designators_to_comma_space
from bom_review.excel_com import SelectionSourceMeta

RANGE_SET_SHEET = "Range_Set"
RANGE_SET_HEADERS = ("역할", "원본파일", "원본시트", "원본주소", "복사시트")


def new_snapshot_workbook_path(folder: Path) -> Path:
    """작업 폴더에 저장할 결과 통합문서 경로(타임스탬프 파일명)."""
    folder = folder.resolve()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return folder / f"BOM_Review_Result_{ts}.xlsx"


def destination_sheet_name_for_role(role: str) -> str:
    """역할별 복사 대상 시트 이름(Excel 제한 내)."""
    if role == "BOM":
        return "BOM_검토복사"
    if role == "원본":
        return "원본_검토복사"
    base = role.replace("[", "").replace("]", "").replace("*", "").replace("?", "").replace("/", "").replace("\\", "")[:28]
    return base or "데이터"


def _ensure_range_set_sheet(wb: Any) -> Any:
    """Range_Set 시트가 없으면 맨 앞에 만든다."""
    if RANGE_SET_SHEET in wb.sheetnames:
        return wb[RANGE_SET_SHEET]
    ws = wb.create_sheet(RANGE_SET_SHEET, 0)
    ws.append(list(RANGE_SET_HEADERS))
    return ws


def _upsert_range_set_row(wb: Any, role: str, meta: SelectionSourceMeta, copy_sheet: str) -> None:
    """동일 역할 행은 제거한 뒤 새 메타 한 줄을 추가한다."""
    rs = _ensure_range_set_sheet(wb)
    kept: list[list[Any]] = []
    for row in rs.iter_rows(min_row=2, values_only=True):
        if not row or all(c is None for c in row):
            continue
        r0 = row[0]
        if r0 is not None and str(r0).strip() == role:
            continue
        kept.append([row[i] if i < len(row) else None for i in range(len(RANGE_SET_HEADERS))])
    if rs.max_row >= 2:
        rs.delete_rows(2, rs.max_row - 1)
    for r in kept:
        rs.append(r)
    rs.append([role, meta.source_file, meta.source_sheet, meta.source_address, copy_sheet])


def _bom_rows_with_normalized_coord_column(
    headers: list[str],
    data_rows: list[list[Any]],
    coord_column: str,
) -> list[list[Any]]:
    """BOM 복사 시 좌표명 열만 ', ' 구분으로 통일한 데이터 행."""
    try:
        idx = headers.index(coord_column)
    except ValueError:
        return data_rows
    out: list[list[Any]] = []
    w = len(headers)
    for row in data_rows:
        new_r = list(row)
        if len(new_r) < w:
            new_r.extend([None] * (w - len(new_r)))
        new_r = new_r[:w]
        new_r[idx] = normalize_designators_to_comma_space(new_r[idx])
        out.append(new_r)
    return out


def write_role_range_to_snapshot(
    snapshot_path: Path,
    *,
    role: str,
    headers: list[str],
    data_rows: list[list[Any]],
    meta: SelectionSourceMeta,
    create_new_workbook: bool,
    bom_coord_column: str | None = None,
) -> str:
    """
    결과 통합문서에 헤더+데이터 시트를 쓰고 Range_Set을 갱신한다.

    create_new_workbook가 True이면 파일이 없을 때 새로 만든다.
    role이 BOM이고 bom_coord_column이 헤더에 있으면 해당 열만 좌표명을 ', ' 형태로 정규화해 복사한다.
    반환: 복사된 데이터가 들어 있는 시트 이름.
    """
    snapshot_path = snapshot_path.resolve()
    dest_name = destination_sheet_name_for_role(role)

    rows_to_write = data_rows
    if role == "BOM" and bom_coord_column and headers:
        rows_to_write = _bom_rows_with_normalized_coord_column(
            headers, data_rows, bom_coord_column
        )

    if snapshot_path.exists():
        wb = load_workbook(snapshot_path)
    else:
        if not create_new_workbook:
            raise FileNotFoundError(f"결과 통합문서가 없습니다: {snapshot_path}")
        wb = Workbook()
        default = wb.active
        default.title = RANGE_SET_SHEET
        default.append(list(RANGE_SET_HEADERS))

    _ensure_range_set_sheet(wb)

    if dest_name in wb.sheetnames:
        wb.remove(wb[dest_name])
    ws_data = wb.create_sheet(dest_name)
    ws_data.append(list(headers))
    for row in rows_to_write:
        ws_data.append(list(row))

    _upsert_range_set_row(wb, role, meta, dest_name)

    snapshot_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(snapshot_path)
    wb.close()
    return dest_name
