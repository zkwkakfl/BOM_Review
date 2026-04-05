"""검토용 통합문서 스냅샷 기록."""

from pathlib import Path

from bom_review.excel_com import SelectionSourceMeta
from bom_review.excel_snapshot import (
    RANGE_SET_SHEET,
    destination_sheet_name_for_role,
    new_snapshot_workbook_path,
    write_role_range_to_snapshot,
)
from bom_review.table_io import load_header_and_rows_by_sheet_name


def test_destination_sheet_name_for_role() -> None:
    assert destination_sheet_name_for_role("BOM") == "BOM_검토복사"
    assert destination_sheet_name_for_role("원본") == "원본_검토복사"


def test_write_two_roles_updates_range_set(tmp_path: Path) -> None:
    p = new_snapshot_workbook_path(tmp_path)
    meta_b = SelectionSourceMeta("bom.xlsx", "Sheet1", "$A$1:$B$5")
    s1 = write_role_range_to_snapshot(
        p,
        role="BOM",
        headers=["Ref", "Qty"],
        data_rows=[["R1", 1]],
        meta=meta_b,
        create_new_workbook=True,
    )
    assert s1 == "BOM_검토복사"
    meta_s = SelectionSourceMeta("src.xlsx", "좌표", "$C$1:$C$9")
    s2 = write_role_range_to_snapshot(
        p,
        role="원본",
        headers=["Ref"],
        data_rows=[["R1"]],
        meta=meta_s,
        create_new_workbook=True,
    )
    assert s2 == "원본_검토복사"

    h1, r1 = load_header_and_rows_by_sheet_name(p, sheet_name=s1, max_data_rows=None)
    assert "Ref" in h1
    assert r1[0][h1.index("Ref")] == "R1"

    from openpyxl import load_workbook

    wb = load_workbook(p)
    try:
        rs = wb[RANGE_SET_SHEET]
        assert rs.max_row == 3
        roles = [rs.cell(row=i, column=1).value for i in range(2, rs.max_row + 1)]
        assert "BOM" in roles
        assert "원본" in roles
    finally:
        wb.close()


def test_bom_snapshot_normalizes_coord_column_to_comma_space(tmp_path: Path) -> None:
    """BOM 복사 시 지정 좌표명 열은 ', ' 형태로 통일."""
    p = new_snapshot_workbook_path(tmp_path)
    meta = SelectionSourceMeta("b.xlsx", "S", "$A$1:$C$3")
    write_role_range_to_snapshot(
        p,
        role="BOM",
        headers=["Part", "Ref"],
        data_rows=[
            ["X", "R1 R2\tR3"],
            ["Y", "R4,R5"],
            ["Z", "R6;R7"],
        ],
        meta=meta,
        create_new_workbook=True,
        bom_coord_column="Ref",
    )
    h, rows = load_header_and_rows_by_sheet_name(p, sheet_name="BOM_검토복사", max_data_rows=None)
    ri = h.index("Ref")
    assert rows[0][ri] == "R1, R2, R3"
    assert rows[1][ri] == "R4, R5"
    assert rows[2][ri] == "R6, R7"
    pi = h.index("Part")
    assert rows[0][pi] == "X"


def test_reapply_same_role_replaces_range_set_row(tmp_path: Path) -> None:
    p = new_snapshot_workbook_path(tmp_path)
    meta1 = SelectionSourceMeta("a.xlsx", "S", "$A$1:$A$2")
    write_role_range_to_snapshot(
        p,
        role="BOM",
        headers=["H"],
        data_rows=[["1"]],
        meta=meta1,
        create_new_workbook=True,
    )
    meta2 = SelectionSourceMeta("a.xlsx", "S", "$B$1:$B$3")
    write_role_range_to_snapshot(
        p,
        role="BOM",
        headers=["H"],
        data_rows=[["2"]],
        meta=meta2,
        create_new_workbook=True,
    )
    from openpyxl import load_workbook

    wb = load_workbook(p)
    try:
        rs = wb[RANGE_SET_SHEET]
        assert rs.max_row == 2
        assert rs.cell(row=2, column=4).value == "$B$1:$B$3"
    finally:
        wb.close()
