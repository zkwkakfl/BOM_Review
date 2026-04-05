"""테이블 읽기."""

import tempfile
from pathlib import Path

from bom_review.table_io import (
    list_files_in_folder,
    load_header_and_rows,
    load_header_and_rows_by_sheet_name,
    resolve_column_index,
    values_for_column,
)


def test_load_csv_with_header():
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "t.csv"
        p.write_text("Reference,Qty\nR1, 1\nR2 R3, 2\n", encoding="utf-8-sig")
        h, rows = load_header_and_rows(p, max_data_rows=None)
        assert "Reference" in h
        idx = h.index("Reference")
        assert rows[0][idx] == "R1"


def test_resolve_column_index_strip_and_case():
    h = [" Part ", "REF", "qty"]
    assert resolve_column_index(h, "Part") == 0
    assert resolve_column_index(h, "REF") == 1
    assert resolve_column_index(h, "Qty") == 2


def test_values_for_column():
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "t.csv"
        p.write_text("A,B\n1,2\n3,4\n", encoding="utf-8")
        h, rows = load_header_and_rows(p, max_data_rows=None)
        assert values_for_column(h, rows, "B") == ["2", "4"]


def test_load_header_and_rows_by_sheet_name(tmp_path: Path) -> None:
    from openpyxl import Workbook

    p = tmp_path / "wb.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "DataSheet"
    ws.append(["Ref", "Q"])
    ws.append(["R1", 1])
    wb.save(p)
    wb.close()
    h, rows = load_header_and_rows_by_sheet_name(p, sheet_name="DataSheet", max_data_rows=None)
    assert "Ref" in h
    assert len(rows) == 1


def test_list_files_in_folder_filters_ext():
    with tempfile.TemporaryDirectory() as d:
        base = Path(d)
        (base / "a.csv").write_text("x", encoding="utf-8")
        (base / "b.txt").write_text("x", encoding="utf-8")
        fs = list_files_in_folder(base)
        assert len(fs) == 1
        assert fs[0].name == "a.csv"
