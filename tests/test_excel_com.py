"""excel_com 순수 함수 (COM 없음)."""

from pathlib import Path

import pytest

from bom_review.excel_com import is_excel_path, normalize_com_value


def test_is_excel_path() -> None:
    assert is_excel_path(Path("a.XLSX"))
    assert is_excel_path(Path("b.xls"))
    assert not is_excel_path(Path("c.csv"))


@pytest.mark.parametrize(
    ("val", "expected_rows"),
    [
        (None, []),
        (1, [[1]]),
        (("a", "b"), [["a", "b"]]),
        ((("x",), ("y", "z")), [["x"], ["y", "z"]]),
    ],
)
def test_normalize_com_value(val: object, expected_rows: list) -> None:  # noqa: ANN001
    assert normalize_com_value(val) == expected_rows
