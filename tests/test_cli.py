"""CLI 데모·진입 동작."""

import io
from contextlib import redirect_stdout
from unittest.mock import patch

from bom_review.cli import cmd_demo, main


def test_cmd_demo_prints_matching_sample():
    buf = io.StringIO()
    with redirect_stdout(buf):
        assert cmd_demo() == 0
    out = buf.getvalue()
    assert "매칭 데모" in out
    assert "GHOST" in out
    assert "TP1" in out or "FID1" in out


def test_main_empty_argv_starts_gui():
    """인자 없음은 정식 GUI — 테스트에서는 tkinter 대신 cmd_gui만 스텁."""
    with patch("bom_review.cli.cmd_gui", return_value=0) as stub:
        assert main([]) == 0
        stub.assert_called_once()


def test_main_demo_subcommand():
    assert main(["demo"]) == 0
