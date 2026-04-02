"""CLI 데모·진입 동작."""

import io
from contextlib import redirect_stdout

from bom_review.cli import cmd_demo, main


def test_cmd_demo_prints_matching_sample():
    buf = io.StringIO()
    with redirect_stdout(buf):
        assert cmd_demo() == 0
    out = buf.getvalue()
    assert "매칭 데모" in out
    assert "GHOST" in out
    assert "TP1" in out or "FID1" in out


def test_main_empty_argv_runs_demo():
    assert main([]) == 0
