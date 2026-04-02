"""매칭 정책: 원본에만 있음은 오류가 아님."""

import pytest

from bom_review.matching import (
    FindingKind,
    bom_vs_source_findings,
    duplicate_reference_findings,
    iter_error_findings,
)


def test_bom_not_in_source_is_error():
    report = bom_vs_source_findings(
        bom_references=["R1", "R2", "GHOST"],
        source_references=["R1", "R2"],
    )
    assert len(report.bom_only_errors) == 1
    assert report.bom_only_errors[0].reference == "GHOST"
    assert report.bom_only_errors[0].kind == FindingKind.ERROR


def test_source_only_is_info_not_error():
    """원본에만 있고 BOM에 없음 → INFO, 오류 목록에 넣지 않음."""
    report = bom_vs_source_findings(
        bom_references=["R1"],
        source_references=["R1", "TP1", "FID1"],
    )
    refs_info = {f.reference for f in report.source_only_info}
    assert refs_info == {"FID1", "TP1"}
    assert all(f.kind == FindingKind.INFO for f in report.source_only_info)
    assert all(f.code == "SOURCE_ONLY" for f in report.source_only_info)

    errors = list(iter_error_findings(report))
    assert not any(f.reference in ("TP1", "FID1") for f in errors)
    assert report.has_errors is False


def test_iter_error_findings_excludes_source_only():
    report = bom_vs_source_findings(
        bom_references=["A", "BAD"],
        source_references=["A", "ONLY_SRC"],
    )
    err_refs = {f.reference for f in iter_error_findings(report)}
    assert err_refs == {"BAD"}


def test_duplicate_reference_errors():
    f = duplicate_reference_findings(["U1", "U1", "U2"])
    assert len(f) == 1
    assert f[0].reference == "U1"
    assert f[0].kind == FindingKind.ERROR
