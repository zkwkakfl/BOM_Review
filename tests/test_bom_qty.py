"""BOM 수량 vs 좌표명 개수."""

from bom_review.bom_parse import normalize_designators_to_comma_space
from bom_review.bom_qty import bom_quantity_mismatch_findings


def test_qty_matches_split_count():
    assert not bom_quantity_mismatch_findings(["C1, C2"], ["2"])


def test_qty_matches_space_separated_tokens():
    """쉼표 없이 공백만 있어도 토큰 2개로 센다."""
    assert not bom_quantity_mismatch_findings(["C1 C2"], ["2"])


def test_qty_mismatch():
    f = bom_quantity_mismatch_findings(["C1"], ["2"])
    assert len(f) == 1
    assert f[0].code == "BOM_QTY_MISMATCH"


def test_normalize_designators_to_comma_space():
    assert normalize_designators_to_comma_space(None) is None
    assert normalize_designators_to_comma_space("") == ""
    assert normalize_designators_to_comma_space("R1 R2") == "R1, R2"
    assert normalize_designators_to_comma_space("R1,R2") == "R1, R2"
    assert normalize_designators_to_comma_space("R1; R2") == "R1, R2"
