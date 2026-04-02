"""BOM 수량 vs 좌표명 개수."""

from bom_review.bom_qty import bom_quantity_mismatch_findings


def test_qty_matches_split_count():
    assert not bom_quantity_mismatch_findings(
        ["C1, C2"], ["2"], delimiter=", "
    )


def test_qty_mismatch():
    f = bom_quantity_mismatch_findings(["C1"], ["2"], delimiter=", ")
    assert len(f) == 1
    assert f[0].code == "BOM_QTY_MISMATCH"
