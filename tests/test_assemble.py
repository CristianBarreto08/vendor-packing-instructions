from __future__ import annotations

from vendor_instructions.assemble import assemble, drop_reason
from vendor_instructions.inputs import PlanRow
from vendor_instructions.lookups import BarcodeCatalog, CaseDimensions


def row(n, sku, parent="", total=100.0, pack=10.0, case_packs=5.0) -> PlanRow:
    return PlanRow(row=n, sku=sku, product="", parent=parent, case_packs=case_packs,
                   pack=pack, ordered=None, cases=None, total=total)


def test_drop_reason_reports_the_root_cause_first():
    # Total #N/A usually comes with other blanks; the Total is what to fix.
    assert drop_reason(row(2, "A-P1", "A", total=None, pack=None)) == "TOTAL_NA"
    assert drop_reason(row(2, "A-P1", "A", pack=0)) == "PACK_ZERO"
    assert drop_reason(row(2, "A-P1", "A", case_packs=None)) == "CASE_PACKS_MISSING"
    assert drop_reason(row(2, "A-P1", "A")) is None


def test_double_lookup_label_fallbacks_and_inheritance():
    rows = [row(2, "HX-1", total=300), row(3, "FBA-HX-1-P5", "HX-1", total=200),
            row(4, "MF-HX-1-P2", "HX-1", total=100), row(5, "FBA-HX-1-P9", "HX-1", total=None)]
    marketplace = {"FBA-HX-1-P5": "MS-AAAA-0001"}
    barcodes = BarcodeCatalog(by_marketplace={"MS-AAAA-0001": ("Z0X", "Hex Key - 5pc")})
    dims = CaseDimensions(measures={"FBA-HX-1-P5": (400.0, 300.0, 200.0)})

    result = assemble(rows, marketplace, barcodes, dims, {"HX-1": 3})
    (group,) = result.groups
    fba, mf = group.packs
    assert (fba.fba_sku, fba.barcode, fba.dims) == ("MS-AAAA-0001", "Z0X", (400.0, 300.0, 200.0))
    # Merchant-fulfilled pack: no marketplace SKU, no barcode, description from sibling.
    assert mf.fba_sku is None and mf.barcode is None
    assert mf.description == "Hex Key - 5pc"
    assert result.inherited == [(4, "MF-HX-1-P2", "Hex Key - 5pc")]
    assert result.missing_barcode == [(4, "MF-HX-1-P2", None)]
    assert result.missing_dims == [(4, "MF-HX-1-P2")]
    assert result.dropped == [(5, "FBA-HX-1-P9", "TOTAL_NA")]
    assert group.priority == 3 and not result.unreconciled


def test_parent_total_reconciliation_priority_and_order():
    rows = [row(2, "B-1", total=999), row(3, "FBA-B-1-P2", "B-1", total=500),
            row(4, "A-1", total=100), row(5, "FBA-A-1-P2", "A-1", total=100),
            row(6, "FBA-ORPHAN-P1", "ORPHAN", total=50)]
    result = assemble(rows, {}, BarcodeCatalog(), None, {"A-1": 2})
    assert [g.parent for g in result.groups] == ["B-1", "A-1", "ORPHAN"]  # plan order kept
    assert result.unreconciled == [("B-1", 999, 500)]
    assert result.missing_priority == ["B-1", "ORPHAN"]
    assert result.total_pieces == 650
    assert result.missing_dims == []          # no export at all -> nothing flagged amber
