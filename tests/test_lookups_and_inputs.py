from __future__ import annotations

from pathlib import Path

import pytest

from vendor_instructions import inputs
from vendor_instructions.export_source import ExportError, LocalFolderSource, coverage_gaps
from vendor_instructions.lookups import BarcodeCatalog, describe, parse_case_dimensions


@pytest.mark.parametrize("text, expected", [
    ("470 x 270 x 315 Confirmed 2041", ("confirmed", (470.0, 270.0, 315.0))),
    ("325X325X325 confirmed", ("confirmed", (325.0, 325.0, 325.0))),
    ("402,5 * 300 * 210 - Confirmed", ("confirmed", (402.5, 300.0, 210.0))),
    ("45 x 32 x 34", ("unconfirmed", None)),       # legacy, centimeters: never used
    ("", ("missing", None)),
    (None, ("missing", None)),
    ("N/A (not applicable)", ("missing", None)),
])
def test_only_confirmed_case_dimensions_are_used(text, expected):
    assert parse_case_dimensions(text) == expected


def test_description_has_no_dangling_dash():
    assert describe("Zipper Pouch", "") == "Zipper Pouch"
    assert describe("Cable Ties", "8in Black") == "Cable Ties - 8in Black"
    assert describe("", "") is None


def test_barcode_lookup_falls_back_to_internal_sku():
    catalog = BarcodeCatalog(by_marketplace={"MS-1": ("Z0AAA", "Item - A")},
                             by_internal={"MF-X-P2": ("Z0BBB", "Item - B")})
    assert catalog.lookup("MS-1", "FBA-X-P1") == ("Z0AAA", "Item - A")
    assert catalog.lookup(None, "MF-X-P2") == ("Z0BBB", "Item - B")
    assert catalog.lookup(None, "MF-X-P9") == (None, None)


@pytest.mark.parametrize("name, po", [
    ("Packing Plan 2026-09-10 - Vendor A PO 2041", 2041),
    ("plan vendor b po#156", 156),
    ("Plan - P.O. 77", 77),
    ("Packing Plan 2026-09-10 - Vendor A", None),   # dates are not PO numbers
])
def test_po_number_parsed_from_filename(name, po):
    assert inputs.po_from_filename(Path(name + ".xlsx")) == po


def test_files_are_classified_by_content(workspace):
    found = inputs.classify(workspace.input)
    assert found["plan"].name.startswith("Packing Plan")
    assert found["sales"].name.startswith("Sales Report")
    assert "(old)" in found["dimensions"].name          # stale copy is recognized...
    assert inputs.classify(workspace.inbox)["dimensions"].name.endswith("481302.xlsx")


def test_na_quantities_are_read_as_none_not_zero(workspace):
    plan = inputs.classify(workspace.input)["plan"]
    rows = {r.sku: r for r in inputs.read_plan_tab(plan, "Northwind Manufacturing")}
    assert rows["FBA-SC-9014-P12"].total is None
    assert rows["FBA-LB-6620-P8"].pack == 0
    assert rows["FBA-HX-1042-P5"].parent == "HX-1042"


def test_local_folder_source_and_coverage(workspace, tmp_path):
    path = LocalFolderSource(workspace.inbox).fetch_case_dimensions(["HX-1042"])
    assert coverage_gaps(path, ["HX-1042", "FBA-HX-1042-P5", "NOPE-1"]) == ["NOPE-1"]
    with pytest.raises(ExportError):
        LocalFolderSource(tmp_path / "empty").fetch_case_dimensions(["HX-1042"])
