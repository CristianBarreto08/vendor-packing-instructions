"""The build pipeline, in order.

    1. find the inputs (by content)                 inputs.py
    2. pick the vendor tab, parse the PO number     inputs.py
    3. load the four lookup tables                  lookups.py, export_source.py
    4. assemble parent/child groups                 assemble.py
    5. write onto the template                      writer.py, media.py
    6. re-open and self-check the saved file        self_check.py
    7. write REVIEW.txt                             review.py

Nothing is fixed silently: anything unresolved is highlighted in the workbook
(red = blocking, amber = missing data) and listed in REVIEW.txt.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from . import inputs, lookups, review, self_check, writer
from .assemble import BuildResult, assemble
from .export_source import ExportError, ExportSource, LocalFolderSource
from .inputs import InputError

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class Workspace:
    root: Path

    @property
    def input(self) -> Path:
        return self.root / "Input"

    @property
    def inbox(self) -> Path:
        return self.root / "Inbox"

    @property
    def reference(self) -> Path:
        return self.root / "Reference"


@dataclass
class BuildOutcome:
    vendor: str
    po: int | None
    output: Path
    review: Path
    result: BuildResult
    problems: list[str]
    warnings: list[str] = field(default_factory=list)
    open_items: int = 0


def _reference_file(folder: Path, *, sheet: str | None = None,
                    name_part: str | None = None) -> Path:
    for path in inputs.xlsx_files(folder):
        if sheet and sheet in inputs.sheet_names(path):
            return path
        if name_part and name_part.lower() in path.name.lower():
            return path
    raise InputError(f"missing reference file in {folder} "
                     f"({'tab ' + repr(sheet) if sheet else repr(name_part)})")


def choose_vendor(tabs: list[str], requested: str | None) -> str:
    if not tabs:
        raise InputError("the packing plan has no vendor tab "
                         "(a tab with SKU / Product / Parent SKU / Pack / Total)")
    if requested:
        for tab in tabs:
            if tab.strip().lower() == requested.strip().lower():
                return tab
        raise InputError(f"no vendor tab {requested!r}; available: {', '.join(tabs)}")
    if len(tabs) > 1:
        raise InputError(f"the plan has several vendors; pass --vendor one of: {', '.join(tabs)}")
    return tabs[0]


def build(workspace: Workspace, *, vendor: str | None = None, po: int | None = None,
          require_priority: bool = True, output_dir: Path | None = None,
          dims_source: ExportSource | None = None) -> BuildOutcome:
    warnings: list[str] = []

    # 1. Inputs. Dimensions come only from Inbox/; a stale copy in Input/ is
    #    ignored and flagged, never used.
    found = inputs.classify(workspace.input)
    plan = found.get("plan")
    if not plan:
        raise InputError(f"no packing plan in {workspace.input}")
    sales = found.get("sales")
    if not sales and require_priority:
        raise InputError(f"no sales report in {workspace.input} (it supplies PRIORITY); "
                         "add it or run with --no-priority")
    if "dimensions" in found:
        warnings.append(f"ignored a dimensions export in Input/ ({found['dimensions'].name}); "
                        "only Inbox/ is used, so stale box sizes never leak in")
    reference = workspace.reference
    template_path = _reference_file(reference, name_part="template")
    marketplace_path = _reference_file(reference, sheet="Marketplace SKUs")
    barcode_path = _reference_file(reference, sheet="All")

    # 2. Vendor tab and PO.
    tab = choose_vendor(inputs.vendor_tabs(plan), vendor)
    po = po if po is not None else inputs.po_from_filename(plan)
    rows = inputs.read_plan_tab(plan, tab)

    # 3. Lookups.
    source = dims_source or LocalFolderSource(workspace.inbox)
    try:
        # Parents and children are both requested: some parents carry the
        # dimensions; surplus SKUs cost nothing, missing ones do.
        skus = list(dict.fromkeys(r.sku for r in rows))
        dims_path: Path | None = source.fetch_case_dimensions(skus)
    except ExportError as exc:
        warnings.append(f"no case dimensions: {exc}")
        dims_path = None
    dims = lookups.load_case_dimensions(dims_path) if dims_path else None
    priorities = lookups.load_priorities(sales) if sales else None

    # 4. Assemble.
    result = assemble(rows, lookups.load_marketplace_skus(marketplace_path),
                      lookups.load_barcodes(barcode_path), dims, priorities)
    log.info("%d products, %d packs", len(result.groups), result.pack_count)

    # 5-7. Write, self-check, review.
    out_dir = output_dir or workspace.root / "Output"
    stem = " - ".join([tab.strip(), "{kind}"] + ([f"PO {po}"] if po is not None else [])
                      + [date.today().isoformat()])
    output = out_dir / (stem.format(kind="Packing Instructions") + ".xlsx")
    review_path = out_dir / (stem.format(kind="REVIEW") + ".txt")
    warnings += writer.write_workbook(template_path, output, result, po, tab, dims is not None)
    problems = self_check.check_output(output, result, po, template_path)
    open_items = review.write_review(review_path, result, dims, {
        "vendor": tab, "PO": po if po is not None else "(blank)", "plan": plan.name,
        "dimensions": dims_path.name if dims_path else "(none)",
        "sales report": sales.name if sales else "(not used)", "output": output.name,
    }, problems)
    return BuildOutcome(tab, po, output, review_path, result, problems, warnings, open_items)
