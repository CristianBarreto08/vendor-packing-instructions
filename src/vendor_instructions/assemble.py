"""Build the parent/child groups that become the vendor workbook (the core logic).

Output is a list of product groups, each a parent with its pack sizes::

    HX-1042                                         <- parent row
      FBA-HX-1042-P5    MS-K4F7-Q2B9    600 pcs     <- child rows (packs)
      FBA-HX-1042-P10   MS-8DJ2-LL0P    400 pcs

Decisions made here
-------------------
1. **Double lookup.** The vendor prints the *marketplace* SKU on the bag, and
   that only exists per child: internal child SKU -> marketplace SKU ->
   barcode + description. A miss leaves the cell empty and red. The internal
   SKU is never used as a stand-in: the vendor would print it and the inventory
   would arrive mislabeled at the fulfillment center.
2. **Description inherited from siblings.** Packs without a marketplace SKU
   find no description, but they are the same product in a different pack
   size, so they take a sibling's. (This used to be a copy-paste by hand.)
3. **Rows with #N/A quantities are dropped.** Sending a line with an empty
   quantity is worse than not sending it; each drop is listed with its reason.
4. **The plan is reconciled against itself.** A parent's declared total must
   equal the sum of its packs; a mismatch is flagged *before* the file leaves.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

from .inputs import PlanRow
from .lookups import BarcodeCatalog, CaseDimensions, Dims


@dataclass
class PackLine:
    """One pack size: the row the vendor actually packs."""

    plan_row: int
    internal_sku: str
    total: float
    pcs_per_bag: float
    bags_per_case: float
    fba_sku: str | None = None
    barcode: str | None = None
    description: str | None = None
    dims: Dims | None = None
    note: str | None = None
    box_row: int | None = None  # set by the writer


@dataclass
class ProductGroup:
    parent: str
    order: int                      # position in the plan, keeps the planner's order
    declared_total: float | None    # Total on the parent's own row
    packs: list[PackLine] = field(default_factory=list)
    priority: int | str | None = None
    box_row: int | None = None      # set by the writer

    @property
    def packs_total(self) -> float:
        return sum(p.total for p in self.packs)

    @property
    def description(self) -> str | None:
        return next((p.description for p in self.packs if p.description), None)


@dataclass
class BuildResult:
    groups: list[ProductGroup] = field(default_factory=list)
    dropped: list[tuple[int, str, str]] = field(default_factory=list)       # row, sku, reason
    missing_fba: list[tuple[int, str]] = field(default_factory=list)
    missing_description: list[tuple[int, str]] = field(default_factory=list)
    missing_barcode: list[tuple[int, str, str | None]] = field(default_factory=list)
    missing_dims: list[tuple[int, str]] = field(default_factory=list)
    inherited: list[tuple[int, str, str]] = field(default_factory=list)
    unreconciled: list[tuple[str, float, float]] = field(default_factory=list)
    missing_priority: list[str] = field(default_factory=list)

    @property
    def pack_count(self) -> int:
        return sum(len(g.packs) for g in self.groups)

    @property
    def row_count(self) -> int:
        return len(self.groups) + self.pack_count

    @property
    def total_pieces(self) -> float:
        return sum(g.packs_total for g in self.groups)


# Reasons a child row cannot go to the vendor, most fundamental first: a row
# whose Total is #N/A usually has Cases in #N/A too, and "Total is #N/A" is the
# thing to actually fix. Adding a reason = add it here and in review.REASONS.
DROP_RULES: tuple[tuple[str, Callable[[PlanRow], bool]], ...] = (
    ("TOTAL_NA", lambda r: r.total is None),
    ("TOTAL_ZERO", lambda r: r.total is not None and r.total <= 0),
    ("PACK_MISSING", lambda r: r.pack is None),
    ("PACK_ZERO", lambda r: not r.pack),
    ("CASE_PACKS_MISSING", lambda r: r.case_packs is None),
    ("CASE_PACKS_ZERO", lambda r: not r.case_packs),
)


def drop_reason(row: PlanRow) -> str | None:
    """Code of the first rule that excludes this row, or None if it ships."""
    return next((code for code, applies in DROP_RULES if applies(row)), None)


def assemble(rows: list[PlanRow],
             marketplace: dict[str, str | None],
             barcodes: BarcodeCatalog,
             dims: CaseDimensions | None,
             priorities: dict[str, int | str] | None) -> BuildResult:
    """Turn plan rows into product groups plus everything worth reporting."""
    result = BuildResult()
    parents = {r.sku: r for r in rows if not r.is_child}
    groups: dict[str, ProductGroup] = {}

    for row in rows:
        if not row.is_child:
            continue
        reason = drop_reason(row)
        if reason:
            result.dropped.append((row.row, row.sku, reason))
            continue
        group = groups.get(row.parent)
        if group is None:
            parent_row = parents.get(row.parent)
            # Ordered by the parent's position; an orphan child group (no
            # parent row) sorts by its first child. Output keeps plan order.
            group = ProductGroup(
                parent=row.parent,
                order=parent_row.row if parent_row else row.row,
                declared_total=parent_row.total if parent_row else None,
            )
            groups[row.parent] = group
        group.packs.append(PackLine(
            plan_row=row.row, internal_sku=row.sku, total=row.total or 0.0,
            pcs_per_bag=row.pack or 0.0, bags_per_case=row.case_packs or 0.0))

    for group in groups.values():
        for pack in group.packs:
            _resolve(pack, marketplace, barcodes, dims, result)
        _inherit_description(group, result)

        if priorities is not None:
            group.priority = priorities.get(group.parent)
            if group.priority is None:
                result.missing_priority.append(group.parent)

        if (group.declared_total is not None
                and abs(group.declared_total - group.packs_total) > 0.5):
            result.unreconciled.append((group.parent, group.declared_total, group.packs_total))

    result.groups = sorted(groups.values(), key=lambda g: g.order)
    return result


def _resolve(pack: PackLine, marketplace: dict[str, str | None], barcodes: BarcodeCatalog,
             dims: CaseDimensions | None, result: BuildResult) -> None:
    pack.fba_sku = marketplace.get(pack.internal_sku)
    if not pack.fba_sku:
        result.missing_fba.append((pack.plan_row, pack.internal_sku))

    pack.barcode, pack.description = barcodes.lookup(pack.fba_sku, pack.internal_sku)
    if not pack.barcode:
        result.missing_barcode.append((pack.plan_row, pack.internal_sku, pack.fba_sku))

    if dims is not None:
        pack.dims = dims.measures.get(pack.internal_sku)
        pack.note = dims.notes.get(pack.internal_sku)
        if not pack.dims:
            result.missing_dims.append((pack.plan_row, pack.internal_sku))


def _inherit_description(group: ProductGroup, result: BuildResult) -> None:
    known = group.description
    for pack in group.packs:
        if pack.description:
            continue
        if known:
            pack.description = known
            result.inherited.append((pack.plan_row, pack.internal_sku, known))
        else:
            result.missing_description.append((pack.plan_row, pack.internal_sku))
