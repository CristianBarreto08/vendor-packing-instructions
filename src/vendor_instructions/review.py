"""REVIEW.txt: everything that did not make it into the vendor file, and why.

Why a file and not just console output: the console gets closed. When someone
asks two weeks later "why didn't SKU X go to the vendor?", the answer needs to
be written down, with the date and the PO.

Every block states, in order: THE RULE (why a row lands here), WHAT TO DO
(where to fix it), and the list (plan row, SKU, the offending value).
"""

from __future__ import annotations

import textwrap
from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path

from .assemble import BuildResult
from .lookups import CaseDimensions

# reason code -> (title, rule, what to do). Keep in sync with assemble.DROP_RULES.
REASONS = {
    "TOTAL_NA": ("Total is #N/A",
                 "The Total column of the vendor tab is #N/A or blank, usually because the "
                 "Ordered formula did not resolve.",
                 "Fix it in the packing plan: check the row and where its Ordered comes from."),
    "TOTAL_ZERO": ("Total is zero",
                   "Nothing is being ordered for this pack size.",
                   "If intentional, nothing to do. Otherwise enter the quantity in the plan."),
    "PACK_MISSING": ("Pieces per bag missing",
                     "The Pack column is blank, so pieces per bag are unknown.",
                     "Fill it in in the packing plan."),
    "PACK_ZERO": ("Pieces per bag is zero",
                  "The Pack column is zero; # OF BAGS would divide by zero.",
                  "Correct it in the packing plan."),
    "CASE_PACKS_MISSING": ("Bags per case missing",
                           "The Case qty Packs column is blank.",
                           "Fill it in in the packing plan, or confirm with the vendor."),
    "CASE_PACKS_ZERO": ("Bags per case is zero",
                        "Case qty Packs is zero; # OF CASES would divide by zero.",
                        "Correct it in the packing plan."),
}
_UNKNOWN = ("Other reason", "No rule text for this code.", "Add it to review.REASONS.")
WIDTH = 78


def _wrap(label: str, text: str) -> list[str]:
    return textwrap.wrap(text, WIDTH, initial_indent="  " + label,
                         subsequent_indent=" " * (2 + len(label)))


class _Doc:
    def __init__(self) -> None:
        self.lines: list[str] = []
        self.open_items = 0

    def title(self, text: str) -> None:
        self.lines += ["=" * WIDTH, "  " + text, "=" * WIDTH, ""]

    def section(self, text: str) -> None:
        self.lines += ["", "-" * WIDTH, text, "-" * WIDTH, ""]

    def block(self, heading: str, rule: str, todo: str, items: Iterable[str]) -> None:
        self.lines += ["", "  " + heading, "  " + "." * len(heading),
                       *_wrap("THE RULE:   ", rule), *_wrap("WHAT TO DO: ", todo), ""]
        self.lines += [f"    {item}" for item in items]

    def missing(self, items: list, heading: str, rule: str, todo: str,
                fmt: Callable[[tuple], str]) -> bool:
        if not items:
            return False
        self.open_items += len(items)
        self.block(f"{heading}  -  {len(items)}", rule, todo, (fmt(x) for x in items))
        return True


def write_review(path: Path, result: BuildResult, dims: CaseDimensions | None,
                 context: dict[str, object], self_check: list[str]) -> int:
    """Write REVIEW.txt. Returns the number of open items."""
    doc = _Doc()
    doc.title("REVIEW - PACKING INSTRUCTIONS")
    doc.lines += [f"  {k:<14} {v}" for k, v in context.items()]
    doc.lines += [f"  {'generated':<14} {datetime.now():%Y-%m-%d %H:%M}", ""]

    doc.section("1. ROWS NOT INCLUDED IN THE VENDOR FILE")
    by_reason: dict[str, list[tuple[int, str]]] = {}
    for row, sku, code in result.dropped:
        by_reason.setdefault(code, []).append((row, sku))
    if not by_reason:
        doc.lines.append("  None. Every row of the vendor tab was included.")
    else:
        doc.lines += textwrap.wrap(
            f"{len(result.dropped)} plan rows were left out. The vendor will not see them: "
            "a line ships complete or not at all.", WIDTH, initial_indent="  ",
            subsequent_indent="  ")
        for code, rows in by_reason.items():
            title, rule, todo = REASONS.get(code, _UNKNOWN)
            doc.open_items += len(rows)
            doc.block(f"{title}  [{code}]  -  {len(rows)} rows", rule, todo,
                      (f"row {r:>4}   {sku}" for r, sku in rows))

    doc.section("2. ROWS INCLUDED, BUT WITH MISSING DATA")
    any_missing = False
    any_missing |= doc.missing(
        result.missing_fba, "No FBA SKU (RED in the workbook)",
        "The pack is not in the marketplace SKU map, or has no marketplace SKU there.",
        "The vendor cannot label this pack. Check the SKU map.",
        lambda x: f"row {x[0]:>4}   {x[1]}")
    any_missing |= doc.missing(
        result.missing_description, "No description (RED in the workbook)",
        "Not found in the barcode list and no sibling pack had one to inherit.",
        "Add the product to the barcode list.",
        lambda x: f"row {x[0]:>4}   {x[1]}")
    any_missing |= doc.missing(
        result.missing_barcode, 'No barcode -> "DO NOT PRINT LABEL" (RED)',
        "No barcode by marketplace SKU nor by internal SKU.",
        "Expected for merchant-fulfilled packs. Otherwise add the barcode to the list.",
        lambda x: f"row {x[0]:>4}   {x[1]:<26} {x[2] or '(no FBA SKU)'}")
    if dims is not None:
        any_missing |= doc.missing(
            result.missing_dims, "No case dimensions (AMBER in the workbook)",
            "The SKU has no CONFIRMED case size in the ERP, so L/W/H were left blank.",
            "Measure and confirm it in the ERP, or let the vendor choose the carton.",
            lambda x: f"row {x[0]:>4}   {x[1]}")
        any_missing |= doc.missing(
            dims.unconfirmed, 'Dimensions DISCARDED for not saying "Confirmed"',
            "A size is typed in the ERP but nobody confirmed it (legacy values are in cm; "
            "the CBM formula expects mm).",
            'If the size is right, add "Confirmed" (and the PO #) to it in the ERP.',
            lambda x: f'{x[0]:<26} "{x[1]}"')
    any_missing |= doc.missing(
        [(p,) for p in result.missing_priority], "No PRIORITY (AMBER in Product Labels)",
        "The parent SKU is not in the sales report, or its priority is blank.",
        "The pre-send checklist requires every priority to be filled in.",
        lambda x: x[0])
    if not any_missing:
        doc.lines.append("  None. Every row has all its data.")

    doc.section("3. THE PACKING PLAN DOES NOT RECONCILE")
    if not result.unreconciled:
        doc.lines.append("  Nothing. Every parent total equals the sum of its packs.")
    else:
        doc.open_items += len(result.unreconciled)
        doc.lines.append("  THE RULE:   a parent's Total must equal the sum of its packs' Totals.")
        doc.lines.append("  WHAT TO DO: fix it in the packing plan, NOT in the output.")
        doc.lines.append("")
        doc.lines.append(f"    {'PARENT':<20}{'PARENT SAYS':>14}{'PACKS SUM':>14}{'DIFF':>12}")
        for parent, declared, packs in result.unreconciled:
            doc.lines.append(f"    {parent:<20}{declared:>14,.0f}{packs:>14,.0f}"
                             f"{packs - declared:>+12,.0f}")

    doc.section("4. AUTOMATIC CHECK OF THE SAVED FILE")
    doc.lines += [f"  - {p}" for p in self_check] or ["  Passed."]
    doc.open_items += len(self_check)

    doc.section("BEFORE SENDING TO THE VENDOR")
    doc.lines += ["  [ ] Fix what needs fixing above and re-run",
                  '  [ ] Review the "Check before Sending" tab, then delete it',
                  "  [ ] Run `vendor-instructions verify` against the ERP PO export"]
    path.write_text("\n".join(doc.lines) + "\n", encoding="utf8")
    return doc.open_items
