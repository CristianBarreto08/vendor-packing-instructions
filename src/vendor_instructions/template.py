"""Locate each table inside the vendor template -- never by a hard-coded row.

The template is owned by the operations team, not by this tool. Every time
someone adds or removes an instruction line *above* a table, the whole table
shifts. That already happened in production: the header moved down two rows,
the tool kept writing at the old fixed row, and the first product landed *on
top of* the column titles. No error was raised; the vendor simply received a
table without headers.

So the header row is searched for by its titles on every run. If it is not
found the run stops, instead of writing blind. ``self_check.py`` runs the same
search against the saved file to prove the header survived.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .columns import (
    BOX_SHEET,
    BOX_SIGNATURE,
    INVOICE_SHEET,
    INVOICE_TEMPLATE_SIGNATURE,
    LABELS_SHEET,
    LABELS_SIGNATURE,
    normalize,
)

SEARCH_DEPTH = 120  # instructions sit above the table; this leaves ample margin

# Which row to clone the formatting from, counted from the header. Not the
# first data row on purpose: in the template the first rows carry odd font
# sizes from manual edits, and the third is consistently formatted.
STYLE_OFFSET_BOX = 3
STYLE_OFFSET_LABELS = 2


class TemplateChanged(Exception):
    """The template no longer has the table where we know how to find it."""


@dataclass(frozen=True)
class TableLocation:
    header: int
    first_data: int
    style_row: int


def find_row(ws: Worksheet, signature: dict[int, str], depth: int = SEARCH_DEPTH) -> int | None:
    """First row where every column in ``signature`` has exactly that title."""
    for row in range(1, min(ws.max_row, depth) + 1):
        if all(normalize(ws.cell(row, col).value) == title for col, title in signature.items()):
            return row
    return None


def _locate(ws: Worksheet, signature: dict[int, str], offset: int, name: str) -> TableLocation:
    header = find_row(ws, signature)
    if header is None:
        raise TemplateChanged(_missing(name, signature))
    return TableLocation(header, header + 1, header + offset)


def box_table(ws: Worksheet) -> TableLocation:
    return _locate(ws, BOX_SIGNATURE, STYLE_OFFSET_BOX, BOX_SHEET)


def labels_table(ws: Worksheet) -> TableLocation:
    return _locate(ws, LABELS_SIGNATURE, STYLE_OFFSET_LABELS, LABELS_SHEET)


def invoice_example_row(ws: Worksheet) -> int:
    """Header of the example SKU list the template ships with.

    Everything from here down is wiped before writing. Overwriting without
    wiping was a real bug: when a vendor had fewer products than the example,
    the leftover example SKUs (from another order) were sent along.
    """
    row = find_row(ws, INVOICE_TEMPLATE_SIGNATURE)
    if row is None:
        raise TemplateChanged(_missing(INVOICE_SHEET, INVOICE_TEMPLATE_SIGNATURE))
    return row


def _missing(name: str, signature: dict[int, str]) -> str:
    cols = ", ".join(f"{get_column_letter(c)}={t.upper()}" for c, t in sorted(signature.items()))
    return (f'could not find the table header in tab "{name}" (looking for a row with '
            f"{cols}). If the column titles were renamed, update the signature in columns.py.")
