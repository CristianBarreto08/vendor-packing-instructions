"""The column map: the single source of truth for where every value lives.

Every other module refers to columns by *name*, never by letter. When a planner
moves a column in the packing plan, or someone edits the vendor template, this
is the only file that changes.

Parent / child hierarchy in the packing plan
--------------------------------------------
Each product occupies several rows of the vendor tab::

    parent row   SKU = HX-1042             Parent SKU = (blank)
    child rows   SKU = FBA-HX-1042-P5      Parent SKU = HX-1042
                 SKU = FBA-HX-1042-P10     Parent SKU = HX-1042

The parent is the product; the children are its pack sizes. The vendor never
sees the internal child SKU: it needs the *marketplace* SKU, which only exists
at child level. Hence the double lookup in ``assemble.py``.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# INPUT: vendor tab of the approved packing plan (0-based indexes)
# ---------------------------------------------------------------------------
PLAN = {
    "sku": 0,          # A  SKU (internal)
    "product": 1,      # B  Product
    "parent": 8,       # I  Parent SKU   (blank = this row IS the parent)
    "case_qty": 9,     # J  Case qty
    "case_packs": 10,  # K  Case qty Packs  -> BAGS (BOXES)/CASE
    "pack": 11,        # L  Pack            -> PCS/BAG
    "ordered": 12,     # M  Ordered (bags)
    "cases": 13,       # N  Cases
    "total": 14,       # O  Total (pieces)  -> TOTAL PIECES
}
PLAN_LAST_COL = 25

# A tab only counts as a vendor tab if its header matches this. Checked before
# reading anything, so a re-arranged tab aborts instead of being read shifted.
PLAN_SIGNATURE = {0: "SKU", 1: "Product", 8: "Parent SKU",
                  10: "Case qty Packs", 11: "Pack", 14: "Total"}

# Helper tabs of the packing plan that are never vendor tabs.
PLAN_FIXED_TABS = {"instructions", "uploadtemplate", "inventory", "kits",
                   "30 packs", "30 singles", "90 packs", "90 singles"}

# ---------------------------------------------------------------------------
# OUTPUT: "Box Information" (1-based column numbers)
# ---------------------------------------------------------------------------
BOX_SHEET = "Box Information"

# The header row is never hard-coded: it is located by these titles on every
# run (see template.py), because the template owner keeps adding instruction
# rows above the table.
BOX_SIGNATURE = {1: "item sku", 2: "fba sku", 3: "product description",
                 4: "total pieces"}

BOX = {
    "item_sku": 1,      # A  ITEM SKU (parent)
    "fba_sku": 2,       # B  FBA SKU (marketplace SKU)
    "description": 3,   # C  PRODUCT DESCRIPTION
    "total_pcs": 4,     # D  TOTAL PIECES
    "invoice_total": 5,  # E  Invoice total per item (formula)
    "final_qty": 6,     # F  filled in by the vendor
    "balance": 7,       # G  filled in by the vendor
    "pcs_bag": 8,       # H  PCS/BAG
    "bags_case": 9,     # I  BAGS (BOXES)/CASE
    "pcs_case": 10,     # J  PCS/CASE    (formula)
    "n_bags": 11,       # K  # OF BAGS   (formula)
    "n_cases": 12,      # L  # OF CASES  (formula)
    "box_no": 13,       # M  Box #
    "po": 14,           # N  PO #
    "length_mm": 18,    # R  LENGTH (mm)
    "width_mm": 19,     # S  WIDTH (mm)
    "height_mm": 20,    # T  HEIGHT (mm)
    "cbm": 21,          # U  CBM         (formula)
    "net_weight": 22,   # V  N.W(KGS)    filled in by the vendor
    "gross_weight": 23,  # W  G.W(KGS)   (formula over V)
    "measurement": 24,  # X  MEAS(M3)    (formula over U)
    "notes": 25,        # Y  Additional instructions
}
BOX_LAST_COL = 25

# ---------------------------------------------------------------------------
# OUTPUT: "Product Labels"
# ---------------------------------------------------------------------------
LABELS_SHEET = "Product Labels"
LABELS_SIGNATURE = {1: "priority", 2: "item sku", 3: "1. pack size", 6: "2. sku"}
LABELS = {
    "priority": 1,      # A  PRIORITY (parent row only)
    "item_sku": 2,      # B  ITEM SKU (parent row only)
    "pack_size": 3,     # C  1. Pack Size
    "packs_req": 4,     # D  Packs Required (formula -> Box Information)
    "total": 5,         # E  Total (formula, parent row only)
    "sku": 6,           # F  2. SKU (marketplace SKU)
    "barcode": 7,       # G  3. Barcode
    "description": 8,   # H  Description
}
LABELS_LAST_COL = 8

# ---------------------------------------------------------------------------
# OUTPUT: "Packing List & Invoice Info" (one line per product)
# ---------------------------------------------------------------------------
INVOICE_SHEET = "Packing List & Invoice Info"
# The template ships an example list from some other order under this header.
INVOICE_TEMPLATE_SIGNATURE = {1: "item sku", 2: "total"}
# The header this tool writes, so the self-check can find it again.
INVOICE_HEADER = ("ITEM SKU", "PRODUCT DESCRIPTION", "TOTAL PIECES")
INVOICE_SIGNATURE = {1: "item sku", 2: "product description", 3: "total pieces"}

# ---------------------------------------------------------------------------
# Reference: marketplace SKU map (internal child SKU -> marketplace SKU)
# ---------------------------------------------------------------------------
MARKETPLACE_SHEET = "Marketplace SKUs"
MARKETPLACE = {"internal_sku": 0, "parent": 1, "product": 2, "pack": 3,
               "marketplace_sku": 4}

# ---------------------------------------------------------------------------
# Reference: barcode list (marketplace SKU -> barcode + label name)
# ---------------------------------------------------------------------------
BARCODE_SHEET = "All"
BARCODE = {
    "marketplace_sku": 0,  # A
    "label": 2,            # C  Label Name
    "sub": 3,              # D  Sub (variant)
    "barcode": 5,          # F  Barcode
    "barcode2": 6,         # G  Barcode2 (backup copy of F)
    "internal_sku": 7,     # H  Internal SKU (fallback key)
}

# ---------------------------------------------------------------------------
# ERP catalog export with case dimensions
# ---------------------------------------------------------------------------
DIMENSIONS_COLUMNS = ("ProductID", "SUPPLIERCASEDIMENSIONS",
                      "PACKINGINSTRUCTIONS", "PACKINGINSTRUCTIONS2")

# ---------------------------------------------------------------------------
# Sales report: SKU -> priority (months of cover, 90-day average)
# ---------------------------------------------------------------------------
SALES_SHEET = "Main_Report"
SALES_SKU_HEADER = "sku"
SALES_PRIORITY_HEADER = "priority (months) 90 days"

# ---------------------------------------------------------------------------
# ERP purchase-order export (used by the `verify` command)
# ---------------------------------------------------------------------------
# Header aliases, normalized (lower-case, single spaces). ERPs rename columns
# between export templates, so each field accepts a few spellings.
PO_EXPORT_ALIASES = {
    "po": ("po #", "po number", "purchaseorderid", "po"),
    "sku": ("productid", "sku", "product id"),
    "qty": ("qty ordered", "qtyordered", "quantity ordered", "quantity"),
    "name": ("productname", "product name", "name"),
}


def normalize(value: object) -> str:
    """Cell text made comparable: no line breaks, single spaces, case-folded.

    Template titles contain embedded line breaks ("BAGS (BOXES)\\n/CASE") and
    trailing spaces ("2. SKU "), so comparing raw strings would fail for
    nothing.
    """
    if value is None:
        return ""
    return " ".join(str(value).split()).casefold()
