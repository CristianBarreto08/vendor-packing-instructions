"""Generate the synthetic sample workspace used by the demo and the tests.

Everything here is invented: vendors, SKUs, product names, barcodes and
quantities. The layout mirrors the real inputs (column positions, header
quirks, stale template rows) so every edge case the tool handles is exercised:

- a pack whose Total is #N/A, and one with zero pieces per bag (both dropped)
- a product whose only pack is #N/A (dropped entirely; still on the PO)
- merchant-fulfilled packs with no marketplace SKU (red), one with a barcode
  found by internal SKU, one with no barcode at all (DO NOT PRINT LABEL)
- a parent whose declared total does not match its packs
- a parent missing from the sales report (no priority)
- case sizes that are confirmed (mm), unconfirmed legacy (cm), or blank
- a stale dimensions export left in Input/ (must be ignored)
- a template whose table sits below instruction rows, with stale example rows,
  cross-workbook formulas, a comment, and two sample images
- a PO export with a quantity mismatch and a SKU that was never packed

Run:  python examples/generate_sample_data.py
"""

from __future__ import annotations

import random
import shutil
import string
import struct
import zipfile
import zlib
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SEED = 2041
PO = 2041
VENDOR = "Northwind Manufacturing"
OTHER_VENDOR = "Pacific Components"
OUT = Path(__file__).resolve().parent / "sample_data"
FIXED_TIME = datetime(2026, 9, 10, 9, 0, 0)

# parent, label name, sub, [(channel, pcs per bag)]
PRODUCTS = [
    ("HX-1042", "Hex Key Set", "10pc Metric", [("FBA", 5), ("FBA", 10)]),
    ("CT-2210", "Cable Ties", "8in Black", [("FBA", 50), ("FBA", 100), ("FBA", 250)]),
    ("SB-3305", "Storage Bin", "Small", [("FBA", 2), ("FBA", 4)]),
    ("MT-0450", "Measuring Tape", "25ft", [("FBA", 1), ("FBA", 3)]),
    ("WG-7718", "Work Gloves", "Large", [("FBA", 3), ("FBA", 6)]),
    ("ZP-1180", "Zipper Pouch", "", [("FBA", 5), ("MF", 20)]),
    ("LB-6620", "LED Bulb", "A19 Warm White", [("FBA", 4), ("FBA", 8)]),
    ("SC-9014", "Spring Clamp", "2in", [("FBA", 4), ("FBA", 8), ("FBA", 12)]),
    ("BC-5521", "Bungee Cord", "24in", [("FBA", 6), ("MF", 2)]),
    ("KS-3307", "Knife Sharpener", "Two Stage", [("FBA", 1)]),
]
OTHER_PRODUCTS = [
    ("DG-4410", "Drawer Glide", "14in", [("FBA", 2)]),
    ("HB-2290", "Hinge Bracket", "Brass", [("FBA", 4), ("FBA", 10)]),
]

TOTAL_NA = {"SC-9014-P12", "KS-3307-P1"}
PACK_ZERO = {"LB-6620-P8"}
SHORT_CASE = {"CT-2210-P250"}           # last carton not full -> verify warning
NO_BARCODE = {"MF-BC-5521-P2"}
BARCODE_BY_INTERNAL = {"MF-ZP-1180-P20"}
BARCODE2_ONLY = {"FBA-HX-1042-P10"}
UNCONFIRMED_DIMS = {"FBA-SB-3305-P2", "FBA-SB-3305-P4"}
NO_DIMS = {"FBA-MT-0450-P3"}
NO_PRIORITY = {"MT-0450"}
DECLARED_TOTAL_OFF = {"WG-7718": -30}
PO_QTY_OFF = {"CT-2210": 500}

HEADER_FILL = PatternFill("solid", fgColor="FFFFFF00")
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

rng = random.Random(SEED)


def code(n: int, alphabet: str = string.ascii_uppercase + string.digits) -> str:
    return "".join(rng.choice(alphabet) for _ in range(n))


def child_sku(channel: str, parent: str, pack: int) -> str:
    return f"{channel}-{parent}-P{pack}"


def short(sku: str) -> str:
    """Key used by the edge-case sets: channel prefix dropped for FBA packs."""
    return sku[4:] if sku.startswith("FBA-") else sku


# ---------------------------------------------------------------------------
# Plan rows
# ---------------------------------------------------------------------------

def plan_rows(products):
    """[(row values...)] for a vendor tab plus metadata per child."""
    rows, children = [], []
    for parent, label, sub, packs in products:
        kids = []
        for channel, pack in packs:
            sku = child_sku(channel, parent, pack)
            bags_case = rng.choice([10, 20, 25, 40, 50])
            cases = rng.randint(2, 24)
            bags = cases * bags_case
            if short(sku) in SHORT_CASE:
                bags = cases * bags_case + bags_case // 4
            total = bags * pack
            kids.append(dict(sku=sku, parent=parent, pack=pack, bags_case=bags_case,
                             bags=bags, cases=round(bags / bags_case, 2), total=total,
                             label=label, sub=sub, channel=channel))
        valid = [k for k in kids if short(k["sku"]) not in TOTAL_NA | PACK_ZERO]
        declared = sum(k["total"] for k in valid) + DECLARED_TOTAL_OFF.get(parent, 0)
        sales = [rng.randint(50, 4000) for _ in range(6)]
        rows.append([parent, f"{label} {sub}".strip(), *sales, None, None, None, None,
                     None, None, declared or "#N/A"])
        for k in kids:
            s = short(k["sku"])
            total = "#N/A" if s in TOTAL_NA else k["total"]
            ordered = "#N/A" if s in TOTAL_NA else k["bags"]
            pack = 0 if s in PACK_ZERO else k["pack"]
            rows.append([k["sku"], f"{label} {sub} Pack of {k['pack']}".strip(),
                         *[rng.randint(5, 900) for _ in range(6)], parent,
                         k["pack"] * k["bags_case"], k["bags_case"], pack, ordered,
                         "#N/A" if s in TOTAL_NA else k["cases"], total])
            children.append(k)
    return rows, children


PLAN_HEADER = ["SKU", "Product", "30", "30T", "90", "90T", "OnOrder", "AvailableQty",
               "Parent SKU", "Case qty", "Case qty Packs", "Pack", "Ordered", "Cases", "Total",
               "Left", "Pcs Sold - 30 dys", "Pcs Sold-90 dys", "% Pcs Sold-30", "% Pcs Sold-90",
               "Avg % Pcs Sold", "Pcs Needed", "Packs Needed", "Cases Needed"]


def write_plan(path: Path, main_rows, other_rows, upload) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    ws.append(["Product Quantity Sold", "30 days and 90 days", "Synthetic sample data"])
    up = wb.create_sheet("UploadTemplate")
    up.append(["ProductID", "Vendor SKU", "Unit Price", "Qty Ordered", "Discount",
               "DiscountType", "Case qty"])
    for sku, qty in upload:
        up.append([sku, None, None, qty])
    for title, rows in ((VENDOR, main_rows), (OTHER_VENDOR, other_rows)):
        tab = wb.create_sheet(title)
        tab.append(PLAN_HEADER)
        for r in rows:
            tab.append(r)
        for cell in tab[1]:
            cell.font = Font(b=True)
    wb.create_sheet("30 Packs").append(["SKU", "Product Name", "Qty Sold"])
    _save(wb, path)


# ---------------------------------------------------------------------------
# Reference files
# ---------------------------------------------------------------------------

def write_marketplace_map(path: Path, children, marketplace) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Marketplace SKUs"
    ws.append(["Internal SKU", "Parent SKU", "Product", "Pack size", "Marketplace SKU",
               "Cost", "Notes"])
    seen = set()
    for k in children:
        if k["parent"] not in seen:
            seen.add(k["parent"])
            ws.append([k["parent"], None, f"{k['label']} {k['sub']}".strip()])
        ws.append([k["sku"], k["parent"], f"{k['label']} Pack of {k['pack']}", k["pack"],
                   marketplace.get(k["sku"]), round(rng.uniform(0.2, 9.0), 2)])
    _save(wb, path)


def write_barcodes(path: Path, children, marketplace, barcodes) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "All"
    ws.append(["Marketplace SKU", "Product Type", "Label Name", "Sub", "Notes", "Barcode",
               "Barcode2", "Internal SKU"])
    for k in children:
        sku = k["sku"]
        if sku in NO_BARCODE:
            continue
        if sku in BARCODE_BY_INTERNAL:
            ws.append([None, "Merchant", k["label"], k["sub"] or None, None, barcodes[sku],
                       None, sku])
            continue
        if sku not in marketplace:
            continue
        bc = barcodes[sku]
        ws.append([marketplace[sku], "FBA", k["label"], k["sub"] or None, None,
                   None if short(sku) in BARCODE2_ONLY else bc, bc, sku])
    _save(wb, path)


def write_sales_report(path: Path, products) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Main_Report"
    ws.append(["SKU", "Product Name", "Vendor", "30", "60", "90", "FBA Needed", "Stock",
               "On Order", "Order Qty", "Priority (Months)\nLast year",
               "Priority (Months)\n30 days", "Priority (Months)\n90 days", "New Notes"])
    for parent, label, sub, _ in products:
        if parent in NO_PRIORITY:
            continue
        ws.append([parent, f"{label} {sub}".strip(), VENDOR,
                   *[rng.randint(100, 9000) for _ in range(7)],
                   rng.randint(1, 6), rng.randint(1, 6), rng.randint(1, 6)])
    _save(wb, path)


def write_dimensions(path: Path, children, stale: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ProductID", "ProductName", "FBACASEQTY", "QtyPerCase", "SUPPLIERCASEQTY",
               "SUPPLIERCASEDIMENSIONS", "CaseLength", "CaseWidth", "CaseHeight",
               "PACKINGINSTRUCTIONS", "PACKINGINSTRUCTIONS2"])
    seen = set()
    for k in children:
        if k["parent"] not in seen:  # parents are requested too, usually blank
            seen.add(k["parent"])
            ws.append([k["parent"], k["label"], None, 0, None, None, 0, 0, 0, None, None])
        sku = k["sku"]
        dims = [rng.randrange(250, 600, 5) for _ in range(3)]
        if stale:
            text = f"{dims[0] // 10} x {dims[1] // 10} x {dims[2] // 10}"
        elif sku in NO_DIMS:
            text = ""
        elif sku in UNCONFIRMED_DIMS:
            text = f"{dims[0] // 10} x {dims[1] // 10} x {dims[2] // 10}"
        else:
            text = f"{dims[0]} x {dims[1]} x {dims[2]} Confirmed" + rng.choice(["", f" {PO - 7}"])
        note = rng.choice(["", "", "1pc per inner box", "Double-wall carton",
                           "N/A (not applicable)"])
        ws.append([sku, f"{k['label']} Pack of {k['pack']}", 1, k["pack"] * k["bags_case"],
                   None, text, round(dims[0] / 25.4, 2), round(dims[1] / 25.4, 2),
                   round(dims[2] / 25.4, 2), note, "N/A (not applicable)"])
    _save(wb, path)


def write_po_export(path: Path, upload, products) -> None:
    names = {p: f"{label} {sub}".strip() for p, label, sub, _ in products}
    wb = Workbook()
    ws = wb.active
    ws.title = "PurchaseOrderItems"
    ws.append(["PO #", "ProductID", "ProductName", "Vendor SKU", "Qty Ordered", "Qty Received"])
    for sku, qty in upload:
        ws.append([PO, sku, names[sku], None, qty + PO_QTY_OFF.get(sku, 0), 0])
    _save(wb, path)


# ---------------------------------------------------------------------------
# Vendor template
# ---------------------------------------------------------------------------

BOX_HEADER = ["ITEM SKU", "FBA SKU", "PRODUCT DESCRIPTION", "TOTAL PIECES",
              "Invoice Total per ITEM SKU", "Final total qty", "Balance", "PCS/BAG",
              "BAGS (BOXES)\n/CASE", "PCS/CASE", "# OF BAGS", "# OF CASES", "Box #", "PO #",
              "L (Inch)", "W (Inch)", "H (Inch)", "LENGTH (mm)", "WIDTH (mm)", "HEIGHT (mm)",
              "CBM", "N.W(KGS)", "G.W(KGS)", "MEAS(M3)", "Additional Instructions"]
LABELS_HEADER = ["PRIORITY", "ITEM SKU", "1. Pack Size", "Packs Required", "Total", "2. SKU ",
                 "3. Barcode", "Description"]
NUMBER_FORMATS = {4: "#,##0", 10: "#,##0", 11: "#,##0", 12: "#,##0.0", 21: "0.0000",
                  24: "0.000"}


def _header(ws, row: int, titles: list[str]) -> None:
    for c, title in enumerate(titles, 1):
        cell = ws.cell(row, c, title)
        cell.font = Font(b=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 32


def _styled_rows(ws, first: int, count: int, ncols: int) -> None:
    for r in range(first, first + count):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            # The first data row carries an odd font size from manual edits,
            # like the real template; the writer clones a later row.
            cell.font = Font(sz=9 if r == first else 11)
            if c in NUMBER_FORMATS:
                cell.number_format = NUMBER_FORMATS[c]


def write_template(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    for line in ["GENERAL PACKING INSTRUCTIONS FOR VENDORS",
                 "1. Pack each bag with the exact PCS/BAG shown in Box Information.",
                 "2. Print and apply the product label (see 'Product Label Sample').",
                 "3. Label every carton (see 'Box Label Sample').",
                 "4. Fill in N.W(KGS) per carton; gross weight and volume calculate themselves.",
                 "5. Bags over 5in opening need a suffocation warning."]:
        ws.append([line])
    ws["A1"].font = Font(b=True, sz=14)

    box = wb.create_sheet("Box Information")
    for i, line in enumerate(["BOX INFORMATION", "Fill in columns F, G and V only.",
                              "Do not mix SKUs in one carton.",
                              "Cartons must not exceed 22 kg.",
                              "Dimensions are outer carton size in millimeters."], 1):
        box.cell(i, 1, line)
    box["A1"].font = Font(b=True, sz=14)
    header = 8
    _header(box, header, BOX_HEADER)
    _styled_rows(box, header + 1, 40, len(BOX_HEADER))
    for i in range(14):  # stale example rows from another order
        r = header + 1 + i
        box.cell(r, 1, f"PC-{8800 + i}")
        box.cell(r, 2, f"=XLOOKUP(A{r},[1]All!$H:$H,[1]All!$A:$A)")
        box.cell(r, 3, f'=XLOOKUP(B{r},[1]All!$A:$A,[1]All!$C:$C)&" - "&'
                       f"XLOOKUP(B{r},[1]All!$A:$A,[1]All!$D:$D)")
        box.cell(r, 4, 120 * (i + 1))
    box.cell(header + 4, 4).comment = Comment("Internal: confirm qty with buyer", "ops")

    labels = wb.create_sheet("Product Labels")
    for i, line in enumerate(["PRODUCT LABELS", "Print one label per bag.",
                              "Rows marked DO NOT PRINT LABEL ship without a label."], 1):
        labels.cell(i, 1, line)
    _header(labels, 5, LABELS_HEADER)
    _styled_rows(labels, 6, 40, len(LABELS_HEADER))
    for i in range(12):
        r = 6 + i
        labels.cell(r, 1, f"=XLOOKUP(B{r},[2]Main_Report!$A:$A,[2]Main_Report!$AA:$AA)")
        labels.cell(r, 4, f"=XLOOKUP(F{r},'Box Information'!B:B,'Box Information'!K:K)")
        labels.cell(r, 6, f"MS-OLD{i:02d}")

    inv = wb.create_sheet("Packing List & Invoice Info")
    for i, line in enumerate(["PACKING LIST & INVOICE INFO", "Shipper: (vendor)",
                              "Consignee: (importer of record)",
                              "Use one invoice line per ITEM SKU."], 1):
        inv.cell(i, 1, line)
    inv.cell(6, 1, "ITEM SKU").font = Font(b=True)
    inv.cell(6, 2, "TOTAL").font = Font(b=True)
    for i in range(14):
        inv.append([f"PC-{8800 + i}", 120 * (i + 1)])
    inv.append(["TOTAL", "=SUM(B7:B20)"])

    for title in ("Box Label Sample", "Product Label Sample"):
        s = wb.create_sheet(title)
        s["A1"] = f"{title.upper()} - your label must look like this"
        s["A1"].font = Font(b=True, sz=14)
    check = wb.create_sheet("Check before Sending")
    for line in ["CHECK BEFORE SENDING (delete this tab before sending)",
                 "[ ] Priority filled in, no zeros", "[ ] No formulas pointing to other files",
                 "[ ] Rows without barcode say DO NOT PRINT LABEL", "[ ] PO # on every row"]:
        check.append([line])
    _save(wb, path)
    _embed_images(path, {wb.sheetnames.index("Box Label Sample"): _png(360, 180, 3),
                         wb.sheetnames.index("Product Label Sample"): _png(300, 120, 5)})


def _png(width: int, height: int, stripes: int) -> bytes:
    """A tiny mock label image (white card, black frame, barcode-like bars)."""
    def pixel(x: int, y: int) -> bytes:
        frame = x < 3 or y < 3 or x >= width - 3 or y >= height - 3
        bars = height * 0.55 < y < height * 0.9 and (x // stripes) % 3 == 0 and 20 < x < width - 20
        text = height * 0.15 < y < height * 0.3 and 20 < x < width * 0.7
        return b"\x00\x00\x00" if frame or bars else (b"\x60\x60\x60" if text else b"\xff\xff\xff")

    raw = b"".join(b"\x00" + b"".join(pixel(x, y) for x in range(width)) for y in range(height))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def _embed_images(path: Path, images: dict[int, bytes]) -> None:
    """Add one anchored picture per sheet by editing the package directly."""
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    types = parts["[Content_Types].xml"].decode()
    additions = '<Default Extension="png" ContentType="image/png"/>'
    for n, (sheet_index, data) in enumerate(sorted(images.items()), 1):
        sheet = f"xl/worksheets/sheet{sheet_index + 1}.xml"
        parts[f"xl/media/image{n}.png"] = data
        parts[f"xl/drawings/drawing{n}.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
            'spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'xmlns:r="{ns_r}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0'
            "</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            '<xdr:ext cx="3429000" cy="1714500"/><xdr:pic><xdr:nvPicPr>'
            f'<xdr:cNvPr id="{n + 1}" name="Picture {n}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch>'
            '</xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3429000" '
            'cy="1714500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            "</xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>").encode()
        parts[f"xl/drawings/_rels/drawing{n}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{ns_r}/image" Target="../media/image{n}.png"/>'
            "</Relationships>").encode()
        parts[f"xl/worksheets/_rels/sheet{sheet_index + 1}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{ns_r}/drawing" Target="../drawings/drawing{n}.xml"/>'
            "</Relationships>").encode()
        xml = parts[sheet].decode()
        if "xmlns:r=" not in xml:
            xml = xml.replace("<worksheet ", f'<worksheet xmlns:r="{ns_r}" ', 1)
        parts[sheet] = xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>').encode()
        additions += (f'<Override PartName="/xl/drawings/drawing{n}.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    parts["[Content_Types].xml"] = types.replace("</Types>", additions + "</Types>").encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def _save(wb: Workbook, path: Path) -> None:
    wb.properties.creator = "sample-data-generator"
    wb.properties.created = wb.properties.modified = FIXED_TIME
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------------------

def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    main_rows, children = plan_rows(PRODUCTS)
    other_rows, other_children = plan_rows(OTHER_PRODUCTS)
    all_children = children + other_children

    marketplace = {k["sku"]: f"MS-{code(4)}-{code(4)}" for k in all_children
                   if k["channel"] == "FBA"}
    barcodes = {k["sku"]: "Z0" + code(8) for k in all_children}

    upload = []
    for parent, *_ in PRODUCTS:
        packs = [k for k in children if k["parent"] == parent]
        shipped = [k for k in packs if short(k["sku"]) not in PACK_ZERO | TOTAL_NA]
        upload.append((parent, sum(k["total"] for k in shipped or packs)))

    write_plan(OUT / "Input" / f"Packing Plan 2026-09-10 - {VENDOR} PO {PO}.xlsx",
               main_rows, other_rows, upload)
    write_sales_report(OUT / "Input" / "Sales Report for Ordering 09-08-26.xlsx", PRODUCTS)
    write_dimensions(OUT / "Input" / "Catalog Dimensions job 480917 (old).xlsx", children, True)
    write_dimensions(OUT / "Inbox" / "Catalog Dimensions job 481302.xlsx", all_children, False)
    write_template(OUT / "Reference" / "Packing Instructions Template.xlsx")
    write_marketplace_map(OUT / "Reference" / "Marketplace SKU Map.xlsx", all_children, marketplace)
    write_barcodes(OUT / "Reference" / "Barcode List.xlsx", all_children, marketplace, barcodes)
    write_po_export(OUT / "PO Export" / f"PO {PO} export.xlsx", upload, PRODUCTS)
    print(f"Sample data written to {OUT}")


if __name__ == "__main__":
    main()
