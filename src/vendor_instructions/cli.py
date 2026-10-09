"""Command-line entry point: ``vendor-instructions build | verify | fetch-dims | demo``."""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

from . import inputs, po_check
from .export_source import BrowserConfig, ExportError, PlaywrightExportSource
from .pipeline import BuildOutcome, Workspace, build, choose_vendor
from .template import TemplateChanged

log = logging.getLogger("vendor_instructions")


def _print_build(outcome: BuildOutcome) -> None:
    r = outcome.result
    print(f"\nVendor {outcome.vendor} | PO {outcome.po}")
    print(f"  {len(r.groups)} products, {r.pack_count} packs, {r.total_pieces:,.0f} pieces")
    print(f"  workbook  {outcome.output}")
    print(f"  review    {outcome.review}  ({outcome.open_items} open items)")
    summary = [
        ("dropped rows (not sent)", len(r.dropped)),
        ("no FBA SKU (red)", len(r.missing_fba)),
        ("DO NOT PRINT LABEL (red)", len(r.missing_barcode)),
        ("no confirmed dimensions (amber)", len(r.missing_dims)),
        ("no priority (amber)", len(r.missing_priority)),
        ("descriptions inherited from siblings", len(r.inherited)),
        ("parents that do not reconcile", len(r.unreconciled)),
    ]
    for label, n in summary:
        if n:
            print(f"  - {label}: {n}")
    for w in outcome.warnings:
        print(f"  WARNING: {w}")
    print("  self-check: " + ("passed" if not outcome.problems else "FAILED"))
    for p in outcome.problems:
        print(f"    - {p}")


def _print_verify(report: po_check.CheckReport, checked: Path, notes: Path) -> None:
    status = "PASSED" if not report.errors else "FAILED"
    print(f"\nPO {report.po_number} check: {status} "
          f"({len(report.errors)} errors, {len(report.warnings)} warnings)")
    for f in report.errors:
        print(f"  ERROR    {f.sku:<12} {f.check:<10} {f.message}")
    print(f"  checked copy  {checked}")
    print(f"  notes         {notes}")


def cmd_build(args: argparse.Namespace) -> int:
    outcome = build(Workspace(Path(args.workspace)), vendor=args.vendor, po=args.po,
                    require_priority=not args.no_priority,
                    output_dir=Path(args.output_dir) if args.output_dir else None)
    _print_build(outcome)
    return 0 if not outcome.problems else 2


def cmd_verify(args: argparse.Namespace) -> int:
    instructions = Path(args.instructions)
    out_dir = Path(args.output_dir) if args.output_dir else instructions.parent
    report, checked, notes = po_check.verify(instructions, Path(args.po_export), out_dir)
    _print_verify(report, checked, notes)
    return 0 if not report.errors else 1


def cmd_fetch_dims(args: argparse.Namespace) -> int:
    ws = Workspace(Path(args.workspace))
    plan = inputs.classify(ws.input).get("plan")
    if not plan:
        raise inputs.InputError(f"no packing plan in {ws.input}")
    tab = choose_vendor(inputs.vendor_tabs(plan), args.vendor)
    skus = list(dict.fromkeys(r.sku for r in inputs.read_plan_tab(plan, tab)))
    config = json.loads(Path(args.config).read_text(encoding="utf8"))
    config = {k: Path(v) if k.endswith("_dir") else v for k, v in config.items()}
    config["download_dir"] = ws.inbox   # downloads always land in Inbox/, never Input/
    source = PlaywrightExportSource(BrowserConfig(**config), args.headless)
    path = source.fetch_case_dimensions(skus)
    print(f"Downloaded {path} covering {len(skus)} SKUs")
    return 0


def cmd_demo(args: argparse.Namespace) -> int:
    sample = Path(args.sample_dir)
    out_dir = Path(args.output_dir)
    outcome = build(Workspace(sample), vendor="Northwind Manufacturing", output_dir=out_dir)
    _print_build(outcome)
    po_export = next((sample / "PO Export").glob("*.xlsx"))
    report, checked, notes = po_check.verify(outcome.output, po_export, out_dir)
    _print_verify(report, checked, notes)
    print("\nThe sample PO export contains deliberate discrepancies, so the check reports errors.")
    return 0 if not outcome.problems else 2


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="vendor-instructions",
                                description="Build and verify vendor packing instructions.")
    p.add_argument("-v", "--verbose", action="store_true")
    sub = p.add_subparsers(dest="command", required=True)

    b = sub.add_parser("build", help="generate the vendor workbook + REVIEW.txt")
    b.add_argument("--workspace", default=".", help="folder with Input/, Inbox/, Reference/")
    b.add_argument("--vendor", help="vendor tab (required when the plan has several)")
    b.add_argument("--po", type=int, help="PO number (default: parsed from the plan filename)")
    b.add_argument("--no-priority", action="store_true", help="run without the sales report")
    b.add_argument("--output-dir", help="default: <workspace>/Output")
    b.set_defaults(func=cmd_build)

    v = sub.add_parser("verify", help="check a finished workbook against the ERP PO export")
    v.add_argument("instructions")
    v.add_argument("--po-export", required=True)
    v.add_argument("--output-dir")
    v.set_defaults(func=cmd_verify)

    f = sub.add_parser("fetch-dims", help="download case dimensions via the browser (optional)")
    f.add_argument("--workspace", default=".")
    f.add_argument("--vendor")
    f.add_argument("--config", required=True, help="JSON with BrowserConfig fields")
    f.add_argument("--headless", action="store_true")
    f.set_defaults(func=cmd_fetch_dims)

    d = sub.add_parser("demo", help="run build + verify on the synthetic sample data")
    d.add_argument("--sample-dir", default="examples/sample_data")
    d.add_argument("--output-dir", default="output")
    d.set_defaults(func=cmd_demo)
    return p


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format="%(levelname)s %(name)s: %(message)s")
    try:
        return args.func(args)
    except (inputs.InputError, TemplateChanged, ExportError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
