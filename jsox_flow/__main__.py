"""CLI entry point: ``python -m jsox_flow <subcommand>``.

Designed for autonomous agents: every subcommand emits structured JSON
on stdout (``--json`` flag) and uses exit codes + stderr for errors,
so callers can parse without resorting to log scraping.

Subcommands
-----------
``layout``   flow.yaml → LayoutResult JSON
``render``   flow.yaml + path → write SVG / XLSX / Mermaid; print report JSON
``validate`` flow.yaml → validation errors/warnings JSON

Examples
--------
::

    python -m jsox_flow layout examples/purchase_process.yaml --json
    python -m jsox_flow render examples/purchase_process.yaml out.xlsx --format xlsx
    python -m jsox_flow validate examples/purchase_process.yaml --json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict

from . import load_flow
from .layout import LayoutError, LayoutOptions, compute_layout
from .render import save_xlsx, to_mermaid, to_svg
from .validate import ValidationError, validate
from .verify import VerifyError, verify_xlsx


def _print_json(obj: Dict[str, Any]) -> None:
    json.dump(obj, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")


def _print_error(kind: str, message: str, details: Dict[str, Any] | None = None) -> int:
    _print_json({"ok": False, "error": {"kind": kind, "message": message, "details": details or {}}})
    return 1


def _load(flow_path: str):
    try:
        return load_flow(flow_path)
    except FileNotFoundError as e:
        raise SystemExit(_print_error("file_not_found", str(e), {"path": flow_path}))
    except Exception as e:
        raise SystemExit(_print_error("invalid_yaml", str(e), {"path": flow_path}))


def _cmd_layout(args: argparse.Namespace) -> int:
    flow = _load(args.flow)
    opts = LayoutOptions(
        orientation=args.orientation,
        dagre_ranker=args.ranker,
        dagre_nodesep=args.nodesep,
        dagre_ranksep=args.ranksep,
    )
    try:
        result = compute_layout(flow, opts)
    except LayoutError as e:
        return _print_error(e.kind, e.message, e.details)
    _print_json({"ok": True, "layout": result.to_dict(), "options": opts.to_dict()})
    return 0


def _cmd_render(args: argparse.Namespace) -> int:
    flow = _load(args.flow)
    out = Path(args.out)
    fmt = args.format or _infer_format(out)
    if fmt not in {"xlsx", "svg", "mermaid", "mmd"}:
        return _print_error("bad_format", f"unknown format {fmt!r}")

    opts = LayoutOptions(orientation=args.orientation)
    try:
        if fmt == "xlsx":
            vertical = args.orientation == "vertical"
            path, layout = save_xlsx(flow, out, vertical=vertical, return_layout=True)
            report = {
                "format": "xlsx",
                "path": str(path),
                "orientation": args.orientation,
                "layout": layout.to_dict(),
            }
        elif fmt == "svg":
            layout = compute_layout(flow, opts)
            out.write_text(to_svg(flow, layout=layout), encoding="utf-8")
            report = {
                "format": "svg",
                "path": str(out),
                "orientation": args.orientation,
                "layout": layout.to_dict(),
            }
        else:
            out.write_text(to_mermaid(flow), encoding="utf-8")
            report = {"format": "mermaid", "path": str(out)}
    except LayoutError as e:
        return _print_error(e.kind, e.message, e.details)

    _print_json({"ok": True, "report": report})
    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    flow = _load(args.flow)
    try:
        warnings = validate(flow)
    except ValidationError as e:
        return _print_error("validation_failed", str(e))
    _print_json({"ok": True, "warnings": warnings})
    return 0


def _cmd_verify(args: argparse.Namespace) -> int:
    flow = _load(args.flow)
    try:
        result = verify_xlsx(
            args.xlsx,
            flow,
            out_dir=args.out_dir,
            render_png=args.png,
            png_dpi=args.dpi,
            keep_pdf=not args.no_pdf,
        )
    except VerifyError as e:
        return _print_error(e.kind, e.message, e.details)

    payload = result.to_dict()
    if not args.include_text:
        # keep the stdout report compact by default
        payload["extracted_text_len"] = len(payload.pop("extracted_text"))
    _print_json({"ok": result.ok, "report": payload})
    return 0 if result.ok else 2


def _infer_format(out: Path) -> str:
    ext = out.suffix.lower()
    return {
        ".xlsx": "xlsx",
        ".svg": "svg",
        ".mmd": "mermaid",
        ".mermaid": "mermaid",
    }.get(ext, "")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="jsox_flow", description="JSOX flowchart CLI.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_common(sp):
        sp.add_argument("flow", help="Path to the flow YAML file.")
        sp.add_argument(
            "--orientation", choices=["horizontal", "vertical"],
            default="horizontal",
        )

    lp = sub.add_parser("layout", help="Compute and print a LayoutResult as JSON.")
    add_common(lp)
    lp.add_argument("--ranker", default="network-simplex",
                    choices=["network-simplex", "tight-tree", "longest-path"])
    lp.add_argument("--nodesep", type=int, default=30)
    lp.add_argument("--ranksep", type=int, default=60)
    lp.set_defaults(func=_cmd_layout)

    rp = sub.add_parser("render", help="Render flow to file; report JSON on stdout.")
    add_common(rp)
    rp.add_argument("out", help="Output path.")
    rp.add_argument("--format", choices=["xlsx", "svg", "mermaid"], default=None,
                    help="Output format (default: infer from extension).")
    rp.set_defaults(func=_cmd_render)

    vp = sub.add_parser("validate", help="Run flow validation; print warnings JSON.")
    add_common(vp)
    vp.set_defaults(func=_cmd_validate)

    vy = sub.add_parser(
        "verify",
        help="Convert xlsx to PDF via LibreOffice and verify labels are present.",
    )
    vy.add_argument("flow", help="Path to the source flow YAML (for expected labels).")
    vy.add_argument("xlsx", help="Path to the xlsx file to verify.")
    vy.add_argument("--out-dir", default=None,
                    help="Directory to write the PDF/PNG into (default: xlsx's dir).")
    vy.add_argument("--png", action="store_true",
                    help="Also rasterise each PDF page to PNG.")
    vy.add_argument("--dpi", type=int, default=120,
                    help="PNG raster DPI (default: 120).")
    vy.add_argument("--no-pdf", action="store_true",
                    help="Delete the intermediate PDF after verification.")
    vy.add_argument("--include-text", action="store_true",
                    help="Include the full extracted PDF text in the JSON report.")
    vy.set_defaults(func=_cmd_verify)

    return p


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
