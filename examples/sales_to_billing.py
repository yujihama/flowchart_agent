"""Render sales_to_billing.yaml to SVG / Mermaid / Excel.

Run from project root:
    python examples/sales_to_billing.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

from jsox_flow import load_flow, validate
from jsox_flow.render import save_xlsx, to_mermaid, to_svg


def main() -> int:
    here = Path(__file__).resolve().parent
    yaml_path = here / "sales_to_billing.yaml"

    flow = load_flow(yaml_path)

    warnings = validate(flow)
    for w in warnings:
        print(f"WARNING: {w}", file=sys.stderr)

    svg_path = here / "sales_to_billing.svg"
    mmd_path = here / "sales_to_billing.mmd"
    xlsx_path = here / "sales_to_billing.xlsx"
    svg_path.write_text(to_svg(flow), encoding="utf-8")
    mmd_path.write_text(to_mermaid(flow), encoding="utf-8")
    save_xlsx(flow, xlsx_path, vertical=True)

    print(f"read  {yaml_path}")
    print(f"wrote {svg_path}")
    print(f"wrote {mmd_path}")
    print(f"wrote {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
