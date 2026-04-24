"""Regenerate every example and run verification end-to-end.

For each YAML in ``examples/``:
    1. Render to ``.xlsx`` (and ``.svg`` / ``.mmd``).
    2. Convert the xlsx to PDF via LibreOffice.
    3. Rasterise each page to PNG.
    4. Verify every expected label made it into the rendered PDF.
    5. Print a compact JSON report per example.

Exit code is non-zero if any example fails verification, so this script
is safe to wire into CI.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

from jsox_flow import load_flow, verify_xlsx, VerifyError
from jsox_flow.render import save_xlsx, to_mermaid, to_svg


CASES = [
    ("sales_to_billing.yaml", True),   # (yaml_name, vertical?)
    ("purchase_process.yaml", False),
]


def main() -> int:
    here = Path(__file__).resolve().parent
    failures: list[str] = []

    for yaml_name, vertical in CASES:
        yaml_path = here / yaml_name
        stem = yaml_path.stem
        xlsx = here / f"{stem}.xlsx"
        svg = here / f"{stem}.svg"
        mmd = here / f"{stem}.mmd"

        flow = load_flow(yaml_path)
        save_xlsx(flow, xlsx, vertical=vertical)
        svg.write_text(to_svg(flow), encoding="utf-8")
        mmd.write_text(to_mermaid(flow), encoding="utf-8")

        try:
            result = verify_xlsx(xlsx, flow, render_png=True, png_dpi=120)
        except VerifyError as e:
            report = {"ok": False, "error": e.to_dict()}
            print(json.dumps({stem: report}, ensure_ascii=False, indent=2))
            failures.append(stem)
            continue

        report = {
            "ok": result.ok,
            "pages": result.page_count,
            "page_size_pt": list(result.page_size_pt),
            "missing_labels": result.missing_labels,
            "warnings": result.warnings,
            "pdf": result.pdf_path,
            "png": result.png_paths,
        }
        print(json.dumps({stem: report}, ensure_ascii=False, indent=2))
        if not result.ok:
            failures.append(stem)

    if failures:
        print(f"\nFAILED: {', '.join(failures)}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
