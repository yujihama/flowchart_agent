"""Visual / textual verification of generated xlsx files.

The renderer can produce structurally valid xlsx files that still look
wrong when opened — shapes clipped off-page, missing labels, text boxes
overwriting one another. Programmatic verification closes that gap by:

1. Converting the xlsx to PDF via LibreOffice headless (vector, same
   engine Excel-like spreadsheets render through).
2. Extracting text via ``pdftotext``.
3. Asserting every expected node label / edge condition appears.
4. Optionally rasterising to PNG (``pdftoppm``) for human review.

The result is a :class:`VerifyResult` — JSON-serialisable — so an
autonomous agent can read it and decide whether to regenerate with
different :class:`LayoutOptions`.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from .model import Flow


# --- error -----------------------------------------------------------------

class VerifyError(Exception):
    """Structured error; ``kind`` is a machine-readable constant."""

    KIND_TOOL_NOT_FOUND = "tool_not_found"
    KIND_CONVERT_FAILED = "convert_failed"
    KIND_EXTRACT_FAILED = "extract_failed"
    KIND_XLSX_MISSING = "xlsx_missing"

    def __init__(self, kind: str, message: str, details: Optional[Dict[str, Any]] = None):
        self.kind = kind
        self.message = message
        self.details = details or {}
        super().__init__(f"[{kind}] {message}")

    def to_dict(self) -> Dict[str, Any]:
        return {"kind": self.kind, "message": self.message, "details": self.details}


# --- result ----------------------------------------------------------------

@dataclass
class VerifyResult:
    xlsx_path: str
    pdf_path: Optional[str]
    page_count: int
    page_size_pt: Tuple[float, float]
    extracted_text: str
    expected_labels: List[str] = field(default_factory=list)
    missing_labels: List[str] = field(default_factory=list)
    png_paths: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    ok: bool = True

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "VerifyResult":
        return cls(
            xlsx_path=d["xlsx_path"],
            pdf_path=d.get("pdf_path"),
            page_count=d["page_count"],
            page_size_pt=tuple(d["page_size_pt"]),  # type: ignore[arg-type]
            extracted_text=d["extracted_text"],
            expected_labels=list(d.get("expected_labels", [])),
            missing_labels=list(d.get("missing_labels", [])),
            png_paths=list(d.get("png_paths", [])),
            warnings=list(d.get("warnings", [])),
            ok=d.get("ok", True),
        )


# --- public API ------------------------------------------------------------

def verify_xlsx(
    xlsx_path: Union[str, Path],
    flow: Flow,
    *,
    out_dir: Union[str, Path, None] = None,
    render_png: bool = False,
    png_dpi: int = 120,
    keep_pdf: bool = True,
    timeout: int = 60,
) -> VerifyResult:
    """Render ``xlsx_path`` to PDF via LibreOffice, extract text, and
    verify every expected label from ``flow`` appears.

    ``out_dir`` defaults to the xlsx's parent directory. The PDF is kept
    unless ``keep_pdf=False``. PNG rasters are written when
    ``render_png=True`` and paths returned in ``png_paths``.
    """
    xlsx = Path(xlsx_path)
    if not xlsx.exists():
        raise VerifyError(
            VerifyError.KIND_XLSX_MISSING,
            f"xlsx file not found: {xlsx}",
            {"path": str(xlsx)},
        )

    target_dir = Path(out_dir) if out_dir else xlsx.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = _convert_to_pdf(xlsx, target_dir, timeout=timeout)

    try:
        text = _extract_text(pdf_path, timeout=timeout)
        page_count, page_size = _probe_pdf(pdf_path)

        png_paths: List[str] = []
        if render_png:
            png_paths = _rasterise(pdf_path, target_dir, png_dpi, timeout=timeout)

        expected = _expected_labels(flow)
        # pdftotext preserves 2D layout, so multi-character labels inside
        # narrow shapes (diamonds especially) get split across lines.
        # Normalise by stripping all whitespace from both sides before
        # substring matching.
        normalised = _strip_ws(text)
        missing = [lbl for lbl in expected if _strip_ws(lbl) not in normalised]
        warnings: List[str] = []
        if missing:
            warnings.append(
                f"{len(missing)} expected labels not found in rendered PDF: "
                + ", ".join(missing[:5])
                + ("..." if len(missing) > 5 else "")
            )
        if page_count == 0:
            warnings.append("PDF has zero pages.")

        return VerifyResult(
            xlsx_path=str(xlsx),
            pdf_path=str(pdf_path) if keep_pdf else None,
            page_count=page_count,
            page_size_pt=page_size,
            extracted_text=text,
            expected_labels=expected,
            missing_labels=missing,
            png_paths=png_paths,
            warnings=warnings,
            ok=not missing and page_count > 0,
        )
    finally:
        if not keep_pdf and pdf_path.exists():
            pdf_path.unlink()


# --- external tools --------------------------------------------------------

def _find_tool(env_var: str, names: List[str]) -> str:
    exe = os.environ.get(env_var)
    if exe:
        return exe
    for name in names:
        p = shutil.which(name)
        if p:
            return p
    raise VerifyError(
        VerifyError.KIND_TOOL_NOT_FOUND,
        f"none of {names} found on PATH (set ${env_var} to override).",
        {"env_var": env_var, "candidates": names},
    )


def _convert_to_pdf(xlsx: Path, out_dir: Path, *, timeout: int) -> Path:
    soffice = _find_tool("JSOX_SOFFICE", ["soffice", "libreoffice"])
    with tempfile.TemporaryDirectory(prefix="jsox-lo-") as profile:
        cmd = [
            soffice,
            "--headless", "--norestore", "--nologo", "--nofirststartwizard",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(xlsx),
        ]
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=timeout, check=False,
            )
        except subprocess.TimeoutExpired as e:
            raise VerifyError(
                VerifyError.KIND_CONVERT_FAILED,
                f"LibreOffice timed out after {timeout}s",
                {"stderr": e.stderr or ""},
            )

    pdf = out_dir / (xlsx.stem + ".pdf")
    if proc.returncode != 0 or not pdf.exists():
        raise VerifyError(
            VerifyError.KIND_CONVERT_FAILED,
            f"LibreOffice failed to produce PDF (rc={proc.returncode}).",
            {
                "returncode": proc.returncode,
                "stderr": proc.stderr,
                "stdout": proc.stdout,
                "expected_pdf": str(pdf),
            },
        )
    return pdf


def _extract_text(pdf: Path, *, timeout: int) -> str:
    exe = _find_tool("JSOX_PDFTOTEXT", ["pdftotext"])
    # ``-raw`` extracts in reading order (left-to-right, top-to-bottom)
    # rather than reconstructing 2D layout, so labels that wrap inside a
    # shape stay on consecutive lines and survive whitespace-stripping.
    proc = subprocess.run(
        [exe, "-raw", str(pdf), "-"],
        capture_output=True, text=True, timeout=timeout, check=False,
    )
    if proc.returncode != 0:
        raise VerifyError(
            VerifyError.KIND_EXTRACT_FAILED,
            f"pdftotext failed (rc={proc.returncode}): {proc.stderr.strip()}",
            {"returncode": proc.returncode, "stderr": proc.stderr},
        )
    return proc.stdout


def _probe_pdf(pdf: Path) -> Tuple[int, Tuple[float, float]]:
    exe = shutil.which("pdfinfo")
    if not exe:
        # pdfinfo is optional — fall back to zero/unknown size
        return (_count_pages_fallback(pdf), (0.0, 0.0))
    proc = subprocess.run(
        [exe, str(pdf)], capture_output=True, text=True, timeout=30, check=False,
    )
    page_count = 0
    size = (0.0, 0.0)
    for line in proc.stdout.splitlines():
        if line.startswith("Pages:"):
            try:
                page_count = int(line.split(":", 1)[1].strip())
            except ValueError:
                pass
        elif line.startswith("Page size:"):
            # e.g. "Page size: 595.304 x 841.89 pts (A4)"
            try:
                spec = line.split(":", 1)[1].strip()
                w, _, rest = spec.partition(" x ")
                h = rest.split(" ")[0]
                size = (float(w), float(h))
            except Exception:
                pass
    return page_count, size


def _count_pages_fallback(pdf: Path) -> int:
    try:
        data = pdf.read_bytes()
    except OSError:
        return 0
    return data.count(b"/Type /Page\n") + data.count(b"/Type/Page\n") or 1


def _rasterise(pdf: Path, out_dir: Path, dpi: int, *, timeout: int) -> List[str]:
    exe = _find_tool("JSOX_PDFTOPPM", ["pdftoppm"])
    # Clean stale PNGs from previous runs so the result exactly matches
    # the current PDF's page count.
    for stale in out_dir.glob(f"{pdf.stem}-*.png"):
        try:
            stale.unlink()
        except OSError:
            pass
    prefix = out_dir / pdf.stem
    proc = subprocess.run(
        [exe, "-r", str(dpi), "-png", str(pdf), str(prefix)],
        capture_output=True, text=True, timeout=timeout, check=False,
    )
    if proc.returncode != 0:
        raise VerifyError(
            VerifyError.KIND_CONVERT_FAILED,
            f"pdftoppm failed (rc={proc.returncode}): {proc.stderr.strip()}",
            {"returncode": proc.returncode, "stderr": proc.stderr},
        )
    return sorted(str(p) for p in out_dir.glob(f"{pdf.stem}-*.png"))


# --- expectations ----------------------------------------------------------

def _strip_ws(s: str) -> str:
    return "".join(ch for ch in s if not ch.isspace())


def _expected_labels(flow: Flow) -> List[str]:
    labels: List[str] = []
    seen: set = set()
    for n in flow.nodes:
        if n.label and n.label not in seen:
            seen.add(n.label)
            labels.append(n.label)
    for e in flow.edges:
        if e.condition and e.condition not in seen:
            seen.add(e.condition)
            labels.append(e.condition)
    for l in flow.lanes:
        if l.name and l.name not in seen:
            seen.add(l.name)
            labels.append(l.name)
    return labels
