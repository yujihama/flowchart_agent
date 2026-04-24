"""Excel (.xlsx) renderer — real Shape objects + swimlane background.

Design
------
* Sheet ``フロー図``: swimlane bands as cell backgrounds + nodes/edges as
  real Excel Shape objects (drawing). The user can drag, resize and retext
  shapes; the sheet as a whole is overwritten on regeneration.
* Sheet ``注記``: free-form notes. Preserved across regeneration — if the
  target file already exists, its ``注記`` sheet is read and carried over.

Layout
------
All rank / collision / back-edge classification lives in
:mod:`jsox_flow.layout`. This module consumes a :class:`LayoutResult` and
translates its px coordinates into Excel's EMU / cell-anchor space.

Back-edge routing is delegated to Excel itself: connectors are wired via
``stCxn``/``endCxn`` to ``bentConnector3`` shapes and Excel auto-routes.
No custom "top channel" logic here.

Implementation
--------------
1. Build the base workbook with openpyxl (cells + formatting only).
2. Save to an in-memory buffer.
3. Post-process the xlsx ZIP: inject ``xl/drawings/drawing1.xml`` with
   ``<xdr:sp>`` for nodes and ``<xdr:cxnSp>`` for edges, and wire up the
   relationship (``sheet1.xml.rels``), content-type, and ``<drawing>``
   reference in ``sheet1.xml``.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union
from xml.sax.saxutils import escape as xml_escape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..layout import LayoutOptions, LayoutResult, compute_layout
from ..model import Flow, Node


# --- EMU conversion ---------------------------------------------------------

EMU_PER_PX = 9525
EMU_PER_PT = 12700
PX_PER_CHAR = 7.0
PX_PER_PT = 96 / 72


# --- sheet structure --------------------------------------------------------

SHEET_FLOW = "フロー図"
SHEET_NOTES = "注記"

TITLE_ROW = 1
WARNING_ROW = 2

# Horizontal: lane-header in col A, one cell per rank from col B.
H_FIRST_LANE_ROW = 3
H_LANE_HEADER_COL = 1
H_FIRST_STEP_COL = 2

# Vertical: lane-header in row 3, one cell per rank from row 4.
V_LANE_HEADER_ROW = 3
V_FIRST_LANE_COL = 1
V_FIRST_STEP_ROW = 4

ROW_TITLE_HEIGHT_PT = 28
ROW_WARNING_HEIGHT_PT = 32

H_HEADER_WIDTH_CHARS = 14
V_LEFT_CHANNEL_WIDTH_CHARS = 12  # unused cell for aesthetics in vertical


# --- styling ----------------------------------------------------------------

NODE_PRESET = {
    "start":      "flowChartTerminator",
    "end":        "flowChartTerminator",
    "task":       "flowChartProcess",
    "decision":   "flowChartDecision",
    "subprocess": "flowChartPredefinedProcess",
    "document":   "flowChartDocument",
}
NODE_FILL = {
    "start":      "FFF3CD",
    "end":        "FFF3CD",
    "task":       "FFFFFF",
    "decision":   "D1ECF1",
    "subprocess": "E2E3E5",
    "document":   "FEFEFE",
}
LANE_HEADER_FILL = "D4DFF0"
LANE_BODY_FILLS = ("F7F9FC", "EEF2F8")
TITLE_FILL = "2F5597"
WARNING_FILL = "FFF2CC"
EDGE_COLOR = "333333"
LABEL_COLOR = "C0392B"

LABEL_W_PX = 50
LABEL_H_PX = 18


# --- connection-point indices ----------------------------------------------
# per the OOXML preset definitions (<cxnLst> order) for all the flowChart*
# presets we emit: TOP, LEFT, BOTTOM, RIGHT.
SITE_TOP = 0
SITE_LEFT = 1
SITE_BOTTOM = 2
SITE_RIGHT = 3


# --- grid record -----------------------------------------------------------

@dataclass
class _Grid:
    """The Excel cell grid used as a visual backing for the swimlanes."""
    vertical: bool
    n_lanes: int
    n_ranks: int
    col_widths_chars: List[float]   # 1-indexed; index 0 unused
    row_heights_pt: List[float]
    col_edges_emu: List[int]        # cumulative EMU at each column edge
    row_edges_emu: List[int]

    def body_origin_emu(self) -> Tuple[int, int]:
        """(x, y) EMU offset from top-left of sheet to top-left of body."""
        if self.vertical:
            # body starts at col B (index 1 edge) and row 4 (index 3 edge)
            return (self.col_edges_emu[V_FIRST_LANE_COL - 1],
                    self.row_edges_emu[V_FIRST_STEP_ROW - 1])
        return (self.col_edges_emu[H_FIRST_STEP_COL - 1],
                self.row_edges_emu[H_FIRST_LANE_ROW - 1])

    def last_col(self) -> int:
        if self.vertical:
            return V_FIRST_LANE_COL + self.n_lanes - 1
        return H_FIRST_STEP_COL + self.n_ranks - 1

    def last_row(self) -> int:
        if self.vertical:
            return V_FIRST_STEP_ROW + self.n_ranks - 1
        return H_FIRST_LANE_ROW + self.n_lanes - 1


# --- public API -------------------------------------------------------------

def save_xlsx(
    flow: Flow,
    path: Union[str, Path],
    *,
    vertical: bool = False,
    layout: Optional[LayoutResult] = None,
    return_layout: bool = False,
) -> Union[Path, Tuple[Path, LayoutResult]]:
    """Render ``flow`` to an .xlsx file.

    Parameters
    ----------
    layout:
        Optional pre-computed layout. If omitted, one is computed with
        the default :class:`LayoutOptions` for the requested orientation.
    return_layout:
        If True, return a ``(path, layout)`` tuple so the caller (agent)
        can inspect metrics / warnings without recomputing.
    """
    target = Path(path)
    preserved_notes = _extract_notes(target)

    orientation = "vertical" if vertical else "horizontal"
    if layout is None:
        layout = compute_layout(flow, LayoutOptions(orientation=orientation))
    elif layout.orientation != orientation:
        raise ValueError(
            f"layout.orientation={layout.orientation!r} but "
            f"vertical={vertical!r} requires {orientation!r}"
        )

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_FLOW

    grid = _make_grid(flow, layout)
    _build_flow_sheet(ws, flow, layout, grid)

    notes_ws = wb.create_sheet(SHEET_NOTES)
    _build_notes_sheet(notes_ws, preserved_notes)

    buf = io.BytesIO()
    wb.save(buf)

    drawing_xml = _build_drawing_xml(flow, layout, grid)
    final_bytes = _inject_drawing(buf.getvalue(), drawing_xml)
    target.write_bytes(final_bytes)

    if return_layout:
        return target, layout
    return target


# --- grid construction ------------------------------------------------------

def _make_grid(flow: Flow, layout: LayoutResult) -> _Grid:
    vertical = layout.orientation == "vertical"
    n_lanes = len(layout.lane_order)
    n_ranks = layout.metrics.n_ranks or 1

    # Convert px grid sizes to Excel cell sizes.
    step_size_px = _step_size_px(layout)
    lane_size_px = _lane_size_px(layout)

    step_width_chars = step_size_px / PX_PER_CHAR
    lane_height_pt = lane_size_px / PX_PER_PT
    lane_width_chars = lane_size_px / PX_PER_CHAR
    step_height_pt = step_size_px / PX_PER_PT

    if vertical:
        col_widths_chars = [0.0, V_LEFT_CHANNEL_WIDTH_CHARS]
        col_widths_chars.extend([lane_width_chars] * n_lanes)
        row_heights_pt = [
            0.0,
            ROW_TITLE_HEIGHT_PT,
            ROW_WARNING_HEIGHT_PT,
            lane_size_px / PX_PER_PT / 2,   # lane header row
        ]
        row_heights_pt.extend([step_height_pt] * n_ranks)
    else:
        col_widths_chars = [0.0, H_HEADER_WIDTH_CHARS]
        col_widths_chars.extend([step_width_chars] * n_ranks)
        row_heights_pt = [
            0.0,
            ROW_TITLE_HEIGHT_PT,
            ROW_WARNING_HEIGHT_PT,
        ]
        row_heights_pt.extend([lane_height_pt] * n_lanes)

    col_edges = [0]
    for w in col_widths_chars[1:]:
        col_edges.append(col_edges[-1] + int(w * PX_PER_CHAR * EMU_PER_PX))
    row_edges = [0]
    for h in row_heights_pt[1:]:
        row_edges.append(row_edges[-1] + int(h * EMU_PER_PT))

    return _Grid(
        vertical=vertical,
        n_lanes=n_lanes,
        n_ranks=n_ranks,
        col_widths_chars=col_widths_chars,
        row_heights_pt=row_heights_pt,
        col_edges_emu=col_edges,
        row_edges_emu=row_edges,
    )


def _step_size_px(layout: LayoutResult) -> int:
    if layout.metrics.n_ranks <= 1:
        return layout.canvas_width if layout.orientation == "horizontal" else layout.canvas_height
    total = layout.canvas_width if layout.orientation == "horizontal" else layout.canvas_height
    return total // layout.metrics.n_ranks


def _lane_size_px(layout: LayoutResult) -> int:
    if layout.metrics.n_lanes <= 0:
        return 120
    total = layout.canvas_height if layout.orientation == "horizontal" else layout.canvas_width
    return total // layout.metrics.n_lanes


# --- notes sheet preservation ----------------------------------------------

def _extract_notes(path: Path) -> List[List]:
    if not path.exists():
        return []
    try:
        existing = load_workbook(path)
    except Exception:
        return []
    if SHEET_NOTES not in existing.sheetnames:
        return []
    ws = existing[SHEET_NOTES]
    rows: List[List] = []
    for row in ws.iter_rows(values_only=True):
        if any(cell not in (None, "") for cell in row):
            rows.append(list(row))
    return rows


def _build_notes_sheet(ws, preserved: List[List]) -> None:
    ws.column_dimensions["A"].width = 90
    ws.column_dimensions["B"].width = 20
    if not preserved:
        ws["A1"] = "このシートは手書き専用です。再生成時も内容は保持されます。"
        ws["A1"].font = Font(italic=True, color="808080")
        return
    for r, row in enumerate(preserved, start=1):
        for c, val in enumerate(row, start=1):
            ws.cell(r, c, val)


# --- flow sheet (cells) -----------------------------------------------------

def _build_flow_sheet(
    ws, flow: Flow, layout: LayoutResult, grid: _Grid,
) -> None:
    for c in range(1, len(grid.col_widths_chars)):
        ws.column_dimensions[get_column_letter(c)].width = grid.col_widths_chars[c]
    for r in range(1, len(grid.row_heights_pt)):
        ws.row_dimensions[r].height = grid.row_heights_pt[r]

    last_col = grid.last_col()
    last_row = grid.last_row()

    _write_title(ws, last_col)
    _write_warning(ws, last_col, layout.warnings)
    _write_lane_bands(ws, flow, layout, grid, last_row, last_col)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B4"


def _write_title(ws, last_col: int) -> None:
    ws.cell(TITLE_ROW, 1, "業務フロー図")
    ws.merge_cells(
        start_row=TITLE_ROW, start_column=1,
        end_row=TITLE_ROW, end_column=last_col,
    )
    c = ws.cell(TITLE_ROW, 1)
    c.font = Font(bold=True, size=16, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=TITLE_FILL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def _write_warning(ws, last_col: int, layout_warnings: List[str]) -> None:
    msg = (
        "⚠ このシートは自動生成です。図形はドラッグ/リサイズ/テキスト編集できますが、"
        "再生成時に「フロー図」シートは上書きされます。"
        "手書きメモは「注記」シートに残してください。"
    )
    if layout_warnings:
        msg += " [layout warnings: " + " / ".join(layout_warnings[:3]) + "]"
    ws.cell(WARNING_ROW, 1, msg)
    ws.merge_cells(
        start_row=WARNING_ROW, start_column=1,
        end_row=WARNING_ROW, end_column=last_col,
    )
    c = ws.cell(WARNING_ROW, 1)
    c.font = Font(italic=True, color="806000")
    c.fill = PatternFill("solid", fgColor=WARNING_FILL)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _write_lane_bands(
    ws, flow: Flow, layout: LayoutResult, grid: _Grid,
    last_row: int, last_col: int,
) -> None:
    thin = Side(style="thin", color="8A99B3")
    body_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    medium = Side(style="medium", color="333333")
    header_border = Border(left=medium, right=medium, top=medium, bottom=medium)

    lane_by_id = {l.id: l for l in flow.lanes}

    for i, lane_id in enumerate(layout.lane_order):
        lane = lane_by_id[lane_id]
        fill = PatternFill("solid", fgColor=LANE_BODY_FILLS[i % 2])

        if grid.vertical:
            lane_col = V_FIRST_LANE_COL + i
            # Lane header in row 3
            hdr = ws.cell(V_LANE_HEADER_ROW, lane_col, lane.name)
            hdr.fill = PatternFill("solid", fgColor=LANE_HEADER_FILL)
            hdr.font = Font(bold=True, size=12)
            hdr.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            hdr.border = header_border
            for r in range(V_FIRST_STEP_ROW, last_row + 1):
                cell = ws.cell(r, lane_col)
                cell.fill = fill
                cell.border = body_border
        else:
            lane_row = H_FIRST_LANE_ROW + i
            hdr = ws.cell(lane_row, H_LANE_HEADER_COL, lane.name)
            hdr.fill = PatternFill("solid", fgColor=LANE_HEADER_FILL)
            hdr.font = Font(bold=True, size=12)
            hdr.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            hdr.border = header_border
            for c in range(H_FIRST_STEP_COL, last_col + 1):
                cell = ws.cell(lane_row, c)
                cell.fill = fill
                cell.border = body_border


# --- EMU helpers -----------------------------------------------------------

def _emu_to_cell(emu: int, edges: List[int]) -> Tuple[int, int]:
    for i in range(len(edges) - 1):
        if edges[i] <= emu < edges[i + 1]:
            return i, emu - edges[i]
    return len(edges) - 2, emu - edges[-2]


def _two_cell_anchor(
    x_emu: int, y_emu: int, w_emu: int, h_emu: int, grid: _Grid,
) -> str:
    fc, fco = _emu_to_cell(x_emu, grid.col_edges_emu)
    fr, fro = _emu_to_cell(y_emu, grid.row_edges_emu)
    tc, tco = _emu_to_cell(x_emu + w_emu, grid.col_edges_emu)
    tr, tro = _emu_to_cell(y_emu + h_emu, grid.row_edges_emu)
    return (
        f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff>'
        f'<xdr:row>{fr}</xdr:row><xdr:rowOff>{fro}</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff>'
        f'<xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>'
    )


# --- drawing XML ------------------------------------------------------------

def _build_drawing_xml(
    flow: Flow, layout: LayoutResult, grid: _Grid,
) -> str:
    ox, oy = grid.body_origin_emu()

    node_shape_id: Dict[str, int] = {}
    next_id = 2
    for nid in layout.nodes:
        node_shape_id[nid] = next_id
        next_id += 1

    node_by_id = {n.id: n for n in flow.nodes}

    parts: List[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    for nid, nl in layout.nodes.items():
        node = node_by_id[nid]
        w_emu = nl.width * EMU_PER_PX
        h_emu = nl.height * EMU_PER_PX
        x_emu = ox + nl.x * EMU_PER_PX - w_emu // 2
        y_emu = oy + nl.y * EMU_PER_PX - h_emu // 2
        parts.append(_node_shape_xml(
            node_shape_id[nid], node, x_emu, y_emu, w_emu, h_emu, grid,
        ))

    label_w_emu = LABEL_W_PX * EMU_PER_PX
    label_h_emu = LABEL_H_PX * EMU_PER_PX

    for e in layout.edges:
        src_nl = layout.nodes[e.from_id]
        dst_nl = layout.nodes[e.to_id]
        src_cx = ox + src_nl.x * EMU_PER_PX
        src_cy = oy + src_nl.y * EMU_PER_PX
        dst_cx = ox + dst_nl.x * EMU_PER_PX
        dst_cy = oy + dst_nl.y * EMU_PER_PX
        w_emu = src_nl.width * EMU_PER_PX
        h_emu = src_nl.height * EMU_PER_PX

        # Let Excel auto-route: bentConnector3 wired via stCxn/endCxn.
        # Bounding box is just a routing hint; Excel re-computes on save.
        bx = min(src_cx, dst_cx) - w_emu // 2
        by = min(src_cy, dst_cy) - h_emu // 2
        bw = max(abs(dst_cx - src_cx), 1)
        bh = max(abs(dst_cy - src_cy), 1)
        flip_h = src_cx > dst_cx
        flip_v = src_cy > dst_cy

        parts.append(_bent_connector_xml(
            next_id,
            node_shape_id[e.from_id], SITE_RIGHT if not grid.vertical else SITE_BOTTOM,
            node_shape_id[e.to_id],   SITE_LEFT  if not grid.vertical else SITE_TOP,
            bx, by, bw, bh, flip_h, flip_v, grid,
        ))
        next_id += 1

        if e.condition:
            lx = (src_cx + dst_cx) // 2 - label_w_emu // 2
            ly = (src_cy + dst_cy) // 2 - label_h_emu // 2
            parts.append(_label_xml(
                next_id, lx, ly, label_w_emu, label_h_emu, e.condition, grid,
            ))
            next_id += 1

    parts.append("</xdr:wsDr>")
    return "\n".join(parts)


def _node_shape_xml(
    shape_id: int, node: Node,
    x_emu: int, y_emu: int, w_emu: int, h_emu: int,
    grid: _Grid,
) -> str:
    preset = NODE_PRESET.get(node.type, "flowChartProcess")
    fill = NODE_FILL.get(node.type, "FFFFFF")
    text = xml_escape(node.label)
    name = xml_escape(f"{node.id}_{node.label}")
    anchor = _two_cell_anchor(x_emu, y_emu, w_emu, h_emu, grid)
    return (
        '<xdr:twoCellAnchor editAs="oneCell">'
        f'{anchor}'
        '<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
        '<xdr:spPr>'
        f'<a:xfrm><a:off x="{x_emu}" y="{y_emu}"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>'
        '</xdr:spPr>'
        '<xdr:txBody>'
        '<a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        '<a:r><a:rPr lang="ja-JP" sz="1100" b="1">'
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        f'</a:rPr><a:t>{text}</a:t></a:r></a:p>'
        '</xdr:txBody>'
        '</xdr:sp>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
    )


def _bent_connector_xml(
    conn_id: int,
    src_shape_id: int, src_site: int,
    dst_shape_id: int, dst_site: int,
    bx: int, by: int, bw: int, bh: int,
    flip_h: bool, flip_v: bool,
    grid: _Grid,
) -> str:
    flips = []
    if flip_h:
        flips.append('flipH="1"')
    if flip_v:
        flips.append('flipV="1"')
    flip_attr = (" " + " ".join(flips)) if flips else ""

    anchor = _two_cell_anchor(bx, by, bw, bh, grid)

    return (
        '<xdr:twoCellAnchor editAs="oneCell">'
        f'{anchor}'
        '<xdr:cxnSp macro="">'
        f'<xdr:nvCxnSpPr><xdr:cNvPr id="{conn_id}" name="Edge_{conn_id}"/>'
        '<xdr:cNvCxnSpPr>'
        f'<a:stCxn id="{src_shape_id}" idx="{src_site}"/>'
        f'<a:endCxn id="{dst_shape_id}" idx="{dst_site}"/>'
        '</xdr:cNvCxnSpPr>'
        '</xdr:nvCxnSpPr>'
        '<xdr:spPr>'
        f'<a:xfrm{flip_attr}><a:off x="{bx}" y="{by}"/><a:ext cx="{bw}" cy="{bh}"/></a:xfrm>'
        '<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>'
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{EDGE_COLOR}"/></a:solidFill>'
        '<a:tailEnd type="triangle"/></a:ln>'
        '</xdr:spPr>'
        '</xdr:cxnSp>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
    )


def _label_xml(
    shape_id: int,
    x_emu: int, y_emu: int, w_emu: int, h_emu: int,
    text: str, grid: _Grid,
) -> str:
    t = xml_escape(text)
    anchor = _two_cell_anchor(x_emu, y_emu, w_emu, h_emu, grid)
    return (
        '<xdr:twoCellAnchor editAs="oneCell">'
        f'{anchor}'
        '<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="Label_{shape_id}"/>'
        '<xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
        '<xdr:spPr>'
        f'<a:xfrm><a:off x="{x_emu}" y="{y_emu}"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</xdr:spPr>'
        '<xdr:txBody>'
        '<a:bodyPr wrap="square" anchor="ctr" lIns="18000" tIns="9000" rIns="18000" bIns="9000"/>'
        '<a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="ja-JP" sz="1000" b="1">'
        f'<a:solidFill><a:srgbClr val="{LABEL_COLOR}"/></a:solidFill>'
        f'</a:rPr><a:t>{t}</a:t></a:r></a:p>'
        '</xdr:txBody>'
        '</xdr:sp>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
    )


# --- xlsx ZIP injection -----------------------------------------------------

DRAWING_CONTENT_TYPE = (
    '<Override PartName="/xl/drawings/drawing1.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
)
DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _inject_drawing(xlsx_bytes: bytes, drawing_xml: str) -> bytes:
    in_buf = io.BytesIO(xlsx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, "r") as zin:
        names = zin.namelist()
        existing_rels = (
            zin.read("xl/worksheets/_rels/sheet1.xml.rels")
            if "xl/worksheets/_rels/sheet1.xml.rels" in names
            else None
        )
        new_rels_bytes, r_id = _build_sheet_rels(existing_rels)

        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == "xl/worksheets/_rels/sheet1.xml.rels":
                    continue
                data = zin.read(name)
                if name == "xl/worksheets/sheet1.xml":
                    data = _add_drawing_ref(data, r_id)
                elif name == "[Content_Types].xml":
                    data = _ensure_drawing_content_type(data)
                zout.writestr(name, data)

            zout.writestr("xl/drawings/drawing1.xml", drawing_xml.encode("utf-8"))
            zout.writestr("xl/worksheets/_rels/sheet1.xml.rels", new_rels_bytes)

    return out_buf.getvalue()


def _build_sheet_rels(existing: bytes | None) -> Tuple[bytes, str]:
    if existing is None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{DRAWING_REL_TYPE}" Target="../drawings/drawing1.xml"/>'
            "</Relationships>"
        )
        return xml.encode("utf-8"), "rId1"

    text = existing.decode("utf-8")
    max_id = 0
    for m in re.finditer(r'Id="rId(\d+)"', text):
        max_id = max(max_id, int(m.group(1)))
    new_id = f"rId{max_id + 1}"
    new_rel = (
        f'<Relationship Id="{new_id}" Type="{DRAWING_REL_TYPE}" '
        'Target="../drawings/drawing1.xml"/>'
    )
    text = text.replace("</Relationships>", new_rel + "</Relationships>")
    return text.encode("utf-8"), new_id


def _add_drawing_ref(sheet_bytes: bytes, r_id: str) -> bytes:
    text = sheet_bytes.decode("utf-8")
    if "<drawing " in text:
        return sheet_bytes

    if "xmlns:r=" not in text[:500]:
        text = re.sub(
            r"<worksheet\b([^>]*)>",
            lambda m: f'<worksheet{m.group(1)} xmlns:r="{_REL_NS}">',
            text,
            count=1,
        )

    ref = f'<drawing r:id="{r_id}"/>'
    text = text.replace("</worksheet>", ref + "</worksheet>")
    return text.encode("utf-8")


def _ensure_drawing_content_type(ct_bytes: bytes) -> bytes:
    text = ct_bytes.decode("utf-8")
    if "drawings/drawing1.xml" in text:
        return ct_bytes
    text = text.replace("</Types>", DRAWING_CONTENT_TYPE + "</Types>")
    return text.encode("utf-8")
