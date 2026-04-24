"""Excel (.xlsx) renderer — real Shape objects + swimlane background.

Design (C: area-separation)
---------------------------
* Sheet ``フロー図``: swimlane bands as cell backgrounds + nodes/edges as
  real Excel Shape objects (drawing). The user can drag, resize and retext
  shapes; the sheet as a whole is overwritten on regeneration.
* Sheet ``注記``: free-form notes. Preserved across regeneration — if the
  target file already exists, its ``注記`` sheet is read and carried over.

Orientation
-----------
``vertical=False`` (default): lanes are horizontal bands, flow goes LEFT->RIGHT.
``vertical=True``: lanes are vertical columns, flow goes TOP->BOTTOM.

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
from typing import Dict, List, Tuple, Union
from xml.sax.saxutils import escape as xml_escape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .._layout import assign_columns, back_edges, resolve_cell_collisions
from ..model import Edge, Flow, Node


# --- EMU conversion ---------------------------------------------------------

EMU_PER_PX = 9525
EMU_PER_PT = 12700
PX_PER_CHAR = 7.0


# --- grid dimensions --------------------------------------------------------
# Horizontal (default)
H_HEADER_WIDTH_CHARS = 14
H_STEP_WIDTH_CHARS = 22
H_ROW_LANE_HEIGHT_PT = 90
H_ROW_TOP_CHANNEL_HEIGHT_PT = 28
# Vertical
V_LEFT_CHANNEL_WIDTH_CHARS = 10
V_LANE_WIDTH_CHARS = 24
V_ROW_STEP_HEIGHT_PT = 70
V_ROW_LANE_HEADER_HEIGHT_PT = 32
# Common
ROW_TITLE_HEIGHT_PT = 28
ROW_WARNING_HEIGHT_PT = 32

NODE_W_PX = 130
NODE_H_PX = 46
LABEL_W_PX = 50
LABEL_H_PX = 18


# --- sheet addressing -------------------------------------------------------

SHEET_FLOW = "フロー図"
SHEET_NOTES = "注記"

TITLE_ROW = 1
WARNING_ROW = 2

# Horizontal body starts at row 4 and col B. Channel is row 3.
H_TOP_CHANNEL_ROW = 3
H_FIRST_LANE_ROW = 4
H_LANE_HEADER_COL = 1
H_FIRST_STEP_COL = 2

# Vertical body starts at row 4 and col B. Lane header is row 3, left channel is col A.
V_LANE_HEADER_ROW = 3
V_FIRST_STEP_ROW = 4
V_LEFT_CHANNEL_COL = 1
V_FIRST_LANE_COL = 2


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


# --- layout record ----------------------------------------------------------

@dataclass
class _Layout:
    vertical: bool
    n_lanes: int
    n_steps: int
    col_widths_chars: List[float]   # widths per col, 1-indexed (index 0 unused)
    row_heights_pt: List[float]
    col_edges_emu: List[int]
    row_edges_emu: List[int]

    def node_cell(self, lane_idx: int, step_idx: int) -> Tuple[int, int]:
        """Return (excel_row, excel_col), both 1-indexed."""
        if self.vertical:
            return (V_FIRST_STEP_ROW + step_idx, V_FIRST_LANE_COL + lane_idx)
        return (H_FIRST_LANE_ROW + lane_idx, H_FIRST_STEP_COL + step_idx)

    def lane_header_cell(self, lane_idx: int) -> Tuple[int, int]:
        if self.vertical:
            return (V_LANE_HEADER_ROW, V_FIRST_LANE_COL + lane_idx)
        return (H_FIRST_LANE_ROW + lane_idx, H_LANE_HEADER_COL)

    def last_col(self) -> int:
        if self.vertical:
            return V_FIRST_LANE_COL + self.n_lanes - 1
        return H_FIRST_STEP_COL + self.n_steps - 1

    def last_row(self) -> int:
        if self.vertical:
            return V_FIRST_STEP_ROW + self.n_steps - 1
        return H_FIRST_LANE_ROW + self.n_lanes - 1

    def back_channel_pos_emu(self) -> int:
        """Center of the back-edge channel along the flow axis (EMU)."""
        if self.vertical:
            # back edges travel along col A (LEFT channel) — return center x
            ci = V_LEFT_CHANNEL_COL - 1
            return (self.col_edges_emu[ci] + self.col_edges_emu[ci + 1]) // 2
        # horizontal: back edges travel along row 3 (TOP channel) — return center y
        ri = H_TOP_CHANNEL_ROW - 1
        return (self.row_edges_emu[ri] + self.row_edges_emu[ri + 1]) // 2


def _make_layout(flow: Flow, max_step_idx: int, vertical: bool) -> _Layout:
    n_lanes = len(flow.lanes)
    n_steps = max_step_idx + 1

    if vertical:
        col_widths_chars = [0.0, V_LEFT_CHANNEL_WIDTH_CHARS]
        col_widths_chars.extend([V_LANE_WIDTH_CHARS] * n_lanes)
        row_heights_pt = [
            0.0,
            ROW_TITLE_HEIGHT_PT,
            ROW_WARNING_HEIGHT_PT,
            V_ROW_LANE_HEADER_HEIGHT_PT,
        ]
        row_heights_pt.extend([V_ROW_STEP_HEIGHT_PT] * n_steps)
    else:
        col_widths_chars = [0.0, H_HEADER_WIDTH_CHARS]
        col_widths_chars.extend([H_STEP_WIDTH_CHARS] * n_steps)
        row_heights_pt = [
            0.0,
            ROW_TITLE_HEIGHT_PT,
            ROW_WARNING_HEIGHT_PT,
            H_ROW_TOP_CHANNEL_HEIGHT_PT,
        ]
        row_heights_pt.extend([H_ROW_LANE_HEIGHT_PT] * n_lanes)

    col_edges = [0]
    for w in col_widths_chars[1:]:
        col_edges.append(col_edges[-1] + int(w * PX_PER_CHAR * EMU_PER_PX))
    row_edges = [0]
    for h in row_heights_pt[1:]:
        row_edges.append(row_edges[-1] + int(h * EMU_PER_PT))

    return _Layout(
        vertical=vertical,
        n_lanes=n_lanes,
        n_steps=n_steps,
        col_widths_chars=col_widths_chars,
        row_heights_pt=row_heights_pt,
        col_edges_emu=col_edges,
        row_edges_emu=row_edges,
    )


# --- public API -------------------------------------------------------------

def save_xlsx(flow: Flow, path: Union[str, Path], *, vertical: bool = False) -> Path:
    target = Path(path)
    preserved_notes = _extract_notes(target)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_FLOW

    cols = assign_columns(flow)
    cols = resolve_cell_collisions(flow, cols)
    back = back_edges(flow, cols)
    max_step = max(cols.values()) if cols else 0
    layout = _make_layout(flow, max_step, vertical)

    _build_flow_sheet(ws, flow, layout)

    notes_ws = wb.create_sheet(SHEET_NOTES)
    _build_notes_sheet(notes_ws, preserved_notes)

    buf = io.BytesIO()
    wb.save(buf)

    drawing_xml = _build_drawing_xml(flow, cols, back, layout)
    final_bytes = _inject_drawing(buf.getvalue(), drawing_xml)
    target.write_bytes(final_bytes)
    return target


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

def _build_flow_sheet(ws, flow: Flow, layout: _Layout) -> None:
    for c in range(1, len(layout.col_widths_chars)):
        ws.column_dimensions[get_column_letter(c)].width = layout.col_widths_chars[c]
    for r in range(1, len(layout.row_heights_pt)):
        ws.row_dimensions[r].height = layout.row_heights_pt[r]

    last_col = layout.last_col()
    last_row = layout.last_row()

    _write_title(ws, last_col)
    _write_warning(ws, last_col)
    _write_lane_bands(ws, flow, layout, last_row, last_col)

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


def _write_warning(ws, last_col: int) -> None:
    msg = (
        "⚠ このシートは自動生成です。図形はドラッグ/リサイズ/テキスト編集できますが、"
        "再生成時に「フロー図」シートは上書きされます。"
        "手書きメモは「注記」シートに残してください。"
    )
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
    ws, flow: Flow, layout: _Layout, last_row: int, last_col: int
) -> None:
    thin = Side(style="thin", color="8A99B3")
    body_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    medium = Side(style="medium", color="333333")
    header_border = Border(left=medium, right=medium, top=medium, bottom=medium)

    for i, lane in enumerate(flow.lanes):
        fill = PatternFill("solid", fgColor=LANE_BODY_FILLS[i % 2])

        if layout.vertical:
            lane_col = V_FIRST_LANE_COL + i
            for r in range(V_FIRST_STEP_ROW, last_row + 1):
                cell = ws.cell(r, lane_col)
                cell.fill = fill
                cell.border = body_border
        else:
            lane_row = H_FIRST_LANE_ROW + i
            for c in range(H_FIRST_STEP_COL, last_col + 1):
                cell = ws.cell(lane_row, c)
                cell.fill = fill
                cell.border = body_border

        hdr_r, hdr_c = layout.lane_header_cell(i)
        header = ws.cell(hdr_r, hdr_c, lane.name)
        header.fill = PatternFill("solid", fgColor=LANE_HEADER_FILL)
        header.font = Font(bold=True, size=12)
        header.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        header.border = header_border


# --- cell / EMU helpers -----------------------------------------------------

def _cell_center(excel_row: int, excel_col: int, layout: _Layout) -> Tuple[int, int]:
    ci = excel_col - 1
    ri = excel_row - 1
    x = (layout.col_edges_emu[ci] + layout.col_edges_emu[ci + 1]) // 2
    y = (layout.row_edges_emu[ri] + layout.row_edges_emu[ri + 1]) // 2
    return x, y


def _emu_to_cell(emu: int, edges: List[int]) -> Tuple[int, int]:
    for i in range(len(edges) - 1):
        if edges[i] <= emu < edges[i + 1]:
            return i, emu - edges[i]
    return len(edges) - 2, emu - edges[-2]


def _two_cell_anchor(
    x_emu: int, y_emu: int, w_emu: int, h_emu: int, layout: _Layout,
) -> str:
    fc, fco = _emu_to_cell(x_emu, layout.col_edges_emu)
    fr, fro = _emu_to_cell(y_emu, layout.row_edges_emu)
    tc, tco = _emu_to_cell(x_emu + w_emu, layout.col_edges_emu)
    tr, tro = _emu_to_cell(y_emu + h_emu, layout.row_edges_emu)
    return (
        f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco}</xdr:colOff>'
        f'<xdr:row>{fr}</xdr:row><xdr:rowOff>{fro}</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco}</xdr:colOff>'
        f'<xdr:row>{tr}</xdr:row><xdr:rowOff>{tro}</xdr:rowOff></xdr:to>'
    )


# --- drawing XML ------------------------------------------------------------
# Connection-point indices for the flowChart* preset geometries we use.
# Per the OOXML preset definitions (<cxnLst> order), flowChartProcess /
# flowChartDecision / flowChartTerminator / flowChartDocument all expose
# their four sites in the order: TOP, LEFT, BOTTOM, RIGHT.
SITE_TOP = 0
SITE_LEFT = 1
SITE_BOTTOM = 2
SITE_RIGHT = 3


def _build_drawing_xml(
    flow: Flow, cols: Dict[str, int], back: set, layout: _Layout,
) -> str:
    lane_idx = {lane.id: i for i, lane in enumerate(flow.lanes)}

    # Each outgoing edge gets an index among its siblings so that multiple
    # edges leaving the same source node can be placed on different sides.
    outgoing: Dict[str, List[Edge]] = {}
    for e in flow.edges:
        outgoing.setdefault(e.from_id, []).append(e)

    node_center: Dict[str, Tuple[int, int]] = {}
    for n in flow.nodes:
        erow, ecol = layout.node_cell(lane_idx[n.lane_id], cols[n.id])
        node_center[n.id] = _cell_center(erow, ecol, layout)

    node_w_emu = NODE_W_PX * EMU_PER_PX
    node_h_emu = NODE_H_PX * EMU_PER_PX
    label_w_emu = LABEL_W_PX * EMU_PER_PX
    label_h_emu = LABEL_H_PX * EMU_PER_PX

    parts: List[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    # Assign Shape IDs for every node up front so connectors can reference them.
    node_shape_id: Dict[str, int] = {}
    next_id = 2
    for n in flow.nodes:
        node_shape_id[n.id] = next_id
        next_id += 1

    get_node = {n.id: n for n in flow.nodes}

    for n in flow.nodes:
        cx, cy = node_center[n.id]
        x = cx - node_w_emu // 2
        y = cy - node_h_emu // 2
        parts.append(_node_shape_xml(
            node_shape_id[n.id], n, x, y, node_w_emu, node_h_emu, layout
        ))

    back_channel = layout.back_channel_pos_emu()

    for e in flow.edges:
        src = get_node[e.from_id]
        dst = get_node[e.to_id]
        is_back = (e.from_id, e.to_id) in back
        siblings = outgoing[e.from_id]
        edge_idx = siblings.index(e)
        n_siblings = len(siblings)

        if is_back:
            # Back edges: manual 3-segment U-route. bentConnector3 can't
            # produce a clean U when both endpoints are on the same side.
            src_site, _ = _pick_sites(
                src, dst, is_back, layout.vertical,
                lane_idx, edge_idx, n_siblings,
            )
            sx, sy = node_center[src.id]
            dx, dy = node_center[dst.id]
            for seg_start, seg_end, has_arrow in _back_edge_segments(
                sx, sy, dx, dy, node_w_emu, node_h_emu,
                back_channel, layout.vertical,
            ):
                parts.append(_straight_connector_xml(
                    next_id, seg_start, seg_end, has_arrow, layout,
                ))
                next_id += 1

            if e.condition:
                lx, ly = _label_near_site(
                    node_center[src.id], src_site, node_w_emu, node_h_emu,
                )
                parts.append(_label_xml(
                    next_id,
                    lx - label_w_emu // 2,
                    ly - label_h_emu // 2,
                    label_w_emu, label_h_emu,
                    e.condition, layout,
                ))
                next_id += 1
            continue

        # Forward edges: use bentConnector3 wired to shape connection points
        # so Excel auto-routes and the connector sticks to the shapes.
        src_site, dst_site = _pick_sites(
            src, dst, is_back, layout.vertical,
            lane_idx, edge_idx, n_siblings,
        )

        src_pt = _site_position(
            node_center[src.id], node_w_emu, node_h_emu, src_site
        )
        dst_pt = _site_position(
            node_center[dst.id], node_w_emu, node_h_emu, dst_site
        )

        bx, by, bw, bh, flip_h, flip_v = _connector_bbox(
            src_pt, dst_pt, is_back, layout.vertical,
        )

        parts.append(_bent_connector_xml(
            next_id,
            node_shape_id[src.id], src_site,
            node_shape_id[dst.id], dst_site,
            bx, by, bw, bh, flip_h, flip_v, layout,
        ))
        next_id += 1

        if e.condition:
            lx, ly = _label_near_site(
                node_center[src.id], src_site, node_w_emu, node_h_emu,
            )
            parts.append(_label_xml(
                next_id,
                lx - label_w_emu // 2,
                ly - label_h_emu // 2,
                label_w_emu, label_h_emu,
                e.condition, layout,
            ))
            next_id += 1

    parts.append("</xdr:wsDr>")
    return "\n".join(parts)


def _back_edge_segments(
    sx: int, sy: int, dx: int, dy: int,
    node_w: int, node_h: int,
    back_channel: int, vertical: bool,
) -> List[Tuple[Tuple[int, int], Tuple[int, int], bool]]:
    """Three straight segments forming a U-route in the back-edge channel."""
    hw = node_w // 2
    hh = node_h // 2
    if vertical:
        p0 = (sx - hw, sy)            # exit left of source
        p1 = (back_channel, sy)       # into left channel
        p2 = (back_channel, dy)       # up to target row
        p3 = (dx - hw, dy)            # right into target's left
    else:
        p0 = (sx, sy - hh)            # exit top of source
        p1 = (sx, back_channel)       # into top channel
        p2 = (dx, back_channel)       # across to target column
        p3 = (dx, dy - hh)            # down into target's top
    return [(p0, p1, False), (p1, p2, False), (p2, p3, True)]


def _straight_connector_xml(
    conn_id: int, start: Tuple[int, int], end: Tuple[int, int],
    has_arrow: bool, layout: _Layout,
) -> str:
    sx, sy = start
    ex, ey = end
    bx = min(sx, ex)
    by = min(sy, ey)
    dx_abs = abs(ex - sx)
    dy_abs = abs(ey - sy)

    # Keep the perpendicular dimension at 0 so corner points align exactly.
    if dy_abs == 0:
        bw = max(dx_abs, 1)
        bh = 0
    elif dx_abs == 0:
        bw = 0
        bh = max(dy_abs, 1)
    else:
        bw = dx_abs
        bh = dy_abs

    flips = []
    if has_arrow:
        if sx > ex:
            flips.append('flipH="1"')
        if sy > ey:
            flips.append('flipV="1"')
    flip_attr = (" " + " ".join(flips)) if flips else ""

    arrow = '<a:tailEnd type="triangle"/>' if has_arrow else ""
    anchor = _two_cell_anchor(bx, by, bw, bh, layout)

    return (
        '<xdr:twoCellAnchor editAs="oneCell">'
        f'{anchor}'
        '<xdr:cxnSp macro="">'
        f'<xdr:nvCxnSpPr><xdr:cNvPr id="{conn_id}" name="Back_{conn_id}"/>'
        '<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>'
        '<xdr:spPr>'
        f'<a:xfrm{flip_attr}><a:off x="{bx}" y="{by}"/><a:ext cx="{bw}" cy="{bh}"/></a:xfrm>'
        '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{EDGE_COLOR}"/></a:solidFill>'
        f'{arrow}</a:ln>'
        '</xdr:spPr>'
        '</xdr:cxnSp>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
    )


# --- connection-site selection ----------------------------------------------

def _pick_sites(
    src: Node, dst: Node, is_back: bool, vertical: bool,
    lane_idx: Dict[str, int], edge_idx: int, n_siblings: int,
) -> Tuple[int, int]:
    """Return (src_site, dst_site) for a connector.

    Multiple edges from the same source are placed on different sides so
    their labels do not pile up.
    """
    src_lane = lane_idx[src.lane_id]
    dst_lane = lane_idx[dst.lane_id]

    if vertical:
        if is_back:
            # exit LEFT / enter LEFT → Excel draws a U around the left
            return SITE_LEFT, SITE_LEFT

        if n_siblings == 1:
            return SITE_BOTTOM, SITE_TOP

        # multiple forward siblings: first goes straight down,
        # the rest exit sideways toward their destination lane
        if edge_idx == 0:
            return SITE_BOTTOM, SITE_TOP
        if dst_lane > src_lane:
            return SITE_RIGHT, SITE_TOP
        if dst_lane < src_lane:
            return SITE_LEFT, SITE_TOP
        # same-lane sibling: alternate left/right
        return (SITE_RIGHT if edge_idx % 2 == 1 else SITE_LEFT), SITE_TOP

    # horizontal flow
    if is_back:
        return SITE_TOP, SITE_TOP

    if n_siblings == 1:
        return SITE_RIGHT, SITE_LEFT

    if edge_idx == 0:
        return SITE_RIGHT, SITE_LEFT
    if dst_lane > src_lane:
        return SITE_BOTTOM, SITE_LEFT
    if dst_lane < src_lane:
        return SITE_TOP, SITE_LEFT
    return (SITE_BOTTOM if edge_idx % 2 == 1 else SITE_TOP), SITE_LEFT


def _site_position(
    center: Tuple[int, int], node_w: int, node_h: int, site: int,
) -> Tuple[int, int]:
    cx, cy = center
    hw = node_w // 2
    hh = node_h // 2
    if site == SITE_TOP:    return (cx, cy - hh)
    if site == SITE_RIGHT:  return (cx + hw, cy)
    if site == SITE_BOTTOM: return (cx, cy + hh)
    return (cx - hw, cy)    # SITE_LEFT


def _label_near_site(
    center: Tuple[int, int], site: int, node_w: int, node_h: int,
) -> Tuple[int, int]:
    """Anchor the label next to the connector's exit site on the source."""
    cx, cy = center
    hw = node_w // 2
    hh = node_h // 2
    gap = 14 * EMU_PER_PX
    if site == SITE_TOP:    return (cx, cy - hh - gap)
    if site == SITE_RIGHT:  return (cx + hw + gap, cy)
    if site == SITE_BOTTOM: return (cx, cy + hh + gap)
    return (cx - hw - gap, cy)


def _connector_bbox(
    src_pt: Tuple[int, int], dst_pt: Tuple[int, int],
    is_back: bool, vertical: bool,
) -> Tuple[int, int, int, int, bool, bool]:
    """Compute the xfrm hint bbox for bentConnector3.

    For back-edges we extend the bbox outward (left in vertical mode, up in
    horizontal mode) so that bentConnector3 has room to route the U-turn
    rather than collapsing to a straight line.
    """
    sx, sy = src_pt
    dx, dy = dst_pt

    if is_back and vertical:
        detour = 70 * EMU_PER_PX  # extend leftward past the back-channel
        bx = min(sx, dx) - detour
        by = min(sy, dy)
        bw = max(max(sx, dx) - bx, 1)
        bh = max(abs(dy - sy), 1)
    elif is_back and not vertical:
        detour = 70 * EMU_PER_PX  # extend upward past the top channel
        bx = min(sx, dx)
        by = min(sy, dy) - detour
        bw = max(abs(dx - sx), 1)
        bh = max(max(sy, dy) - by, 1)
    else:
        bx = min(sx, dx)
        by = min(sy, dy)
        bw = max(abs(dx - sx), 1)
        bh = max(abs(dy - sy), 1)

    flip_h = sx > dx
    flip_v = sy > dy
    return bx, by, bw, bh, flip_h, flip_v


def _node_shape_xml(
    shape_id: int, node: Node,
    x_emu: int, y_emu: int, w_emu: int, h_emu: int,
    layout: _Layout,
) -> str:
    preset = NODE_PRESET.get(node.type, "flowChartProcess")
    fill = NODE_FILL.get(node.type, "FFFFFF")
    text = xml_escape(node.label)
    name = xml_escape(f"{node.id}_{node.label}")
    anchor = _two_cell_anchor(x_emu, y_emu, w_emu, h_emu, layout)
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
    layout: _Layout,
) -> str:
    """Emit a ``bentConnector3`` wired to two shapes' connection points.

    Excel uses ``stCxn``/``endCxn`` to stick the endpoints to the referenced
    shapes, so moving a shape re-routes the connector automatically. The
    ``xfrm`` acts as a routing hint — for back-edges we expand it outward
    (see ``_connector_bbox``) to coax Excel into drawing a U instead of a
    straight line.
    """
    flips = []
    if flip_h:
        flips.append('flipH="1"')
    if flip_v:
        flips.append('flipV="1"')
    flip_attr = (" " + " ".join(flips)) if flips else ""

    anchor = _two_cell_anchor(bx, by, bw, bh, layout)

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
    text: str, layout: _Layout,
) -> str:
    t = xml_escape(text)
    anchor = _two_cell_anchor(x_emu, y_emu, w_emu, h_emu, layout)
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
