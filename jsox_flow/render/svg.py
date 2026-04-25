"""SVG renderer.

Consumes a :class:`LayoutResult` from :mod:`jsox_flow.layout`, so the
rendering pipeline is purely presentational: no rank-assignment logic
lives here.
"""
from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from ..layout import LayoutOptions, LayoutResult, compute_layout
from ..model import Edge, Flow, Node

# --- geometry ---------------------------------------------------------------

MARGIN = 30
LANE_HEADER_W_PX = 120

LANE_FILL_EVEN = "#f7f9fc"
LANE_FILL_ODD = "#eef2f8"
LANE_HEADER_FILL = "#d4dff0"
LANE_STROKE = "#8a99b3"

NODE_STROKE = "#333"
NODE_FILLS = {
    "start":      "#fff3cd",
    "end":        "#fff3cd",
    "task":       "#ffffff",
    "decision":   "#d1ecf1",
    "subprocess": "#e2e3e5",
    "document":   "#fefefe",
}

EDGE_STROKE = "#333"
COND_COLOR = "#c0392b"


# --- public API -------------------------------------------------------------

def to_svg(flow: Flow, *, layout: Optional[LayoutResult] = None) -> str:
    """Render ``flow`` as SVG.

    Pass a pre-computed ``layout`` to reuse the same positions the XLSX
    renderer saw; otherwise a default horizontal layout is computed.
    """
    if layout is None:
        layout = compute_layout(flow, LayoutOptions(orientation="horizontal"))
    if layout.orientation != "horizontal":
        raise ValueError("SVG renderer currently only supports horizontal layouts.")

    node_by_id = {n.id: n for n in flow.nodes}
    lane_by_id = {l.id: l for l in flow.lanes}

    lane_h = layout.canvas_height // max(len(layout.lane_order), 1)
    body_w = layout.canvas_width
    width = LANE_HEADER_W_PX + body_w + MARGIN
    height = MARGIN * 2 + layout.canvas_height

    parts: List[str] = []
    parts.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" '
        f'height="{height}" viewBox="0 0 {width} {height}" '
        f'font-family="-apple-system, Segoe UI, Meiryo, sans-serif" font-size="13">'
    )
    parts.append(_defs())
    parts.extend(_render_lanes(layout, lane_by_id, lane_h, width))

    # shift nodes by LANE_HEADER_W + MARGIN (body origin), lane header bar on left
    ox = LANE_HEADER_W_PX
    oy = MARGIN

    for e in layout.edges:
        parts.extend(_render_edge(e, layout, node_by_id, ox, oy))

    for nid, nl in layout.nodes.items():
        parts.extend(_render_node(node_by_id[nid], nl.x + ox, nl.y + oy,
                                  nl.width, nl.height))

    parts.append("</svg>")
    return "\n".join(parts)


# --- rendering helpers ------------------------------------------------------

def _defs() -> str:
    return (
        '<defs>'
        '<marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" '
        'markerWidth="8" markerHeight="8" orient="auto-start-reverse">'
        '<path d="M0,0 L10,5 L0,10 z" fill="#333"/>'
        '</marker>'
        '</defs>'
    )


def _render_lanes(
    layout: LayoutResult, lane_by_id: Dict[str, object],
    lane_h: int, width: int,
) -> List[str]:
    out: List[str] = []
    for i, lane_id in enumerate(layout.lane_order):
        lane = lane_by_id[lane_id]
        y = MARGIN + i * lane_h
        fill = LANE_FILL_EVEN if i % 2 == 0 else LANE_FILL_ODD
        out.append(
            f'<rect x="0" y="{y}" width="{width}" height="{lane_h}" '
            f'fill="{fill}" stroke="{LANE_STROKE}"/>'
        )
        out.append(
            f'<rect x="0" y="{y}" width="{LANE_HEADER_W_PX}" height="{lane_h}" '
            f'fill="{LANE_HEADER_FILL}" stroke="{LANE_STROKE}"/>'
        )
        out.append(
            f'<text x="{LANE_HEADER_W_PX / 2}" y="{y + lane_h / 2}" '
            f'text-anchor="middle" dominant-baseline="middle" '
            f'font-weight="bold">{_escape(lane.name)}</text>'
        )
    return out


def _render_node(
    node: Node, cx: float, cy: float, w: int, h: int,
) -> List[str]:
    x = cx - w / 2
    y = cy - h / 2
    fill = NODE_FILLS.get(node.type, NODE_FILLS["task"])
    label = _escape(node.label)
    out: List[str] = []

    if node.type in ("start", "end"):
        rx = h / 2
        out.append(
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
            f'rx="{rx}" ry="{rx}" fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "decision":
        pts = (
            f"{cx},{cy - h / 2} "
            f"{cx + w / 2},{cy} "
            f"{cx},{cy + h / 2} "
            f"{cx - w / 2},{cy}"
        )
        out.append(
            f'<polygon points="{pts}" fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "subprocess":
        out.append(
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
            f'fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
        out.append(
            f'<line x1="{x + 8}" y1="{y}" x2="{x + 8}" y2="{y + h}" '
            f'stroke="{NODE_STROKE}"/>'
        )
        out.append(
            f'<line x1="{x + w - 8}" y1="{y}" x2="{x + w - 8}" '
            f'y2="{y + h}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "document":
        path = (
            f"M{x},{y} L{x + w},{y} L{x + w},{y + h - 8} "
            f"Q{x + w * 0.75},{y + h + 4} "
            f"{x + w / 2},{y + h - 8} "
            f"Q{x + w * 0.25},{y + h - 20} {x},{y + h - 8} Z"
        )
        out.append(f'<path d="{path}" fill="{fill}" stroke="{NODE_STROKE}"/>')
    else:
        out.append(
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
            f'fill="{fill}" stroke="{NODE_STROKE}"/>'
        )

    out.append(
        f'<text x="{cx}" y="{cy}" text-anchor="middle" '
        f'dominant-baseline="middle">{_escape(node.label)}</text>'
    )
    return out


def _render_edge(
    edge, layout: LayoutResult, node_by_id: Dict[str, Node],
    ox: int, oy: int,
) -> List[str]:
    src_nl = layout.nodes[edge.from_id]
    dst_nl = layout.nodes[edge.to_id]
    sx, sy = src_nl.x + ox, src_nl.y + oy
    dx, dy = dst_nl.x + ox, dst_nl.y + oy
    hw_s = src_nl.width // 2
    hh_s = src_nl.height // 2
    hw_d = dst_nl.width // 2
    hh_d = dst_nl.height // 2

    out: List[str] = []

    if not edge.is_back:
        ssx = sx + hw_s
        ssy = sy
        ddx = dx - hw_d
        ddy = dy
        mid_x = (ssx + ddx) / 2
        path = f"M{ssx},{ssy} L{mid_x},{ssy} L{mid_x},{ddy} L{ddx},{ddy}"
        label_x, label_y = ssx + 8, ssy - 6
    else:
        # route over the top margin band, exit top of src, enter top of dst
        ssx = sx
        ssy = sy - hh_s
        ddx = dx
        ddy = dy - hh_d
        top_y = MARGIN / 2 - 8
        path = f"M{ssx},{ssy} L{ssx},{top_y} L{ddx},{top_y} L{ddx},{ddy}"
        label_x, label_y = ssx - 18, ssy - 6

    out.append(
        f'<path d="{path}" fill="none" stroke="{EDGE_STROKE}" '
        f'stroke-width="1.5" marker-end="url(#arrow)"/>'
    )

    if edge.condition:
        out.append(
            f'<text x="{label_x}" y="{label_y}" fill="{COND_COLOR}" '
            f'font-weight="bold">{_escape(edge.condition)}</text>'
        )
    return out


def _escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
