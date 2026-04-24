"""SVG renderer with real horizontal swimlanes.

Layout algorithm
----------------
1. Assign each node to a column via longest-path from `start` (back edges
   — edges that would create a cycle — are ignored for column assignment).
2. Assign each node to a row = index of its lane in `flow.lanes`.
3. Place node centers at
       x = LANE_HEADER_W + col * COL_W + COL_W / 2
       y = MARGIN       + row * LANE_H + LANE_H / 2
4. Edge routing:
       * forward (col grows):   L-shape through vertical midline
       * backward (col shrinks): route over the top margin band
"""
from __future__ import annotations

from typing import Dict, List, Tuple

from .._layout import assign_columns as _assign_columns
from .._layout import back_edges as _back_edges
from .._layout import resolve_cell_collisions as _resolve_cell_collisions
from ..model import Edge, Flow, Node

# --- geometry ---------------------------------------------------------------

MARGIN = 30
LANE_HEADER_W = 120
COL_W = 190
LANE_H = 110
NODE_W = 140
NODE_H = 52

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

def to_svg(flow: Flow) -> str:
    cols = _assign_columns(flow)
    cols = _resolve_cell_collisions(flow, cols)
    max_col = max(cols.values()) if cols else 0
    lane_row = {lane.id: i for i, lane in enumerate(flow.lanes)}

    width = LANE_HEADER_W + (max_col + 1) * COL_W + MARGIN
    height = MARGIN * 2 + len(flow.lanes) * LANE_H

    positions: Dict[str, Tuple[float, float]] = {}
    for n in flow.nodes:
        cx = LANE_HEADER_W + cols[n.id] * COL_W + COL_W / 2
        cy = MARGIN + lane_row[n.lane_id] * LANE_H + LANE_H / 2
        positions[n.id] = (cx, cy)

    back_edges = _back_edges(flow, cols)

    parts: List[str] = []
    parts.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" '
        f'height="{height}" viewBox="0 0 {width} {height}" '
        f'font-family="-apple-system, Segoe UI, Meiryo, sans-serif" font-size="13">'
    )
    parts.append(_defs())

    # swimlanes (drawn first so edges/nodes sit on top)
    parts.extend(_render_lanes(flow, width))

    # edges (behind nodes)
    for e in flow.edges:
        parts.extend(
            _render_edge(e, positions, cols, is_back=(e.from_id, e.to_id) in back_edges)
        )

    # nodes
    for n in flow.nodes:
        parts.extend(_render_node(n, positions[n.id]))

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


def _render_lanes(flow: Flow, width: float) -> List[str]:
    out: List[str] = []
    for i, lane in enumerate(flow.lanes):
        y = MARGIN + i * LANE_H
        fill = LANE_FILL_EVEN if i % 2 == 0 else LANE_FILL_ODD
        out.append(
            f'<rect x="0" y="{y}" width="{width}" height="{LANE_H}" '
            f'fill="{fill}" stroke="{LANE_STROKE}"/>'
        )
        out.append(
            f'<rect x="0" y="{y}" width="{LANE_HEADER_W}" height="{LANE_H}" '
            f'fill="{LANE_HEADER_FILL}" stroke="{LANE_STROKE}"/>'
        )
        out.append(
            f'<text x="{LANE_HEADER_W / 2}" y="{y + LANE_H / 2}" '
            f'text-anchor="middle" dominant-baseline="middle" '
            f'font-weight="bold">{_escape(lane.name)}</text>'
        )
    return out


def _render_node(node: Node, pos: Tuple[float, float]) -> List[str]:
    cx, cy = pos
    x = cx - NODE_W / 2
    y = cy - NODE_H / 2
    fill = NODE_FILLS.get(node.type, NODE_FILLS["task"])
    label = _escape(node.label)
    out: List[str] = []

    if node.type in ("start", "end"):
        rx = NODE_H / 2
        out.append(
            f'<rect x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" '
            f'rx="{rx}" ry="{rx}" fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "decision":
        # diamond fits within the same NODE_W x NODE_H bbox as other shapes
        # so every edge uses the same attachment geometry
        pts = (
            f"{cx},{cy - NODE_H / 2} "
            f"{cx + NODE_W / 2},{cy} "
            f"{cx},{cy + NODE_H / 2} "
            f"{cx - NODE_W / 2},{cy}"
        )
        out.append(
            f'<polygon points="{pts}" fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "subprocess":
        out.append(
            f'<rect x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" '
            f'fill="{fill}" stroke="{NODE_STROKE}"/>'
        )
        out.append(
            f'<line x1="{x + 8}" y1="{y}" x2="{x + 8}" y2="{y + NODE_H}" '
            f'stroke="{NODE_STROKE}"/>'
        )
        out.append(
            f'<line x1="{x + NODE_W - 8}" y1="{y}" x2="{x + NODE_W - 8}" '
            f'y2="{y + NODE_H}" stroke="{NODE_STROKE}"/>'
        )
    elif node.type == "document":
        path = (
            f"M{x},{y} L{x + NODE_W},{y} L{x + NODE_W},{y + NODE_H - 8} "
            f"Q{x + NODE_W * 0.75},{y + NODE_H + 4} "
            f"{x + NODE_W / 2},{y + NODE_H - 8} "
            f"Q{x + NODE_W * 0.25},{y + NODE_H - 20} {x},{y + NODE_H - 8} Z"
        )
        out.append(f'<path d="{path}" fill="{fill}" stroke="{NODE_STROKE}"/>')
    else:  # task
        out.append(
            f'<rect x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" '
            f'fill="{fill}" stroke="{NODE_STROKE}"/>'
        )

    out.append(
        f'<text x="{cx}" y="{cy}" text-anchor="middle" '
        f'dominant-baseline="middle">{label}</text>'
    )
    return out


def _render_edge(
    edge: Edge,
    positions: Dict[str, Tuple[float, float]],
    cols: Dict[str, int],
    *,
    is_back: bool,
) -> List[str]:
    src_x, src_y = positions[edge.from_id]
    dst_x, dst_y = positions[edge.to_id]
    out: List[str] = []

    if not is_back:
        sx = src_x + NODE_W / 2
        sy = src_y
        dx = dst_x - NODE_W / 2
        dy = dst_y
        mid_x = (sx + dx) / 2
        path = f"M{sx},{sy} L{mid_x},{sy} L{mid_x},{dy} L{dx},{dy}"
        label_x, label_y = sx + 8, sy - 6
    else:
        # exit top of src, route over the top margin, enter top of dst
        sx = src_x
        sy = src_y - NODE_H / 2
        dx = dst_x
        dy = dst_y - NODE_H / 2
        top_y = MARGIN / 2 - 8
        path = f"M{sx},{sy} L{sx},{top_y} L{dx},{top_y} L{dx},{dy}"
        label_x, label_y = sx - 18, sy - 6

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
