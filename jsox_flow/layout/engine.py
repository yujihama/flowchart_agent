"""Python wrapper around the dagre Node script.

Public entry point: :func:`compute_layout`. The function is side-effect free
(no file I/O), deterministic for a given Flow + options, and returns a fully
JSON-serialisable :class:`LayoutResult`. This is the shape an autonomous
agent is expected to consume.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from ..model import Flow
from .result import (
    EdgeLayout,
    LayoutError,
    LayoutMetrics,
    LayoutResult,
    NodeLayout,
)


# --- default geometry ------------------------------------------------------
# Keep in sync with the renderers; all dimensions in device-independent px.

DEFAULT_NODE_W = 130
DEFAULT_NODE_H = 46

# step = axis along which the flow progresses; lane = perpendicular axis.
HORIZONTAL_STEP_W = 154     # H_STEP_WIDTH_CHARS (22) * PX_PER_CHAR (7)
HORIZONTAL_LANE_H = 120     # H_ROW_LANE_HEIGHT_PT (90pt) * 96/72
VERTICAL_LANE_W = 168       # V_LANE_WIDTH_CHARS (24) * PX_PER_CHAR (7)
VERTICAL_STEP_H = 93        # V_ROW_STEP_HEIGHT_PT (70pt) * 96/72


@dataclass
class LayoutOptions:
    """Agent-tunable knobs for :func:`compute_layout`."""
    orientation: str = "horizontal"            # "horizontal" | "vertical"
    node_width: int = DEFAULT_NODE_W
    node_height: int = DEFAULT_NODE_H
    step_size: Optional[int] = None            # px along flow axis per rank
    lane_size: Optional[int] = None            # px perpendicular to flow axis per lane
    dagre_nodesep: int = 30
    dagre_edgesep: int = 15
    dagre_ranksep: int = 60
    dagre_ranker: str = "network-simplex"      # "network-simplex" | "tight-tree" | "longest-path"

    def resolved_step(self) -> int:
        if self.step_size is not None:
            return self.step_size
        return HORIZONTAL_STEP_W if self.orientation == "horizontal" else VERTICAL_STEP_H

    def resolved_lane(self) -> int:
        if self.lane_size is not None:
            return self.lane_size
        return HORIZONTAL_LANE_H if self.orientation == "horizontal" else VERTICAL_LANE_W

    def to_dict(self) -> Dict[str, Any]:
        return {
            "orientation": self.orientation,
            "node_width": self.node_width,
            "node_height": self.node_height,
            "step_size": self.resolved_step(),
            "lane_size": self.resolved_lane(),
            "dagre_nodesep": self.dagre_nodesep,
            "dagre_edgesep": self.dagre_edgesep,
            "dagre_ranksep": self.dagre_ranksep,
            "dagre_ranker": self.dagre_ranker,
        }


# --- public API ------------------------------------------------------------

def compute_layout(flow: Flow, options: Optional[LayoutOptions] = None) -> LayoutResult:
    """Compute a swimlane-aware layout for ``flow``.

    Steps:
        1. Invoke the dagre Node script with node / edge metadata.
        2. Derive a discrete ``rank`` per node by bucketing dagre's
           flow-axis coordinate (x for horizontal, y for vertical).
        3. Override the perpendicular axis with the lane index so that
           the swimlane constraint is strictly honoured.
        4. Compute final pixel positions on the target grid.
        5. Collect metrics useful for an agent (collisions, end-rank,
           nodes placed past the end).

    Raises :class:`LayoutError` with a machine-readable ``kind`` on any
    problem. Never raises bare exceptions.
    """
    opts = options or LayoutOptions()
    _validate_flow(flow)

    back_edges = _detect_back_edges(flow)
    payload = _build_dagre_payload(flow, opts, back_edges)
    raw = _run_dagre(payload)

    try:
        dagre_out = json.loads(raw)
    except json.JSONDecodeError as e:
        raise LayoutError(
            LayoutError.KIND_ENGINE_FAILED,
            f"dagre returned non-JSON output: {e}",
            {"raw": raw[:500]},
        )

    return _assemble_result(flow, opts, dagre_out, back_edges)


# --- helpers ---------------------------------------------------------------

def _validate_flow(flow: Flow) -> None:
    lane_ids = {l.id for l in flow.lanes}
    for n in flow.nodes:
        if n.lane_id not in lane_ids:
            raise LayoutError(
                LayoutError.KIND_UNKNOWN_LANE,
                f"Node {n.id!r} references unknown lane {n.lane_id!r}.",
                {"node_id": n.id, "lane_id": n.lane_id},
            )
    node_ids = {n.id for n in flow.nodes}
    for e in flow.edges:
        if e.from_id not in node_ids or e.to_id not in node_ids:
            raise LayoutError(
                LayoutError.KIND_INVALID_FLOW,
                f"Edge references unknown node: {e.from_id} -> {e.to_id}.",
                {"from_id": e.from_id, "to_id": e.to_id},
            )


def _detect_back_edges(flow: Flow) -> Set[Tuple[str, str]]:
    """DFS cycle detection, preferring start nodes as roots.

    Edges that close a cycle become "back" edges. They are withheld from
    dagre so that the rank assignment reflects the forward process order
    and is not perturbed by dagre's own feedback-arc heuristic.
    """
    outgoing: Dict[str, List[str]] = {n.id: [] for n in flow.nodes}
    for e in flow.edges:
        outgoing.setdefault(e.from_id, []).append(e.to_id)

    visited: Set[str] = set()
    on_stack: Set[str] = set()
    back: Set[Tuple[str, str]] = set()

    def visit(nid: str) -> None:
        on_stack.add(nid)
        for tgt in outgoing.get(nid, []):
            if tgt in on_stack:
                back.add((nid, tgt))
            elif tgt not in visited:
                visit(tgt)
        on_stack.discard(nid)
        visited.add(nid)

    for n in flow.nodes:
        if n.type == "start" and n.id not in visited:
            visit(n.id)
    for n in flow.nodes:
        if n.id not in visited:
            visit(n.id)
    return back


def _build_dagre_payload(
    flow: Flow, opts: LayoutOptions, back_edges: Set[Tuple[str, str]]
) -> Dict[str, Any]:
    forward = [
        {"from": e.from_id, "to": e.to_id}
        for e in flow.edges
        if (e.from_id, e.to_id) not in back_edges
    ]
    return {
        "orientation": opts.orientation,
        "nodes": [
            {
                "id": n.id,
                "width": opts.node_width,
                "height": opts.node_height,
                "label": n.label,
            }
            for n in flow.nodes
        ],
        "edges": forward,
        "options": {
            "nodesep": opts.dagre_nodesep,
            "edgesep": opts.dagre_edgesep,
            "ranksep": opts.dagre_ranksep,
            "ranker": opts.dagre_ranker,
        },
    }


def _script_path() -> Path:
    return Path(__file__).resolve().parent / "dagre_layout.js"


def _find_node() -> str:
    exe = os.environ.get("JSOX_NODE") or shutil.which("node")
    if not exe:
        raise LayoutError(
            LayoutError.KIND_ENGINE_NOT_FOUND,
            "Node.js ('node') not found on PATH. Install Node or set JSOX_NODE.",
        )
    return exe


def _run_dagre(payload: Dict[str, Any]) -> str:
    node_exe = _find_node()
    script = _script_path()
    if not script.exists():
        raise LayoutError(
            LayoutError.KIND_ENGINE_NOT_FOUND,
            f"dagre_layout.js not found at {script}",
            {"path": str(script)},
        )

    try:
        proc = subprocess.run(
            [node_exe, str(script)],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
            cwd=str(script.parent),
        )
    except subprocess.TimeoutExpired as e:
        raise LayoutError(
            LayoutError.KIND_ENGINE_FAILED,
            "dagre process timed out after 30s.",
            {"stderr": e.stderr or ""},
        )
    except FileNotFoundError as e:
        raise LayoutError(
            LayoutError.KIND_ENGINE_NOT_FOUND,
            f"failed to spawn node: {e}",
        )

    if proc.returncode != 0:
        raise LayoutError(
            LayoutError.KIND_ENGINE_FAILED,
            f"dagre exited with code {proc.returncode}: {proc.stderr.strip()}",
            {"returncode": proc.returncode, "stderr": proc.stderr},
        )
    return proc.stdout


def _assemble_result(
    flow: Flow,
    opts: LayoutOptions,
    dagre_out: Dict[str, Any],
    back_edges: Set[Tuple[str, str]],
) -> LayoutResult:
    dagre_nodes: Dict[str, Dict[str, int]] = dagre_out.get("nodes") or {}
    dagre_edges: List[Dict[str, Any]] = dagre_out.get("edges") or []

    horizontal = opts.orientation == "horizontal"

    # Derive ranks by bucketing the flow-axis coordinate.
    flow_axis = "x" if horizontal else "y"
    coords = sorted({dagre_nodes[n.id][flow_axis] for n in flow.nodes})
    rank_by_coord: Dict[int, int] = {c: i for i, c in enumerate(coords)}
    rank_by_node: Dict[str, int] = {
        n.id: rank_by_coord[dagre_nodes[n.id][flow_axis]] for n in flow.nodes
    }

    lane_order = [l.id for l in flow.lanes]
    lane_idx = {lane_id: i for i, lane_id in enumerate(lane_order)}

    step_size = opts.resolved_step()
    lane_size = opts.resolved_lane()

    # Detect cells with multiple nodes so we can stagger them inside the
    # lane (rather than push them past the end node). Order siblings by
    # dagre's perpendicular coordinate to preserve dagre's own ordering.
    secondary_axis = "y" if horizontal else "x"
    cell_members: Dict[Tuple[str, int], List[str]] = {}
    for n in flow.nodes:
        cell_members.setdefault((n.lane_id, rank_by_node[n.id]), []).append(n.id)
    slot_offset: Dict[str, int] = {}
    for ids in cell_members.values():
        if len(ids) == 1:
            slot_offset[ids[0]] = 0
            continue
        ids.sort(key=lambda nid: dagre_nodes[nid][secondary_axis])
        slot_gap = (opts.node_height if horizontal else opts.node_width) + 14
        span = slot_gap * (len(ids) - 1)
        for i, nid in enumerate(ids):
            slot_offset[nid] = int(-span / 2 + i * slot_gap)

    node_layouts: Dict[str, NodeLayout] = {}
    for n in flow.nodes:
        rank = rank_by_node[n.id]
        li = lane_idx[n.lane_id]
        offset = slot_offset.get(n.id, 0)
        if horizontal:
            x = step_size * rank + step_size // 2
            y = lane_size * li + lane_size // 2 + offset
        else:
            x = lane_size * li + lane_size // 2 + offset
            y = step_size * rank + step_size // 2
        node_layouts[n.id] = NodeLayout(
            id=n.id,
            lane_id=n.lane_id,
            rank=rank,
            x=x,
            y=y,
            width=opts.node_width,
            height=opts.node_height,
        )

    edge_layouts: List[EdgeLayout] = []
    for e in flow.edges:
        is_back = (e.from_id, e.to_id) in back_edges
        points = _waypoints_for(e, dagre_edges)
        edge_layouts.append(
            EdgeLayout(
                from_id=e.from_id,
                to_id=e.to_id,
                is_back=is_back,
                condition=e.condition,
                points=points,
            )
        )

    metrics = _compute_metrics(flow, node_layouts, edge_layouts, lane_order)
    warnings = _derive_warnings(metrics, lane_size, opts)

    max_rank = metrics.max_rank or 0
    n_ranks = max_rank + 1 if node_layouts else 0
    if horizontal:
        canvas_w = n_ranks * step_size
        canvas_h = len(lane_order) * lane_size
    else:
        canvas_w = len(lane_order) * lane_size
        canvas_h = n_ranks * step_size

    return LayoutResult(
        orientation=opts.orientation,
        nodes=node_layouts,
        edges=edge_layouts,
        lane_order=lane_order,
        canvas_width=canvas_w,
        canvas_height=canvas_h,
        metrics=metrics,
        warnings=warnings,
    )


def _waypoints_for(
    edge, dagre_edges: List[Dict[str, Any]]
) -> List[tuple[int, int]]:
    for de in dagre_edges:
        if de.get("from") == edge.from_id and de.get("to") == edge.to_id:
            return [(int(p[0]), int(p[1])) for p in (de.get("points") or [])]
    return []


def _compute_metrics(
    flow: Flow,
    nodes: Dict[str, NodeLayout],
    edges: List[EdgeLayout],
    lane_order: List[str],
) -> LayoutMetrics:
    cells: Dict[tuple, List[str]] = {}
    for nid, nl in nodes.items():
        cells.setdefault((nl.lane_id, nl.rank), []).append(nid)
    collisions: List[tuple[str, str]] = []
    max_cell_size = 1
    for ids in cells.values():
        if len(ids) > 1:
            max_cell_size = max(max_cell_size, len(ids))
            ids = sorted(ids)
            for i in range(len(ids)):
                for j in range(i + 1, len(ids)):
                    collisions.append((ids[i], ids[j]))

    lane_span: Dict[str, List[int]] = {}
    for nid, nl in nodes.items():
        span = lane_span.get(nl.lane_id)
        if span is None:
            lane_span[nl.lane_id] = [nl.rank, nl.rank]
        else:
            span[0] = min(span[0], nl.rank)
            span[1] = max(span[1], nl.rank)

    end_ranks = [nl.rank for n in flow.nodes if n.type == "end" for nl in [nodes[n.id]]]
    end_rank = max(end_ranks) if end_ranks else None
    max_rank = max((nl.rank for nl in nodes.values()), default=None)

    nodes_past_end: List[str] = []
    if end_rank is not None:
        nodes_past_end = sorted(
            nid for nid, nl in nodes.items() if nl.rank > end_rank
        )

    n_back = sum(1 for e in edges if e.is_back)
    n_ranks = (max_rank + 1) if max_rank is not None else 0

    return LayoutMetrics(
        n_nodes=len(nodes),
        n_edges=len(edges),
        n_back_edges=n_back,
        n_ranks=n_ranks,
        n_lanes=len(lane_order),
        collisions=collisions,
        lane_span=lane_span,
        end_rank=end_rank,
        max_rank=max_rank,
        nodes_past_end=nodes_past_end,
        max_cell_size=max_cell_size,
    )


def _derive_warnings(
    metrics: LayoutMetrics, lane_size: int, opts: "LayoutOptions"
) -> List[str]:
    out: List[str] = []
    # Cell collisions are resolved by in-lane stacking; flag them so the
    # caller knows some cells are doubly-occupied and may look tighter.
    for a, b in metrics.collisions:
        out.append(f"cell-stacked: {a} and {b} share the same (lane, rank).")
    for nid in metrics.nodes_past_end:
        out.append(f"node {nid} is placed past the end node (rank > end_rank).")

    # Lane overflow: if too many nodes are stacked, they exceed the lane.
    node_extent = opts.node_height if opts.orientation == "horizontal" else opts.node_width
    if metrics.max_cell_size >= 2:
        needed = (node_extent + 14) * metrics.max_cell_size
        if needed > lane_size:
            out.append(
                f"lane-overflow: up to {metrics.max_cell_size} nodes per cell "
                f"need ~{needed}px but lane_size={lane_size}px; "
                "consider increasing LayoutOptions.lane_size."
            )
    return out
