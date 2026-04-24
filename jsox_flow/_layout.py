"""Layout utilities shared across renderers.

* Column assignment: longest-path from start, ignoring back edges.
* Back edge detection via DFS.

Both the SVG and the XLSX renderer rely on these so that the logical
positions (column/lane) are identical across output formats.
"""
from __future__ import annotations

from typing import Dict, List, Set, Tuple

from .model import Flow


def topo_order(flow: Flow) -> Tuple[List[str], Set[Tuple[str, str]]]:
    """DFS topological order; returns (order, back_edges)."""
    visited: Set[str] = set()
    on_stack: Set[str] = set()
    back: Set[Tuple[str, str]] = set()
    order: List[str] = []

    def visit(nid: str) -> None:
        if nid in visited:
            return
        on_stack.add(nid)
        for e in flow.outgoing(nid):
            if e.to_id in on_stack:
                back.add((e.from_id, e.to_id))
            elif e.to_id not in visited:
                visit(e.to_id)
        on_stack.discard(nid)
        visited.add(nid)
        order.append(nid)

    for s in (n.id for n in flow.nodes if n.type == "start"):
        visit(s)
    for n in flow.nodes:
        if n.id not in visited:
            visit(n.id)

    order.reverse()
    return order, back


def assign_columns(flow: Flow) -> Dict[str, int]:
    """Longest-path column assignment; back edges do not extend depth."""
    order, back = topo_order(flow)
    depth: Dict[str, int] = {nid: 0 for nid in order}
    for nid in order:
        for e in flow.incoming(nid):
            if (e.from_id, e.to_id) in back:
                continue
            if e.from_id in depth:
                depth[nid] = max(depth[nid], depth[e.from_id] + 1)
    for n in flow.nodes:
        depth.setdefault(n.id, 0)
    return depth


def back_edges(flow: Flow, cols: Dict[str, int]) -> Set[Tuple[str, str]]:
    """Edges whose target column is strictly before the source's column."""
    return {
        (e.from_id, e.to_id)
        for e in flow.edges
        if cols.get(e.to_id, 0) < cols.get(e.from_id, 0)
    }


def resolve_cell_collisions(flow: Flow, cols: Dict[str, int]) -> Dict[str, int]:
    """Bump nodes that share a (lane, step) cell to the next free step.

    Long-path column assignment can place multiple nodes at the same
    (lane, step) — typically a rework branch that loops back to an
    earlier node (e.g. n7 → n8 → n6) ends up at the same step as the
    "main" sibling. We keep the mainline node in place and push the
    rework (side-branch) one forward until it has its own cell.

    Priority to stay put (highest first):
        2: nodes with any forward outgoing edge (mainline)
        1: nodes with no outgoing at all (start / end / isolated)
        0: side-branches whose only outgoing edges are back-edges
    """
    initial_cols = dict(cols)  # detect "back" using the original step
    forward_out = {n.id: 0 for n in flow.nodes}
    total_out = {n.id: 0 for n in flow.nodes}
    for e in flow.edges:
        total_out[e.from_id] += 1
        if initial_cols.get(e.to_id, 0) > initial_cols.get(e.from_id, 0):
            forward_out[e.from_id] += 1

    def priority(nid: str) -> int:
        if forward_out[nid] > 0:
            return 2
        if total_out[nid] == 0:
            return 1
        return 0  # only back-edge outgoing → side branch

    for _ in range(len(flow.nodes) * 4):
        occupancy: Dict[Tuple[str, int], list] = {}
        for n in flow.nodes:
            occupancy.setdefault((n.lane_id, cols[n.id]), []).append(n)

        collided = [ns for ns in occupancy.values() if len(ns) > 1]
        if not collided:
            break

        for nodes in collided:
            nodes.sort(key=lambda n: (-priority(n.id), n.id))
            for n in nodes[1:]:
                cols[n.id] += 1
    return cols
