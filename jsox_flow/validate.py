from __future__ import annotations

from collections import deque
from typing import List

from .model import Flow


class ValidationError(Exception):
    """Hard failure. Flow cannot be rendered as-is."""


def validate(flow: Flow) -> List[str]:
    """Validate a Flow.

    Raises ValidationError on fatal issues (broken refs, missing start/end, etc.).
    Returns a list of non-fatal warnings (isolated nodes, unreachable nodes, ...).
    """
    errors: List[str] = []
    warnings: List[str] = []

    lane_ids = {l.id for l in flow.lanes}
    node_ids = {n.id for n in flow.nodes}

    if len(lane_ids) != len(flow.lanes):
        errors.append("Duplicate lane id detected.")
    if len(node_ids) != len(flow.nodes):
        errors.append("Duplicate node id detected.")

    starts = [n for n in flow.nodes if n.type == "start"]
    ends = [n for n in flow.nodes if n.type == "end"]
    if not starts:
        errors.append("No start node.")
    if not ends:
        errors.append("No end node.")

    for n in flow.nodes:
        if n.lane_id not in lane_ids:
            errors.append(f"Node {n.id} references unknown lane '{n.lane_id}'.")

    for e in flow.edges:
        if e.from_id not in node_ids:
            errors.append(f"Edge references unknown node: from_id='{e.from_id}'.")
        if e.to_id not in node_ids:
            errors.append(f"Edge references unknown node: to_id='{e.to_id}'.")

    for n in flow.nodes:
        if n.type != "decision":
            continue
        out = flow.outgoing(n.id)
        if len(out) < 2:
            errors.append(f"Decision node {n.id} has fewer than 2 outgoing edges.")
        for e in out:
            if not e.condition:
                errors.append(
                    f"Decision edge {e.from_id}->{e.to_id} has no 'condition'."
                )

    if errors:
        raise ValidationError("\n".join(errors))

    touched = {e.from_id for e in flow.edges} | {e.to_id for e in flow.edges}
    for n in flow.nodes:
        if n.id not in touched and n.type not in ("start", "end"):
            warnings.append(f"Isolated node: {n.id} ({n.label}).")

    if starts:
        reachable = _bfs_forward(flow, starts[0].id)
        for n in flow.nodes:
            if n.id not in reachable:
                warnings.append(
                    f"Node {n.id} ({n.label}) is not reachable from start."
                )

    return warnings


def _bfs_forward(flow: Flow, src: str) -> set[str]:
    seen = {src}
    q = deque([src])
    while q:
        cur = q.popleft()
        for e in flow.outgoing(cur):
            if e.to_id not in seen:
                seen.add(e.to_id)
                q.append(e.to_id)
    return seen
