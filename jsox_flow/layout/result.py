"""Layout result types.

The layout layer is intentionally JSON-first so an agent can:

1. Call ``compute_layout`` → inspect ``LayoutResult.to_dict()``
2. Evaluate via ``result.metrics`` (collisions, crossings, wasted cells)
3. Adjust the input Flow or layout options and retry

Everything here is a plain dataclass with deterministic ``to_dict`` /
``from_dict`` so round-trips through JSON preserve content.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Tuple


# --- error -----------------------------------------------------------------

class LayoutError(Exception):
    """Structured error with a machine-readable ``kind``.

    ``kind`` is one of the string constants below so callers (including
    autonomous agents) can branch on failure mode without string parsing
    the message.
    """

    KIND_ENGINE_NOT_FOUND = "engine_not_found"
    KIND_ENGINE_FAILED = "engine_failed"
    KIND_INVALID_FLOW = "invalid_flow"
    KIND_COLLISION_UNRESOLVABLE = "collision_unresolvable"
    KIND_UNKNOWN_LANE = "unknown_lane"

    def __init__(self, kind: str, message: str, details: Optional[Dict[str, Any]] = None):
        self.kind = kind
        self.message = message
        self.details = details or {}
        super().__init__(f"[{kind}] {message}")

    def to_dict(self) -> Dict[str, Any]:
        return {"kind": self.kind, "message": self.message, "details": self.details}


# --- layout records --------------------------------------------------------

@dataclass
class NodeLayout:
    id: str
    lane_id: str
    rank: int                         # 0-indexed step along the flow axis
    x: int                            # final centre, px
    y: int                            # final centre, px
    width: int                        # px
    height: int                       # px


@dataclass
class EdgeLayout:
    from_id: str
    to_id: str
    is_back: bool                     # rank[to] <= rank[from]
    condition: Optional[str] = None
    points: List[Tuple[int, int]] = field(default_factory=list)  # dagre waypoints (informational)


@dataclass
class LayoutMetrics:
    n_nodes: int
    n_edges: int
    n_back_edges: int
    n_ranks: int
    n_lanes: int
    collisions: List[Tuple[str, str]] = field(default_factory=list)  # pairs sharing (lane, rank)
    lane_span: Dict[str, List[int]] = field(default_factory=dict)    # lane_id -> [min_rank, max_rank]
    end_rank: Optional[int] = None                                   # max rank of any end node
    max_rank: Optional[int] = None                                   # overall max rank
    nodes_past_end: List[str] = field(default_factory=list)          # nodes with rank > end_rank
    max_cell_size: int = 1                                           # max nodes sharing one (lane, rank)


@dataclass
class LayoutResult:
    orientation: str                  # "horizontal" | "vertical"
    nodes: Dict[str, NodeLayout]
    edges: List[EdgeLayout]
    lane_order: List[str]
    canvas_width: int
    canvas_height: int
    metrics: LayoutMetrics
    warnings: List[str] = field(default_factory=list)

    # ---- serialization --------------------------------------------------

    def to_dict(self) -> Dict[str, Any]:
        return {
            "orientation": self.orientation,
            "nodes": {nid: asdict(nl) for nid, nl in self.nodes.items()},
            "edges": [asdict(e) for e in self.edges],
            "lane_order": list(self.lane_order),
            "canvas_width": self.canvas_width,
            "canvas_height": self.canvas_height,
            "metrics": asdict(self.metrics),
            "warnings": list(self.warnings),
        }

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "LayoutResult":
        nodes = {nid: NodeLayout(**nl) for nid, nl in d["nodes"].items()}
        edges = [EdgeLayout(**e) for e in d["edges"]]
        metrics = LayoutMetrics(**d["metrics"])
        return cls(
            orientation=d["orientation"],
            nodes=nodes,
            edges=edges,
            lane_order=list(d["lane_order"]),
            canvas_width=d["canvas_width"],
            canvas_height=d["canvas_height"],
            metrics=metrics,
            warnings=list(d.get("warnings", [])),
        )

    # ---- convenience ----------------------------------------------------

    def node(self, node_id: str) -> NodeLayout:
        return self.nodes[node_id]

    def lane_index(self, lane_id: str) -> int:
        return self.lane_order.index(lane_id)
