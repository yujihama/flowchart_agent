from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Literal, Optional

NodeType = Literal["start", "end", "task", "decision", "subprocess", "document"]


@dataclass
class Lane:
    id: str
    name: str


@dataclass
class Node:
    id: str
    lane_id: str
    label: str
    type: NodeType = "task"


@dataclass
class Edge:
    from_id: str
    to_id: str
    condition: Optional[str] = None


@dataclass
class Flow:
    lanes: List[Lane] = field(default_factory=list)
    nodes: List[Node] = field(default_factory=list)
    edges: List[Edge] = field(default_factory=list)

    def get_node(self, node_id: str) -> Optional[Node]:
        return next((n for n in self.nodes if n.id == node_id), None)

    def get_lane(self, lane_id: str) -> Optional[Lane]:
        return next((l for l in self.lanes if l.id == lane_id), None)

    def outgoing(self, node_id: str) -> List[Edge]:
        return [e for e in self.edges if e.from_id == node_id]

    def incoming(self, node_id: str) -> List[Edge]:
        return [e for e in self.edges if e.to_id == node_id]
