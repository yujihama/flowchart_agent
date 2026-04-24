from __future__ import annotations

from ..model import Flow, Node

_SHAPE = {
    "start":      ('(["', '"])'),   # stadium
    "end":        ('(["', '"])'),   # stadium
    "decision":   ('{"',   '"}'),   # diamond
    "task":       ('["',   '"]'),   # rectangle
    "subprocess": ('[["',  '"]]'),  # subroutine
    "document":   ('[("',  '")]'),  # cylinder
}


def _shape(node: Node) -> str:
    open_, close_ = _SHAPE.get(node.type, _SHAPE["task"])
    return f"{node.id}{open_}{node.label}{close_}"


def to_mermaid(flow: Flow, *, direction: str = "TD") -> str:
    """Render a Flow as Mermaid `flowchart` text.

    Swimlanes are expressed as `subgraph` blocks — an approximation of true
    swimlane semantics but widely supported (GitHub, Notion, VS Code, etc.).
    """
    lines: list[str] = [f"flowchart {direction}"]

    nodes_by_lane: dict[str, list[Node]] = {}
    for n in flow.nodes:
        nodes_by_lane.setdefault(n.lane_id, []).append(n)

    for lane in flow.lanes:
        lines.append(f'  subgraph {lane.id}["{lane.name}"]')
        for n in nodes_by_lane.get(lane.id, []):
            lines.append(f"    {_shape(n)}")
        lines.append("  end")

    for e in flow.edges:
        if e.condition:
            lines.append(f"  {e.from_id} -->|{e.condition}| {e.to_id}")
        else:
            lines.append(f"  {e.from_id} --> {e.to_id}")

    return "\n".join(lines)
