"""YAML (de)serialization for Flow.

The YAML schema intentionally uses short, human-friendly keys:

* Node ``lane_id`` → ``lane``
* Edge ``from_id`` / ``to_id`` → ``from`` / ``to``
* Node ``type`` defaults to ``task`` and is omitted when not overridden
* Edge ``condition`` is omitted when absent
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Union

import yaml

from .model import Edge, Flow, Lane, Node


def load_flow(path: Union[str, Path]) -> Flow:
    """Read a Flow from a YAML file on disk."""
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    return from_dict(data)


def from_yaml(text: str) -> Flow:
    """Parse a Flow from a YAML string."""
    return from_dict(yaml.safe_load(text))


def dump_flow(flow: Flow, path: Union[str, Path]) -> Path:
    """Write a Flow to a YAML file on disk."""
    p = Path(path)
    p.write_text(to_yaml(flow), encoding="utf-8")
    return p


def to_yaml(flow: Flow) -> str:
    """Render a Flow as YAML text."""
    return yaml.safe_dump(
        to_dict(flow),
        default_flow_style=False,
        sort_keys=False,
        allow_unicode=True,
        width=120,
    )


def from_dict(data: Dict[str, Any]) -> Flow:
    if not isinstance(data, dict):
        raise ValueError("YAML root must be a mapping with lanes/nodes/edges.")

    lanes = [
        Lane(id=_require(d, "id"), name=_require(d, "name"))
        for d in data.get("lanes") or []
    ]
    nodes = [
        Node(
            id=_require(d, "id"),
            lane_id=_require(d, "lane"),
            label=_require(d, "label"),
            type=d.get("type", "task"),
        )
        for d in data.get("nodes") or []
    ]
    edges = [
        Edge(
            from_id=_require(d, "from"),
            to_id=_require(d, "to"),
            condition=d.get("condition"),
        )
        for d in data.get("edges") or []
    ]
    return Flow(lanes=lanes, nodes=nodes, edges=edges)


def to_dict(flow: Flow) -> Dict[str, Any]:
    lanes = [{"id": l.id, "name": l.name} for l in flow.lanes]

    nodes = []
    for n in flow.nodes:
        d: Dict[str, Any] = {"id": n.id, "lane": n.lane_id, "label": n.label}
        if n.type != "task":
            d["type"] = n.type
        nodes.append(d)

    edges = []
    for e in flow.edges:
        d = {"from": e.from_id, "to": e.to_id}
        if e.condition is not None:
            d["condition"] = e.condition
        edges.append(d)

    return {"lanes": lanes, "nodes": nodes, "edges": edges}


def _require(d: Dict[str, Any], key: str) -> Any:
    if key not in d:
        raise ValueError(f"Missing required field '{key}' in: {d!r}")
    return d[key]
