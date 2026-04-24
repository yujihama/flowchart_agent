"""Layout engine: Flow → (node coordinates, edge waypoints, metrics).

The public surface is intentionally small and JSON-first so an autonomous
agent can iterate (compute → inspect metrics → adjust flow → recompute).
"""
from .engine import LayoutOptions, compute_layout
from .result import (
    EdgeLayout,
    LayoutError,
    LayoutMetrics,
    LayoutResult,
    NodeLayout,
)

__all__ = [
    "compute_layout",
    "LayoutOptions",
    "LayoutResult",
    "LayoutMetrics",
    "NodeLayout",
    "EdgeLayout",
    "LayoutError",
]
