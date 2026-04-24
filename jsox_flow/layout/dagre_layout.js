#!/usr/bin/env node
// Dagre layout driver.
//
// Reads a JSON spec from stdin, runs dagre's layered layout, and writes the
// result to stdout as JSON. The Python side is authoritative for lane
// placement: this script only computes ranks (step positions along the
// flow axis) and waypoints. Lane axis coordinates from dagre are kept in
// the output for informational use but Python overrides them.
//
// Input:
//   {
//     "orientation": "horizontal" | "vertical",
//     "nodes": [{"id": "n1", "width": 130, "height": 46, "label": "..."}, ...],
//     "edges": [{"from": "n1", "to": "n2"}, ...],
//     "options": {             // all optional
//       "nodesep": 30,
//       "edgesep": 15,
//       "ranksep": 60,
//       "ranker": "network-simplex" | "tight-tree" | "longest-path"
//     }
//   }
//
// Output:
//   {
//     "nodes": {
//       "n1": {"x": 65, "y": 23, "width": 130, "height": 46}
//     },
//     "edges": [
//       {"from": "n1", "to": "n2", "points": [[x,y], ...]}
//     ],
//     "canvas": {"width": ..., "height": ...}
//   }
"use strict";

const dagre = require("dagre");

function readStdin() {
  return new Promise((resolve, reject) => {
    let buf = "";
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (c) => (buf += c));
    process.stdin.on("end", () => resolve(buf));
    process.stdin.on("error", reject);
  });
}

function round(n) {
  return Math.round(n);
}

async function main() {
  const raw = await readStdin();
  let input;
  try {
    input = JSON.parse(raw);
  } catch (e) {
    process.stderr.write("invalid JSON on stdin: " + e.message + "\n");
    process.exit(2);
  }

  const orientation = input.orientation === "vertical" ? "vertical" : "horizontal";
  const rankdir = orientation === "vertical" ? "TB" : "LR";
  const opts = input.options || {};

  const g = new dagre.graphlib.Graph({ multigraph: false, compound: false });
  g.setGraph({
    rankdir: rankdir,
    nodesep: opts.nodesep != null ? opts.nodesep : 30,
    edgesep: opts.edgesep != null ? opts.edgesep : 15,
    ranksep: opts.ranksep != null ? opts.ranksep : 60,
    marginx: opts.marginx != null ? opts.marginx : 20,
    marginy: opts.marginy != null ? opts.marginy : 20,
    ranker: opts.ranker || "network-simplex",
  });
  g.setDefaultEdgeLabel(() => ({}));

  const nodes = Array.isArray(input.nodes) ? input.nodes : [];
  const edges = Array.isArray(input.edges) ? input.edges : [];

  for (const n of nodes) {
    if (!n || typeof n.id !== "string") {
      process.stderr.write("invalid node entry\n");
      process.exit(2);
    }
    g.setNode(n.id, {
      width: n.width || 130,
      height: n.height || 46,
      label: n.label || n.id,
    });
  }
  for (const e of edges) {
    if (!e || typeof e.from !== "string" || typeof e.to !== "string") {
      process.stderr.write("invalid edge entry\n");
      process.exit(2);
    }
    g.setEdge(e.from, e.to, {});
  }

  try {
    dagre.layout(g);
  } catch (e) {
    process.stderr.write("dagre.layout failed: " + e.message + "\n");
    process.exit(3);
  }

  const outNodes = {};
  for (const id of g.nodes()) {
    const n = g.node(id);
    outNodes[id] = {
      x: round(n.x),
      y: round(n.y),
      width: round(n.width),
      height: round(n.height),
    };
  }

  const outEdges = [];
  for (const e of g.edges()) {
    const ed = g.edge(e);
    const pts = (ed.points || []).map((p) => [round(p.x), round(p.y)]);
    outEdges.push({ from: e.v, to: e.w, points: pts });
  }

  const gp = g.graph();
  const result = {
    nodes: outNodes,
    edges: outEdges,
    canvas: { width: round(gp.width || 0), height: round(gp.height || 0) },
  };

  process.stdout.write(JSON.stringify(result));
}

main().catch((e) => {
  process.stderr.write("unexpected: " + (e && e.stack ? e.stack : e) + "\n");
  process.exit(1);
});
