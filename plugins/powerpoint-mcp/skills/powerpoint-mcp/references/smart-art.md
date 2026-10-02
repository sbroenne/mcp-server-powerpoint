# SmartArt

Reference for the `smartart` tool — SmartArt diagrams (process, cycle, hierarchy, list, etc.)
backed by PowerPoint's built-in SmartArt gallery, native shapes and nodes (not images).

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `smartart` | `add-smart-art` | `session_id`, `slide_index`, `layout_name`, `left`, `top`, `width`, `height` | Creates a SmartArt diagram shape using a built-in gallery layout, identified by name (see below). Returns `shapeIndex`, `layoutName`, and `nodeCount` (the layout's default node count). |
| `smartart` | `add-node` | `session_id`, `slide_index`, `shape_index`, `text` | Adds a new top-level (root) node with the given text, appended after the existing top-level nodes. Returns `nodeIndex` and the new `nodeCount`. |
| `smartart` | `add-child-node` | `session_id`, `slide_index`, `shape_index`, `parent_node_index`, `text` | Adds a new node nested one level below `parent_node_index`, appended after that parent's existing children. Returns `nodeIndex` and the new `nodeCount`. |
| `smartart` | `set-node-text` | `session_id`, `slide_index`, `shape_index`, `node_index`, `text` | Sets the text of an existing node. |
| `smartart` | `get-node-text` | `session_id`, `slide_index`, `shape_index`, `node_index` | Returns the text of an existing node. |
| `smartart` | `delete-node` | `session_id`, `slide_index`, `shape_index`, `node_index` | Deletes a node (and any of its children). Returns the new `nodeCount`. |
| `smartart` | `get-node-count` | `session_id`, `slide_index`, `shape_index` | Returns the total number of nodes, flat across the whole hierarchy. |

## Layout Names

`layout_name` is the SmartArt gallery's stable, human-readable display name — not a numeric
position or file path. Common examples:

| Category | Example layout names |
|----------|----------------------|
| Process | `"Basic Process"`, `"Step Up Process"`, `"Continuous Block Process"` |
| Cycle | `"Basic Cycle"`, `"Continuous Cycle"` |
| Hierarchy | `"Organization Chart"`, `"Horizontal Hierarchy"` |
| List | `"Basic Block List"`, `"Vertical Bullet List"` |
| Pyramid | `"Basic Pyramid"` |
| Relationship | `"Basic Venn"`, `"Balance"` |

An unrecognized `layout_name` returns `Success=false` — the exact set of available names depends
on the installed Office version/language, so when in doubt, confirm the name is spelled exactly as
it appears in PowerPoint's own **Insert → SmartArt** gallery (case-insensitive match is tolerated).

## Node Addressing

Nodes are addressed by their **1-based position in the diagram's flat node list**
(`SmartArt.AllNodes`, in document order across the whole hierarchy) — the same order PowerPoint's
own object model already uses, not an invented per-level or path-based scheme. After
`add-smart-art`, use `get-node-count`/`get-node-text` to discover the layout's default node
positions before editing them.

## Building a Diagram

```
smartart(action: "add-smart-art", session_id: ..., slide_index: ...,
  layout_name: "Basic Process", left: 60, top: 120, width: 500, height: 250)
# → shapeIndex, nodeCount (e.g. 3 default placeholder nodes)

smartart(action: "set-node-text", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  node_index: 1, text: "Plan")
smartart(action: "set-node-text", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  node_index: 2, text: "Build")
smartart(action: "set-node-text", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  node_index: 3, text: "Ship")

# Need a 4th step? Append a new top-level node instead of assuming a 4th slot already exists:
smartart(action: "add-node", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  text: "Iterate")
```

## Hierarchy Diagrams (Organization Chart)

`"Organization Chart"` starts with a single root node — build out reports using `add-child-node`
against the root (or an existing child, for a deeper tree):

```
smartart(action: "add-smart-art", session_id: ..., slide_index: ...,
  layout_name: "Organization Chart", left: 60, top: 100, width: 500, height: 350)
# → shapeIndex, nodeCount = 1 (just the root)

smartart(action: "set-node-text", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  node_index: 1, text: "CEO")

smartart(action: "add-child-node", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  parent_node_index: 1, text: "VP Engineering")
smartart(action: "add-child-node", session_id: ..., slide_index: ..., shape_index: <shapeIndex>,
  parent_node_index: 1, text: "VP Sales")
```

## Removing Nodes

`delete-node` removes the target node and any of its descendants — always call `get-node-count`
(or check the `nodeCount` from a prior response) before deleting, and re-resolve node positions
afterward rather than assuming subsequent indices are unchanged (a delete shifts every later
node's flat position down by however many nodes were removed).

## Verify Visually

Diagram layout, spacing, and text overflow are only obvious once rendered — always
`export(action: "export-slide-to-image", ...)` after building or editing a SmartArt diagram to
confirm node text fits and the layout reads as intended (see `export-and-verify.md`).
