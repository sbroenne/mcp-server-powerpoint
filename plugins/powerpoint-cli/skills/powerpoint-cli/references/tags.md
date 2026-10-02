> **CLI syntax:** Shared guides may use MCP calls as shorthand. Use `cli-commands.md` or live `--help` for exact commands and kebab-case options.

# String Tags

PowerPoint string tags attach small text metadata to a presentation, slide, or shape without
changing visible slide content. Each owner supports `set-tag`, `get-tag`, `list-tags`, and
`delete-tag`.

## Name and Value Rules

- Tag names are case-insensitive.
- Letter casing is normalized to invariant uppercase, while whitespace is preserved.
  `reviewState` and `ReviewState` identify the same tag; ` reviewState ` is a distinct name.
- Tag values are never case-normalized or trimmed. A value such as `Needs Review ` round-trips
  exactly as stored.
- `list-tags` returns tags in PowerPoint's native 1-based collection order. Each item includes
  `tagIndex`, `name`, and `value`.
- Missing tags return `success: false`. `delete-tag` is not an idempotent success for a missing
  name.
- Only string tags are supported. Binary tag APIs are intentionally not exposed.

## Presentation Tags

Presentation tags use camelCase arguments on the hand-written `presentation` tool:

```
presentation(action: "set-tag", sessionId: ..., tagName: "ReviewState", tagValue: "Needs Review")
presentation(action: "get-tag", sessionId: ..., tagName: "reviewstate")
presentation(action: "list-tags", sessionId: ...)
presentation(action: "delete-tag", sessionId: ..., tagName: "REVIEWSTATE")
```

## Slide Tags

Slide tags require a 1-based `slide_index`:

```
slide(action: "set-tag", session_id: ..., slide_index: 2,
  tag_name: "Section", tag_value: "Financials")
slide(action: "get-tag", session_id: ..., slide_index: 2, tag_name: "section")
slide(action: "list-tags", session_id: ..., slide_index: 2)
slide(action: "delete-tag", session_id: ..., slide_index: 2, tag_name: "SECTION")
```

## Shape Tags

Shape tags require 1-based `slide_index` and `shape_index` values:

```
shape(action: "set-tag", session_id: ..., slide_index: 2, shape_index: 3,
  tag_name: "DataSource", tag_value: "Quarterly Results")
shape(action: "get-tag", session_id: ..., slide_index: 2, shape_index: 3,
  tag_name: "datasource")
shape(action: "list-tags", session_id: ..., slide_index: 2, shape_index: 3)
shape(action: "delete-tag", session_id: ..., slide_index: 2, shape_index: 3,
  tag_name: "DATASOURCE")
```
