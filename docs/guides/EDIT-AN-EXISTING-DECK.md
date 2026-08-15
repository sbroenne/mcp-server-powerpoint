# Edit an Existing PowerPoint Deck

How to update, restructure or repair a `.pptx` file you already have — without rebuilding it and
without losing the formatting someone spent hours on.

## Why this is different from generating a deck

Editing an existing deck means working blind unless you inspect first. The file already has a
theme, a master, custom layouts, and shapes placed by a human. The goal is the smallest possible
change that achieves what was asked, leaving everything else untouched.

The failure mode to avoid is delete-and-recreate: rebuilding a slide from scratch because it was
easier than reading it. That discards the original formatting and is almost never what the person
asking wanted.

## Step 1 — Open and survey

```
presentation(action: "open", filePath: "C:\\Decks\\board-deck.pptx")   → sessionId
slide(action: "get-count", session_id: ...)                            → how many slides
presentation(action: "get-theme-name", sessionId: ...)                 → which theme is in use
```

Then, for any slide you intend to touch:

```
shape(action: "get-count", session_id: ..., slide_index: 4)
textframe(action: "get-text", session_id: ..., slide_index: 4, shape_index: 1)
shape(action: "get-name", session_id: ..., slide_index: 4, shape_index: 1)
```

Reading shape names is the reliable way to find the shape you actually mean. Guessing that "the
title is always shape 1" is wrong often enough to matter, especially in decks built from custom
layouts.

If the deck is unfamiliar, the fastest survey is to render it:

```
export(action: "export-all-slides-to-images", session_id: ..., output_directory: "C:\\Decks\\before")
```

One pass gives an assistant the whole deck visually, which is far more useful than reading
shape metadata slide by slide.

## Step 2 — Make targeted changes

Prefer the narrowest operation that does the job:

| Goal | Use | Not |
|------|-----|-----|
| Change wording | `textframe(action: "set-text", ...)` | Deleting the text box and adding a new one |
| Update a number in a table | `table(action: "set-cell-text", ...)` | Re-adding the table |
| Refresh chart data | `chart(action: "replace-chart-data", ...)` | Deleting and recreating the chart |
| Nudge a shape | `shape(action: "set-position"/"set-size", ...)` | Recreating the shape at new coordinates |
| Restyle text | `textframe(action: "set-font-size"/"set-bold"/"set-font-color", ...)` | Rewriting the whole slide |

Targeted edits inherit everything the original had — theme colors, autofit behaviour, animation
attached to the shape. A recreated shape does not.

## Step 3 — Restructure carefully

Structural operations renumber things, which is where index bugs come from.

```
slide(action: "duplicate", session_id: ..., slide_index: 3)
slide(action: "move-to", session_id: ..., slide_index: 7, to_position: 2)
slide(action: "delete", session_id: ..., slide_index: 5)
```

- `duplicate` inserts the copy directly after the source, so everything below shifts down by one.
- `move-to` renumbers every slide between the old and new position. It returns the slide's new
  `slideIndex` — use that value rather than the one you passed in.
- `delete` shifts everything after it up by one.

**Re-read `slide(action: "get-count", ...)` after any structural change**, and never hold a list
of slide indexes across one. If you need to delete several slides, delete from the highest index
downwards so the earlier indexes stay valid.

To insert a slide in the middle, append it with `add-blank` and then `move-to` the position you
want.

## Step 4 — Verify the slides you touched

```
export(action: "export-slide-to-image", session_id: ..., slide_index: 4, output_path: "C:\\Decks\\after-4.png")
```

Export only the slides you changed while iterating — rendering the whole deck on every fix cycle
is wasteful once you have localized the problem. Do a full
`export-all-slides-to-images` pass at the end to confirm nothing shifted elsewhere.

Comparing the before and after renders is the fastest way to prove an edit did exactly what was
intended and nothing more.

## Step 5 — Save

```
presentation(action: "save", sessionId: ...)
presentation(action: "close", sessionId: ...)
```

If the original matters, work on a copy. There is no undo across a save.

## Common tasks

**Fix a typo across the deck.** Loop slides, `textframe(action: "get-text", ...)` on each shape,
and `set-text` only where the string actually appears. Do not blanket-rewrite.

**Rebrand fonts.** Change them at the master level with
`master(action: "set-title-font"/"set-body-font", ...)` so every slide inherits the change,
instead of restyling shapes one at a time. See
[Apply a corporate template](APPLY-A-CORPORATE-TEMPLATE.md).

**Update last quarter's numbers.** `chart(action: "replace-chart-data", ...)` swaps categories,
series names and values in a single call, including changing the category count.

**Add presenter notes to a deck that has none.** `notes(action: "set-notes-text", ...)` per slide.
An assistant that has just read the whole deck visually is well placed to write them.

## Related

- [Build a deck with AI](BUILD-A-DECK-WITH-AI.md)
- [Automate charts and tables](AUTOMATE-CHARTS-AND-TABLES.md)
- [Slides and shapes reference](https://powerpointmcpserver.dev/reference/slides-and-shapes/)
- [Anti-patterns to avoid](https://powerpointmcpserver.dev/reference/anti-patterns/)
