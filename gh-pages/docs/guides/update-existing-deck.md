---
title: Update an Existing Presentation Safely
description: Inspect an existing PowerPoint deck before editing it, make targeted changes, render affected slides, and choose whether to save on close.
---

# Update an Existing Presentation Safely

Use this workflow when changing a presentation you did not create in the current session.

## 1. Test before opening

```text
presentation(action="test", filePath="C:\Decks\existing.pptx")
```

`test` checks whether PowerPoint can open the file without keeping a live session. If it succeeds,
open the file once and keep the returned session ID.

## 2. Inspect before editing

Start with discovery operations:

- `slide(action="get-count", ...)` for the slide range.
- `shape(action="get-count", ...)` for each slide you will change.
- `shape(action="get-name", ...)` and `shape(action="get-alt-text", ...)` to identify content.
- `textframe(action="get-text", ...)` before replacing text.
- `layout(action="get-layout", ...)` and `pagesetup(action="get-settings", ...)` before changing
  layout or size.

Reading first avoids editing the wrong 1-based index.

## 3. Make focused changes

Change only the slides and shapes needed for the task. If a change is uncertain, render that slide
before continuing rather than making the same guess across the whole deck.

## 4. Verify and close

Render every changed slide with `export(action="export-slide-to-image", ...)`. Run the
accessibility audit when content or reading order changed.

Close with `save=true` to keep the changes or `save=false` to discard them:

```text
presentation(action="close", sessionId="...", save=true)
```
