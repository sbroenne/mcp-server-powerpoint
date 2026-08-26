---
title: Create and Verify a Presentation
description: Create a PowerPoint deck, add content and speaker notes, check accessibility, render every slide, and save safely on close.
---

# Create and Verify a Presentation

Use this workflow when building a new deck from scratch.

## 1. Create one session

```text
presentation(action="create", filePath="C:\Decks\quarterly-review.pptx")
```

Keep the returned `sessionId` for every later call. Do not open the new file again.

## 2. Build slides in the intended order

For each slide:

1. Add a blank slide with `slide(action="add-blank", ...)`.
2. Add text, shapes, tables, charts, SmartArt, or images.
3. Add speaker notes with `notes(action="set-notes-text", ...)`.
4. Set useful alternative text on visual shapes with `shape(action="set-alt-text", ...)`.

Use `pagesetup(action="get-settings", ...)` before relying on assumed slide dimensions.

## 3. Check structure and accessibility

Run `accessibility(action="audit", ...)` and fix every reported missing alternative text or empty
title placeholder. Use `accessibility(action="get-reading-order", ...)` to inspect the order in
which assistive tools will read each slide.

## 4. Render the real PowerPoint result

```text
export(action="export-all-slides-to-images", session_id="...", output_directory="C:\Decks\preview")
```

Review every rendered slide for overlap, clipped text, poor contrast, and inconsistent spacing.
Fix issues and render again. PowerPoint itself produces these images, so they reflect the saved
presentation more accurately than a file parser can.

## 5. Save and close

```text
presentation(action="close", sessionId="...", save=true)
```

Saving is part of close. There is no separate save action.
