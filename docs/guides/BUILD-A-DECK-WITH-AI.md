# Build a PowerPoint Deck with AI

How to have an AI assistant generate a complete, presentable multi-slide deck — and why the
result looks different from what a prompt-to-slides web tool produces.

## The short version

```
1. presentation(action: "create", filePath: "C:\Decks\q4-review.pptx")   → sessionId
2. For each planned slide:
     slide(action: "add-blank", session_id: ...)
     layout(action: "set-layout", session_id: ..., slide_index: N, layout_name: ...)
     shape / table / chart / image calls to place content
     notes(action: "set-notes-text", session_id: ..., slide_index: N, text: ...)
3. export(action: "export-all-slides-to-images", session_id: ..., output_directory: ...)
4. Review every image, fix what looks wrong, re-export
5. presentation(action: "save", sessionId: ...)
6. presentation(action: "close", sessionId: ...)
```

Steps 3 and 4 are the part most tools skip, and the reason this approach produces decks you can
actually show to someone.

## Step 1 — Plan the whole deck first

New slides are appended at the end. You can reorder later with
`slide(action: "move-to", ...)`, but every move renumbers the slides around it and invalidates
indexes you were holding, so it is cheaper to decide the running order up front.

Ask the assistant to produce the outline *before* it touches PowerPoint — one line per slide,
naming the purpose and the composition:

```
1. Title            — deck title + subtitle, no other content
2. Agenda           — 4 bullets
3. Section divider  — large centered heading
4. Revenue trend    — title + line chart
5. Regional split   — title + 3-column table
6. Summary          — 3 bullets + closing statement
```

A good outline varies the composition. If slides 2 through 6 are all "title plus one bulleted
text box", the deck reads as generated rather than designed.

## Step 2 — Create the presentation once

```
presentation(action: "create", filePath: "C:\\Decks\\q4-review.pptx")
```

This creates the file, saves it, and leaves the session **open**. Reuse the returned `sessionId`
for everything that follows — do not call `open` on the same path afterwards.

From the CLI the equivalent is:

```powershell
pptcli session create "C:\Decks\q4-review.pptx"
# → prints a session id; pass it as -s <SESSION> to every later command
```

The CLI keeps the session alive in a background daemon between invocations, so you pay
PowerPoint's startup cost once rather than on every command.

## Step 3 — Build each slide

`slide(action: "add-blank", ...)` always appends a **blank** slide. Applying a layout with
`layout(action: "set-layout", ...)` sets up PowerPoint's placeholder scaffolding, but the actual
content is placed explicitly with the content tools:

| Content | Call |
|---------|------|
| Title or body text | `shape(action: "add-text-box", ...)` then `textframe(action: "set-text"/"set-font-size"/"set-bold", ...)` |
| Structured data | `table(action: "add-table", ...)` then `table(action: "set-cell-text", ...)` |
| Metrics and trends | `chart(action: "add-chart", ...)` — see the [charts and tables guide](AUTOMATE-CHARTS-AND-TABLES.md) |
| Screenshots, logos, diagrams | `image(action: "add-picture", ...)` |
| Presenter script | `notes(action: "set-notes-text", ...)` |

### Positioning

`left`, `top`, `width` and `height` are in **points** (1 inch = 72 points). A standard 16:9 slide
is roughly 960 × 540 points, so:

- Keep `left >= 40` and `top >= 40`.
- Keep `left + width <= 920` and `top + height <= 500`.
- Title band: `top: 30, height: 80`. Body area starts around `top: 130`.
- Two columns: `width: 420` each with a 40-point gutter.

There is no API to query the actual slide dimensions, so assume the 16:9 default and confirm
visually rather than trusting the numbers.

### Always write speaker notes

`notes(action: "set-notes-text", ...)` is cheap and it is what turns a deck into something
presentable. It is also the part an AI assistant is genuinely good at, since it has the full
context of why each slide exists.

## Step 4 — Look at what you built

This is the step that matters:

```
export(action: "export-all-slides-to-images", session_id: ..., output_directory: "C:\\Decks\\preview")
```

PowerPoint renders every slide to `Slide1.PNG`, `Slide2.PNG` and so on. Feed those images back to
the assistant. A `success: true` from a COM call only means PowerPoint accepted the parameters —
it says nothing about whether the title overflows its box or the chart covers the footer.

Typical issues and their fixes:

| What you see | Fix |
|--------------|-----|
| Text cut off at the box edge | Shorten the text, lower `font_size`, or grow the shape with `shape(action: "set-size", ...)` |
| Two shapes overlapping | `shape(action: "set-position", ...)` on one of them |
| Chart too small to read | Increase the chart's `width`/`height`, or reduce the category count |
| Image stretched | Recompute `width`/`height` to match the source aspect ratio |

Expect one or two fix cycles on visually dense slides. That is normal, not a failure.

## Step 5 — Save and close

```
presentation(action: "save", sessionId: ...)
presentation(action: "close", sessionId: ...)
```

`close` returns immediately and disposes the PowerPoint process in the background. Office's own
cleanup can take a couple of minutes afterwards — that is expected behaviour and not something to
work around by killing the process.

## Why this beats prompt-to-slides tools

An offline library or a hosted slide generator writes the file and stops. It has no renderer, so
it cannot tell the difference between a deck that is correct and a deck that merely parses. Here
PowerPoint itself writes the file and PowerPoint itself renders the preview, so the assistant is
checking its work against exactly what a viewer will open.

See the [comparison page](https://powerpointmcpserver.dev/comparison/) for how this stacks up
against `python-pptx`, the Open XML SDK and VBA.

## Related

- [Automate charts and tables](AUTOMATE-CHARTS-AND-TABLES.md)
- [Export slides to images](EXPORT-SLIDES-TO-IMAGES.md)
- [Apply a corporate template](APPLY-A-CORPORATE-TEMPLATE.md)
- [Deck builder reference](https://powerpointmcpserver.dev/reference/deck-builder/)
