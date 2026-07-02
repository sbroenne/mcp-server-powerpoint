# Speaker Notes

Reference for `set_notes_text` and `get_notes_text`.

## Tools

| Tool | Parameters | Notes |
|------|------------|-------|
| `set_notes_text` | `sessionId`, `slideIndex`, `text` | Replaces the slide's entire notes content. |
| `get_notes_text` | `sessionId`, `slideIndex` | Reads current notes — use before editing to avoid discarding existing content. |

## When to Add Notes

Add speaker notes to every content slide you create or substantially modify — not just when
explicitly asked:

- After `add_slide` + building its content, before moving to the next slide (see
  `deck-builder.md`'s slide assembly loop).
- After significantly changing a slide's message (new data, restructured content).

## Note Structure

Keep each note focused and short (well under 150 words):

1. **One-sentence hook** — what this slide is about / why it matters.
2. **2-4 talking points** — context not visible on the slide itself (the "why" behind a stat, a
   data source, an anticipated question).
3. **One-sentence transition** to the next slide's topic.

## Writing Style

- Plain text only — `set_notes_text` takes a plain string, no markdown/bold/bullet syntax is
  rendered by PowerPoint's Notes pane. Use line breaks (`\n`) and leading dashes for structure.
- Write what a presenter would *say*, not a repeat of the slide's own text. If the slide says
  "Revenue grew 24%", the note should explain *why* or *so what*, not restate the number.
- Address the audience conversationally where useful ("You'll notice…", "This matters because…").

## Example

```
set_notes_text(sessionId, slideIndex=2,
  text="Revenue accelerated through the year, with Q4 the strongest quarter on record.\n" +
       "- Growth was broad-based across all three regions, not one outlier.\n" +
       "- Source: Finance close, Q4 2025.\n" +
       "- Next slide breaks this down by region.")
```

## Don't Skip Notes

Every slide deserves notes, even simple section dividers ("Pause here before moving into the
financials section."). Missing notes on a deck built for someone else to present is an incomplete
deliverable.

## Read Before Overwrite

`set_notes_text` replaces the whole notes body. If you need to append to existing notes rather
than replace them, call `get_notes_text` first, compose the full new string, then call
`set_notes_text` with the complete result — there is no append operation.
