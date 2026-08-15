# Apply a Corporate PowerPoint Template

How to force an AI-generated deck onto your organisation's template, theme, fonts and colours —
so what comes out looks like it came from your company, not from a default Office install.

## The problem

A generated deck built from blank slides uses whatever theme PowerPoint's default template
supplies. It is legible, but it is obviously not yours: wrong fonts, wrong accent colour, no
logo, no footer.

There are two ways to fix that, and the order matters.

## Option A — apply your template to the finished deck

```
presentation(action: "apply-template", sessionId: ..., templatePath: "C:\\Brand\\corporate.potx")
presentation(action: "get-theme-name", sessionId: ...)
```

`apply-template` calls PowerPoint's own `ApplyTemplate`, which swaps the design — theme fonts,
theme colours, the slide master and its layouts — across the whole presentation. Content stays
where it is; the styling underneath it is replaced.

Accepted template extensions are `.potx`, `.potm`, `.pot`, `.pptx`, `.pptm` and `.ppt`, so an
existing branded deck works as a template just as well as a real `.potx`.

A missing file or an unsupported extension fails cleanly with `success: false` and an explanatory
message rather than throwing, so it is safe to try a path and check the result.

Follow it with `get-theme-name` to confirm the design actually changed — that is the cheap way to
verify the template was applied rather than silently ignored.

**Always re-render afterwards.** A new theme means different fonts at different metrics, and text
that fitted its box under Calibri may overflow under your brand font:

```
export(action: "export-all-slides-to-images", session_id: ..., output_directory: "C:\\preview")
```

This is the single most common reason a branded deck looks broken, and it is invisible unless you
look at the slides.

## Option B — start from the template

Often better: create the presentation *from* a copy of your template file, so every slide added is
already on-brand and you never have to re-flow anything.

```powershell
Copy-Item "C:\Brand\corporate.potx" "C:\Decks\q4-review.pptx"
```

```
presentation(action: "open", filePath: "C:\\Decks\\q4-review.pptx")   → sessionId
```

Then build normally. Because the master and layouts are already correct, `layout(action:
"set-layout", ...)` picks up your branded placeholder scaffolding instead of the Office defaults.

Use Option B when you control the workflow, and Option A when a deck already exists and needs
rebranding.

## Adjusting the master directly

If you do not have a template file, you can set the shared styling yourself. Master-level changes
cascade to every slide that has not overridden the property, which is far cheaper and more
consistent than restyling shapes one by one.

```
master(action: "set-title-font", session_id: ..., font_name: "Inter", font_size: 40, bold: true,
       red: 31, green: 78, blue: 121)
master(action: "set-body-font", session_id: ..., font_name: "Inter", font_size: 18)
master(action: "set-background-color", session_id: ..., red: 255, green: 255, blue: 255)
```

Every field on `set-title-font` and `set-body-font` is optional — omit what you do not want to
change. Pass `red`, `green` and `blue` together when setting a colour.

For a subtler backdrop:

```
master(action: "set-gradient-background", session_id: ...,
       red1: 245, green1: 247, blue1: 250,
       red2: 255, green2: 255, blue2: 255,
       gradient_style: "msoGradientVertical", gradient_variant: 1)
```

Read the current values first with `get-title-font`, `get-body-font` and `get-background-color` if
you want to match an existing deck rather than impose new values.

## Fonts have to be installed

PowerPoint renders with the fonts available on the machine. If your brand font is not installed,
PowerPoint silently substitutes something else and the render — and the saved file on someone
else's machine — will not match your intent.

Two consequences worth planning for:

- Verify the render, do not trust the font name you set. A successful `set-title-font` only means
  the property was assigned.
- If the deck will be opened on machines that may lack the font, either embed fonts in PowerPoint's
  save options or choose a font that ships with Office.

## Per-slide overrides

Master changes do not reach a slide that has already overridden the property. A slide with an
explicit background colour set via `slide(action: "set-background-color", ...)` keeps it, and text
restyled directly with `textframe(action: "set-font-color", ...)` keeps that too.

If a rebrand does not appear to take on some slides, that is usually why. Reset the override at the
slide or shape level, or apply the change there as well.

## Checklist

1. Apply the template, or start from a copy of it.
2. `get-theme-name` to confirm the design changed.
3. Re-render every slide and look for text that now overflows.
4. Fix overflow by shortening text or growing shapes — not by reverting the font.
5. Save.

## Related

- [Build a deck with AI](BUILD-A-DECK-WITH-AI.md)
- [Edit an existing deck](EDIT-AN-EXISTING-DECK.md)
- [Slide master reference](https://powerpointmcpserver.dev/reference/slide-master/)
- [Layouts reference](https://powerpointmcpserver.dev/reference/layouts/)
