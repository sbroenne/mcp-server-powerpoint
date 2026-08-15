# Export PowerPoint Slides to Images

How to render slides to PNG (or JPG, GIF, BMP, TIF, WMF, EMF) — and why this is the single most
important operation when an AI assistant is producing slides.

## The two actions

| Action | Renders | Parameters |
|--------|---------|------------|
| `export(action: "export-slide-to-image", ...)` | One slide | `session_id`, `slide_index`, `output_path`, `format` (default `PNG`), optional `width`/`height` in pixels |
| `export(action: "export-all-slides-to-images", ...)` | Every slide | `session_id`, `output_directory`, `format` (default `PNG`) |

`export-all-slides-to-images` creates the output directory if it does not exist and lets
PowerPoint name the files `Slide1.PNG`, `Slide2.PNG`, and so on.

From the CLI:

```powershell
pptcli session open "C:\Decks\q4-review.pptx"
# → SESSION_ID

pptcli export export-slide-to-image -s SESSION_ID --slide-index 3 --output-path "C:\preview\slide3.png"
pptcli export export-all-slides-to-images -s SESSION_ID --output-directory "C:\preview"
```

## Why it matters more than it sounds

Rendering is what makes an AI assistant's PowerPoint output trustworthy.

Text-only inspection — reading back shape counts, cell values and chart dimensions — cannot detect
the things that actually make a deck unusable:

- A title that overflows its placeholder and gets clipped.
- A chart sitting on top of the footer.
- Two text boxes overlapping by 20 points.
- A font that was substituted because it is not installed.
- A stretched logo with the wrong aspect ratio.

Every one of those coexists happily with `success: true` on every preceding call. A successful COM
call confirms PowerPoint accepted the parameters, not that the slide looks right.

Offline `.pptx` libraries have no renderer at all, which is why an assistant driving them is
working blind. Here, PowerPoint renders the preview with the same engine that will open the file
later, so what the assistant sees is what the audience gets.

## The verification loop

```
1. Add or move visual content (chart, table, image, shape)
2. export(action: "export-slide-to-image", ...) on that slide
3. Inspect the image
4. If anything is wrong → fix → export again
5. Only move on when it looks right
```

Use the single-slide export while iterating. Once the deck is complete, run
`export-all-slides-to-images` once as a final pass to confirm no slide was missed and the deck
reads coherently from start to finish.

Budget one or two fix cycles for any slide with a chart, a table or several overlapping shapes.
That is the expected cost of getting it right, not a sign something broke.

## Choosing a format and size

- **PNG** is the default and the right choice for feeding images back to a multimodal model —
  lossless, sharp text, no compression artifacts around glyph edges.
- **JPG** only if you need small files for a web preview, and accept the text fringing.
- **WMF/EMF** are vector formats, useful if you need to scale the output, but not for visual
  inspection by a model.

`width` and `height` on `export-slide-to-image` set the output pixel dimensions. Omit them for
PowerPoint's default rendering size, which is usually adequate. Raise them if you need to read
small chart axis labels in the render.

## Other uses for slide rendering

Visual verification is the main one, but the same call covers:

- **Thumbnails** for a document management system or an internal deck library.
- **Deck previews** embedded in a web page or a chat message.
- **Visual regression checks** — render a deck before and after a template change and diff the
  images to see exactly what moved.
- **Handout images** for documentation, wikis or release notes.

## Requirements and limits

Rendering uses PowerPoint's own export, so the PowerPoint desktop application must be installed
and a desktop session must be available. This is not a headless operation — it does not work on a
bare CI runner without an interactive session.

If you need slide rendering with no Office installed, that is a genuine reason to look at an
offline library instead; see the
[comparison page](https://powerpointmcpserver.dev/comparison/) for the trade-offs.

## Related

- [Build a deck with AI](BUILD-A-DECK-WITH-AI.md)
- [Edit an existing deck](EDIT-AN-EXISTING-DECK.md)
- [Export and visual verification reference](https://powerpointmcpserver.dev/reference/export-and-verify/)
