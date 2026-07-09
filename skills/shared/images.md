# Images

Reference for `image(action: "add-picture", ...)` — embeds a local image file into a slide.

## Action

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `image` | `add-picture` | `session_id`, `slide_index`, `image_path`, `left`, `top`, `width`, `height` | Embeds (not links) the file into the presentation. |

## Requirements

- `image_path` must be a **full Windows path** to a local, existing image file (e.g.
  `C:\Assets\logo.png`). There is no URL/remote-fetch parameter — download or generate the image
  to local disk first if it doesn't already exist there.
- The image is **embedded**, not linked: the presentation's file size grows by the image size, and
  the presentation remains valid even if the original file is later moved or deleted.
- `width`/`height` are explicit — this action does not auto-detect or preserve the source image's
  native aspect ratio. If the aspect ratio matters (logos, photos), compute `width`/`height` to
  match the source image's ratio yourself before calling, or the image will appear stretched or
  squashed.

## Placement Guidance

- Logos: small, corner-anchored, e.g. `width=80, height=80` near `left=20, top=20` (top-left) or
  `left=860, top=460` (bottom-right on a 960×540pt slide).
- Full-slide product shots/photos: leave a title band at the top (`top ≥ 80`) unless the image is
  intentionally full-bleed.
- Diagrams/screenshots next to explanatory text: place the image in one half of the slide
  (`width ≈ 420` on a 960pt-wide slide) with a text box in the other half.

## Verify After Adding

Always `export(action: "export-slide-to-image", ...)` after `image(action: "add-picture", ...)`
to confirm: the image loaded (not a broken placeholder), the aspect ratio looks correct, and it
doesn't overlap other shapes on the slide (see `export-and-verify.md`).

## No Editing After Insert

There is no crop/rotate/effects action in this surface. If the image needs cropping or rotation,
prepare the final image file before calling `add-picture` — the only post-insert adjustments
available are `shape(action: "set-position", ...)` and `shape(action: "set-size", ...)` (see
`slides-and-shapes.md`).
