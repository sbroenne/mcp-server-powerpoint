# Images

Reference for the image domain: `image(action: "add-picture", ...)` inserts a picture into a slide.
The image domain provides 7 total actions: 1 insertion action (`add-picture`) plus 6 appearance actions: `set-brightness-contrast`, `get-brightness-contrast`,
`set-recolor`, `get-recolor`, `set-crop`, and `get-crop`.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `image` | `add-picture` | `session_id`, `slide_index`, `image_path`, `left`, `top`, `width`, `height` | Embeds (not links) the file into the presentation. |
| `image` | `set-brightness-contrast` | `session_id`, `slide_index`, `shape_index`, `brightness`, `contrast` | `brightness`/`contrast` are floats in `[0, 1]` (PowerPoint default is `0.5` for both). |
| `image` | `get-brightness-contrast` | `session_id`, `slide_index`, `shape_index` | Returns current `brightness`/`contrast`. |
| `image` | `set-recolor` | `session_id`, `slide_index`, `shape_index`, `color_type` | `color_type` is one of `msoPictureAutomatic` (default/no recolor), `msoPictureGrayscale`, `msoPictureBlackAndWhite`, `msoPictureWatermark`. Unrecognized names fail with `Success=false`. |
| `image` | `get-recolor` | `session_id`, `slide_index`, `shape_index` | Returns current `color_type`. |
| `image` | `set-crop` | `session_id`, `slide_index`, `shape_index`, `crop_left`, `crop_top`, `crop_right`, `crop_bottom` | Crop edges in points from the picture's edges (L-T-R-B order). Negative values expand the displayed image. Requires a properly sized source image for meaningful geometry. |
| `image` | `get-crop` | `session_id`, `slide_index`, `shape_index` | Returns current crop offsets in points (L-T-R-B). A fresh picture with no crop applied returns 0.0 for all four values. |

## Crop Behavior

When using `set-crop`, all four values are relative crop distances in points from the picture's
edges. A positive value crops inward (hides content); a negative value expands the visible area
outward (revealing areas beyond the original picture bounds in PowerPoint's rendering). PowerPoint
allows negative crop values as a valid expansion mechanism.

**Important:** Meaningful crop geometry requires a properly sized source image. The legacy test anomaly
of a 1×1 pixel image should **never** be treated as normal user behavior — it is a degenerate case
used only for testing framework integration and produces unintuitive crop results. Always use real,
properly sized source images (e.g., 100×100 pixels minimum) when working with crop operations.

Results from `get-crop` expose `cropLeft`, `cropTop`, `cropRight`, `cropBottom` as numeric
floats (in points). A fresh picture with no crop applied returns 0.0 for all four values.
These result fields are absent (null in JSON) only when the result comes from a non-crop operation
(e.g., `get-recolor`) — they are never null simply because no crop has been applied.

## PictureFormat Coverage

The image domain exposes these `Microsoft.Office.Interop.PowerPoint.PictureFormat` properties and methods:

| Member | Type | Status | Action | Notes |
|--------|------|--------|--------|-------|
| `Brightness` | Property | ✓ Exposed | `set-brightness-contrast`, `get-brightness-contrast` | Float [0, 1] scale. Default 0.5. |
| `Contrast` | Property | ✓ Exposed | `set-brightness-contrast`, `get-brightness-contrast` | Float [0, 1] scale. Default 0.5. |
| `ColorType` | Property | ✓ Exposed | `set-recolor`, `get-recolor` | MsoPictureColorType enum. |
| `CropLeft` | Property | ✓ Exposed | `set-crop`, `get-crop` | Direct scalar property, in points from left edge. |
| `CropTop` | Property | ✓ Exposed | `set-crop`, `get-crop` | Direct scalar property, in points from top edge. |
| `CropRight` | Property | ✓ Exposed | `set-crop`, `get-crop` | Direct scalar property, in points from right edge. |
| `CropBottom` | Property | ✓ Exposed | `set-crop`, `get-crop` | Direct scalar property, in points from bottom edge. |
| `Crop` | Object | ⊘ Not exposed | — | Separate Office.Core Crop subobject with natural-image/display scaling semantics. Deliberately unexposed; direct scalar crop properties subsume its interface. |
| `IncrementBrightness` | Method | ⊘ Not exposed | — | Already covered by idempotent absolute setter; live COM confirms increments clamp to [0, 1]. |
| `IncrementContrast` | Method | ⊘ Not exposed | — | Already covered by idempotent absolute setter; live COM confirms increments clamp to [0, 1]. |
| `TransparencyColor` | Property | ⊘ Not exposed | — | Live COM proved set/read/persistence on BMP, but deliberately unexposed: BMP defaults to msoTrue, no-key uses int.MinValue sentinel, exact color-key knowledge required, JPEG unsuitable, PNG supports alpha natively. |
| `TransparentBackground` | Property | ⊘ Not exposed | — | Live COM proved set/read/persistence on BMP, but deliberately unexposed: defaults are format-dependent (BMP: msoTrue), exact behavior requires format-specific tuning. |
| `Application` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |
| `Creator` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |
| `Parent` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |

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

## Limited Editing After Insert

Post-insert adjustments available: `shape(action: "set-position", ...)`, `shape(action:
"set-size", ...)` (see `slides-and-shapes.md`), and `set-brightness-contrast`, `set-recolor`,
and `set-crop` actions (see above).
