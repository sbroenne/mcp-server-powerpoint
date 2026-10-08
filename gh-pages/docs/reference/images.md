---
description: Insert and adjust PowerPoint pictures, manage linked images, set crop frames and transparency, and compress images.
---

# Images

Reference for the image domain: `image(action: "add-picture", ...)` inserts a picture into a slide.
The image domain provides 16 actions: picture insertion, appearance controls, scalar and crop-frame
cropping, transparency, and image compression.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `image` | `add-picture` | `presentation_session_id`, `slide_index`, `image_path`, `left`, `top`, `width`, `height`, optional `link_to_file`, `save_with_document`, `compression` | Embeds by default. Set `link_to_file=true` for a linked picture. Compression is `default`, `compress`, or `preserve`. |
| `image` | `set-brightness-contrast` | `presentation_session_id`, `slide_index`, `shape_index`, `brightness`, `contrast` | `brightness`/`contrast` are floats in `[0, 1]` (PowerPoint default is `0.5` for both). |
| `image` | `get-brightness-contrast` | `presentation_session_id`, `slide_index`, `shape_index` | Returns current `brightness`/`contrast`. |
| `image` | `set-recolor` | `presentation_session_id`, `slide_index`, `shape_index`, `color_type` | `color_type` is one of `msoPictureAutomatic` (default/no recolor), `msoPictureGrayscale`, `msoPictureBlackAndWhite`, `msoPictureWatermark`. Unrecognized names fail with `Success=false`. |
| `image` | `get-recolor` | `presentation_session_id`, `slide_index`, `shape_index` | Returns current `color_type`. |
| `image` | `set-crop` | `presentation_session_id`, `slide_index`, `shape_index`, `crop_left`, `crop_top`, `crop_right`, `crop_bottom` | Crop edges in points from the picture's edges (L-T-R-B order). Negative values expand the displayed image. Requires a properly sized source image for meaningful geometry. |
| `image` | `get-crop` | `presentation_session_id`, `slide_index`, `shape_index` | Returns current crop offsets in points (L-T-R-B). A fresh picture with no crop applied returns 0.0 for all four values. |
| `image` | `increment-brightness` | `presentation_session_id`, `slide_index`, `shape_index`, `increment` | Changes brightness by a relative amount; PowerPoint clamps the result to `[0, 1]`. |
| `image` | `increment-contrast` | `presentation_session_id`, `slide_index`, `shape_index`, `increment` | Changes contrast by a relative amount; PowerPoint clamps the result to `[0, 1]`. |
| `image` | `set-transparency-color` | `presentation_session_id`, `slide_index`, `shape_index`, `color_rgb` | Sets the color key using PowerPoint's RGB integer order (`0xBBGGRR`). |
| `image` | `get-transparency-color` | `presentation_session_id`, `slide_index`, `shape_index` | Reads the picture's color key. |
| `image` | `set-transparent-background` | `presentation_session_id`, `slide_index`, `shape_index`, `enabled` | Enables or disables color-key transparency; behavior depends on the image format. |
| `image` | `get-transparent-background` | `presentation_session_id`, `slide_index`, `shape_index` | Reads whether color-key transparency is enabled. |
| `image` | `set-crop-frame` | `presentation_session_id`, `slide_index`, `shape_index`, `picture_width`, `picture_height`, `picture_offset_x`, `picture_offset_y`, `frame_left`, `frame_top`, `frame_width`, `frame_height` | Sets the full source-picture and visible-frame geometry in points. |
| `image` | `get-crop-frame` | `presentation_session_id`, `slide_index`, `shape_index` | Reads source-picture size/offsets and visible-frame position/size in points. |
| `image` | `compress-pictures` | `presentation_session_id`, optional `slide_index`, `shape_index`, `resolution`, `delete_cropped_areas` | Compresses one picture, a slide, or the presentation. |

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
| `Crop` | Object | ✓ Exposed | `set-crop-frame`, `get-crop-frame` | Exposes the source picture dimensions, offsets, and visible frame geometry in points. |
| `IncrementBrightness` | Method | ✓ Exposed | `increment-brightness` | Relative adjustment; PowerPoint clamps brightness to `[0, 1]`. |
| `IncrementContrast` | Method | ✓ Exposed | `increment-contrast` | Relative adjustment; PowerPoint clamps contrast to `[0, 1]`. |
| `TransparencyColor` | Property | ✓ Exposed | `set/get-transparency-color` | Uses PowerPoint's RGB integer order (`0xBBGGRR`); results depend on image format and color-key mode. |
| `TransparentBackground` | Property | ✓ Exposed | `set/get-transparent-background` | Enables or disables color-key transparency; defaults vary by image format. |
| `Application` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |
| `Creator` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |
| `Parent` | Property | ⊘ Not exposed | — | Read-only, non-actionable object reference. |

## Picture Compression

`compress-pictures` can target one picture by supplying both indexes, every picture on one slide by
supplying only `slide_index`, or all slides by omitting both. `shape_index` without `slide_index`
is invalid. The resolution presets are `high-fidelity` (do not reduce pixel dimensions), `hd`
(330 PPI), `print` (220 PPI), `web` (150 PPI), and `email` (96 PPI). Set
`delete_cropped_areas=true` to permanently remove pixels hidden by crop settings.

Compression requires a saved `.pptx` or `.pptm` file and Windows. Embedded PNG, JPEG, BMP, GIF,
and TIFF pictures are processed automatically. Linked, vector, animated, and unsupported pictures
are left unchanged and listed in `skippedPictures`. The operation saves and reopens the active
presentation as part of the update; verify important slides visually afterward. Counts and byte
totals in the result describe pictures actually changed, not every selected picture.

## Requirements

- `image_path` must be a **full Windows path** to a local, existing image file (e.g.
  `C:\Assets\logo.png`). There is no URL/remote-fetch parameter — download or generate the image
  to local disk first if it doesn't already exist there.
- The default is embedded: `link_to_file=false`, `save_with_document=true`. The presentation
  remains valid if the original file is moved or deleted.
- A linked-only picture uses `link_to_file=true`, `save_with_document=false` and depends on the
  full source path remaining available. `true/true` keeps the link and also saves picture data in
  the presentation. `false/false` is rejected because PowerPoint would have no picture data.
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
"set-size", ...)` (see `slides-and-shapes.md`), and the appearance, transparency, crop-frame, and
compression actions described above.

## Linked Picture Lifecycle

Use the `shape` tool after adding a linked picture:

1. `shape(action: "get-link-info", ...)` returns `linkSourceFullName`. `linkAutoUpdate` is null
   because some PowerPoint builds reject reads of the typed automatic-update property.
2. `shape(action: "set-link-auto-update", ..., auto_update: false)` switches between automatic
   and manual refresh. If the installed PowerPoint build rejects this typed operation, the MCP or
   CLI boundary returns the PowerPoint error rather than a successful or validation-shaped result.
3. `shape(action: "update-link", ...)` refreshes immediately and fails cleanly if the source file
   is missing.
4. `shape(action: "break-link", ...)` permanently embeds the current image and removes the file
   dependency. The picture shape remains in the slide.

These actions require a linked picture shape. Calling them on an embedded picture or ordinary
shape returns a validation error rather than attempting an invalid COM operation.
