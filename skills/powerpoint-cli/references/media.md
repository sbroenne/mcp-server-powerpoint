> **CLI syntax:** Shared guides may use MCP calls as shorthand. Use `cli-commands.md` or live `--help` for exact commands and kebab-case options.

# Audio and Video

Reference for the `media` domain. It inserts playable audio or video as either embedded content or
a file link, then reads PowerPoint's native media type from the resulting shape.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `media` | `add-media` | `session_id`, `slide_index`, `media_path`, `link_to_file`, `save_with_document`, `left`, `top`, `width`, `height` | Adds audio or video through PowerPoint's native media insertion API. |
| `media` | `get-media-info` | `session_id`, `slide_index`, `shape_index` | Returns `mediaTypeName` as `ppMediaTypeSound` or `ppMediaTypeMovie`, plus shape count/index and geometry. |

## Storage Modes

Choose the storage behavior PowerPoint should use:

| Mode | `link_to_file` | `save_with_document` | Source-file behavior |
|------|----------------|----------------------|----------------------|
| Embedded | `false` | `true` | Media is stored in the presentation and survives moving or deleting the source file. |
| Linked only | `true` | `false` | The presentation keeps only a file link. Keep the source path available for reliable playback. |
| Linked and saved | `true` | `true` | The presentation keeps the link and also saves media data, so the media remains available if the source is deleted. |

The `false`/`false` combination is rejected before PowerPoint is called because the presentation
would have neither a link nor saved media data.

## Requirements and Limits

- `media_path` must resolve to an existing local file. Use a full Windows path.
- `slide_index` and `shape_index` are 1-based.
- `width` and `height` must be greater than zero and are measured in points.
- Prefer formats supported natively by the target PowerPoint installation. H.264 video with AAC
  audio in an MP4 container is Microsoft's recommended general-purpose format.
- This first media surface covers insertion and metadata only. It does not control playback,
  trimming, timing, volume, poster frames, or slide-show behavior.

## Verification

Use `get-media-info` after insertion to confirm that PowerPoint classified the shape as sound or
movie. Exporting the slide can verify placement and the video poster frame, but a rendered image
cannot prove audio or video playback.
