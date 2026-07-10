---
"@sbroenne/mcp-server-powerpoint": minor
---

Add picture-effect actions to the `image` domain: `set-brightness-contrast`/`get-brightness-contrast`
(adjust brightness and contrast, each a float in `[0, 1]`) and `set-recolor`/`get-recolor` (recolor
a picture to grayscale, black-and-white, watermark, or automatic/no recolor via
`msoPictureAutomatic`, `msoPictureGrayscale`, `msoPictureBlackAndWhite`, `msoPictureWatermark`).
