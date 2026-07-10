---
"powerpointmcp": minor
---

Add gradient background support to the `slide` and `master` domains:
`set-gradient-background` / `get-gradient-background` set or read a two-color gradient fill
(style is an `MsoGradientStyle` name — `msoGradientHorizontal`, `msoGradientVertical`,
`msoGradientDiagonalUp`, `msoGradientDiagonalDown`, `msoGradientFromCorner`,
`msoGradientFromTitle`, `msoGradientFromCenter` — plus a `1`-`4` variant). `slide`'s version
overrides a single slide's background (same override semantics as the existing
`set-background-color`); `master`'s version sets the gradient for every slide that inherits from
the master.
