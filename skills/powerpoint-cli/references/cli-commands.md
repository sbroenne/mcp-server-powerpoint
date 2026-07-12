# CLI Command Reference

> Auto-generated from `pptcli --help`. Use these exact parameter names.

## Global Commands

### session

Open, create, save, close, or list presentation sessions held by the daemon.

**Commands:** `open <FILE_PATH>`, `create <FILE_PATH>`, `close <SESSION_ID>`, `save <SESSION_ID>`, `list`

| Command | Description |
|---------|-------------|
| `session open <FILE_PATH>` | Open an existing presentation and return a session id |
| `session create <FILE_PATH>` | Create a new presentation and return a session id |
| `session close <SESSION_ID>` | Close a session, optionally saving first |
| `session save <SESSION_ID>` | Save the presentation open in a session |
| `session list` | List every session currently open in the daemon |

### service

Start, stop, or check the status of the `pptcli` background daemon.

**Commands:** `start`, `stop`, `status`

| Command | Description |
|---------|-------------|
| `service start` | Start the daemon if it isn't already running |
| `service stop` | Stop the running daemon |
| `service status` | Report whether the daemon is running |

## Domain Commands

Every domain command below follows the shape `pptcli <domain> <ACTION> [OPTIONS]`, targeting an
already-open session via `-s, --session <SESSION>` (obtained from `session open`/`session
create`). All slide/shape/row/column indices are 1-based, matching PowerPoint's own COM object
model.

### chart

Chart lifecycle and data operations.

**Actions:** `add-chart`, `get-chart-data`, `add-series`, `set-chart-title`, `get-chart-title`, `set-axis-title`, `get-axis-title`, `set-legend-visibility`, `get-legend-visibility`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index. (required) |
| `--chart-type` | Chart type: "bar", "line", or "pie". (required for: add-chart) |
| `--left` | Left position in points. (required for: add-chart) |
| `--top` | Top position in points. (required for: add-chart) |
| `--width` | Width in points. (required for: add-chart) |
| `--height` | Height in points. (required for: add-chart) |
| `--categories` | Category labels (x-axis / pie slice labels). (required for: add-chart) (JSON format) |
| `--series-name` | Name of the data series. (required for: add-chart, add-series) |
| `--values` | Data values, one per category. (required for: add-chart, add-series) (JSON format) |
| `--shape-index` | (required for: get-chart-data, add-series, set-chart-title, get-chart-title, set-axis-title, get-axis-title, set-legend-visibility, get-legend-visibility) |
| `--title` | Chart or axis title text. (required for: set-chart-title, set-axis-title) |
| `--axis-type` | Axis to target: "category" or "value". (required for: set-axis-title, get-axis-title) |
| `--visible` | Whether the legend is shown. (required for: set-legend-visibility) |

### export

Export commands: render presentation slides to raster image files. Operates within an
already-open session.

**Actions:** `export-slide-to-image`, `export-all-slides-to-images`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based index of the slide to export. (required for: export-slide-to-image) |
| `--output-path` | Full path for the output image file (e.g. `C:\output\slide1.png`). (required for: export-slide-to-image) |
| `--format` | PowerPoint filter name for the image format (e.g. "PNG", "JPG", "GIF"). Defaults to "PNG" |
| `--width` | Optional output width in pixels; 0 or null uses PowerPoint's default |
| `--height` | Optional output height in pixels; 0 or null uses PowerPoint's default |
| `--output-directory` | Directory where slide images will be written. Created if it does not exist. PowerPoint names the output files `Slide1.{ext}`, `Slide2.{ext}`, etc. (required for: export-all-slides-to-images) |

### image

Image commands: insert pictures and adjust picture formatting. Operates within an already-open
session, targeting a specific slide and picture shape by their 1-based indices.

**Actions:** `add-picture`, `set-brightness-contrast`, `get-brightness-contrast`, `set-recolor`, `get-recolor`, `set-crop`, `get-crop`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required for: set-brightness-contrast, get-brightness-contrast, set-recolor, get-recolor, set-crop, get-crop) |
| `--image-path` | (required for: add-picture) |
| `--left` | (required for: add-picture) |
| `--top` | (required for: add-picture) |
| `--width` | (required for: add-picture) |
| `--height` | (required for: add-picture) |
| `--brightness` | (required for: set-brightness-contrast) |
| `--contrast` | (required for: set-brightness-contrast) |
| `--color-type` | (required for: set-recolor) |
| `--crop-left` | (required for: set-crop) |
| `--crop-top` | (required for: set-crop) |
| `--crop-right` | (required for: set-crop) |
| `--crop-bottom` | (required for: set-crop) |

### layout

Slide layout commands: apply/read a slide's built-in layout. Operates within an already-open
session, targeting a specific slide by its 1-based index.

**Actions:** `set-layout`, `get-layout`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--layout-name` | (required for: set-layout) — a `PpSlideLayout` enum member name, e.g. `ppLayoutBlank`, `ppLayoutTitle` |

### master

Slide master commands: read/edit the title and body placeholder fonts on the presentation's
slide master, and read/edit the slide master's background fill color. Changes apply to every
slide inheriting from the master (i.e. any slide that doesn't itself override the property).

**Actions:** `get-title-font`, `set-title-font`, `get-body-font`, `set-body-font`, `get-background-color`, `set-background-color`

| Parameter | Description |
|-----------|-------------|
| `--font-name` | Font name, e.g. "Arial". (optional for: set-title-font, set-body-font) |
| `--font-size` | Font size in points. (optional for: set-title-font, set-body-font) |
| `--bold` | Whether the font is bold. (optional for: set-title-font, set-body-font) |
| `--red` | Red channel (0-255). (required for: set-background-color; pass together with green/blue for set-title-font/set-body-font color changes) |
| `--green` | Green channel (0-255). |
| `--blue` | Blue channel (0-255). |

### animation

Animation commands: add/delete shape entrance, emphasis, and exit effects on a slide's
timeline (`Slide.TimeLine.MainSequence`), and read/set a slide's transition
(`Slide.SlideShowTransition`).

**Actions:** `add-effect`, `get-effect-count`, `delete-effect`, `get-transition`, `set-transition`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required for: add-effect) |
| `--effect-name` | (required for: add-effect) — an `MsoAnimEffect` enum member name, e.g. `msoAnimEffectFade`, `msoAnimEffectFly` |
| `--is-exit` | Whether the effect plays as the shape leaving the slide (default false = entrance/emphasis). (optional for: add-effect) |
| `--trigger` | `"on-click"` (default), `"with-previous"`, or `"after-previous"`. (optional for: add-effect) |
| `--effect-index` | (required for: delete-effect) |
| `--transition-name` | (required for: set-transition) — a `PpEntryEffect` enum member name, e.g. `ppEffectFade`, `ppEffectCut` |
| `--duration-seconds` | (optional for: set-transition) |
| `--advance-on-click` | (optional for: set-transition) |
| `--advance-on-time` | (optional for: set-transition) |
| `--advance-time-seconds` | (optional for: set-transition) |

### notes

Speaker notes commands: set/get the notes text for a slide. Operates within an already-open
session, targeting a specific slide by its 1-based index.

**Actions:** `set-notes-text`, `get-notes-text`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--text` | (required for: set-notes-text) |

### presentation

Presentation lifecycle commands: create, open, save, apply-template, get-theme-name.

**Actions:** `create`, `open`, `save`, `apply-template`, `get-theme-name`

| Parameter | Description |
|-----------|-------------|
| `--file-path` | (required for: create, open) |
| `--is-macro-enabled` | IsMacroEnabled |
| `--template-path` | Full path to a `.potx`/`.potm`/`.pot` template file (a `.pptx`/`.pptm` presentation may also be used as a template source, matching PowerPoint's own behavior). (required for: apply-template) |

> Prefer the top-level `session open`/`session create`/`session save` commands for everyday
> session lifecycle management — `presentation <ACTION>` is the lower-level domain dispatch
> underneath them, exposed mainly for `apply-template` and `get-theme-name`, which have no
> `session`-level equivalent.

### shape

Shape commands: add rectangles/text boxes/auto-shapes/lines/connectors, count, delete,
reposition/resize, and format (fill/line/rotation/flip/z-order/shadow/group/name/alt-text).
Operates within an already-open session, targeting a specific slide by its 1-based index.

**Actions:** `add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`, `add-connector`, `get-count`, `delete`, `set-position`, `set-size`, `set-fill`, `get-fill`, `set-line`, `get-line`, `set-rotation`, `get-rotation`, `flip`, `set-z-order`, `set-shadow`, `get-shadow`, `group`, `ungroup`, `set-name`, `get-name`, `set-alt-text`, `get-alt-text`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--left` | (required for: add-rectangle, add-text-box, add-auto-shape, set-position) |
| `--top` | (required for: add-rectangle, add-text-box, add-auto-shape, set-position) |
| `--width` | (required for: add-rectangle, add-text-box, add-auto-shape, set-size) |
| `--height` | (required for: add-rectangle, add-text-box, add-auto-shape, set-size) |
| `--text` | (required for: add-text-box) |
| `--shape-type` | `MsoAutoShapeType` name, e.g. `msoShapeOval`, `msoShapeRightArrow`. (required for: add-auto-shape) |
| `--begin-x` | (required for: add-line, add-connector) |
| `--begin-y` | (required for: add-line, add-connector) |
| `--end-x` | (required for: add-line, add-connector) |
| `--end-y` | (required for: add-line, add-connector) |
| `--connector-type` | `msoConnectorStraight`, `msoConnectorElbow`, or `msoConnectorCurve`. (required for: add-connector) |
| `--shape-index` | (required for: delete, set-position, set-size, set-fill, get-fill, set-line, get-line, set-rotation, get-rotation, flip, set-z-order, set-shadow, get-shadow, ungroup, set-name, get-name, set-alt-text, get-alt-text) |
| `--red` | 0-255. (required for: set-fill; optional for: set-line) |
| `--green` | 0-255. (required for: set-fill; optional for: set-line) |
| `--blue` | 0-255. (required for: set-fill; optional for: set-line) |
| `--weight` | Line weight in points. (optional for: set-line) |
| `--dash-style` | `MsoLineDashStyle` name: `msoLineSolid`, `msoLineSquareDot`, `msoLineRoundDot`, `msoLineDash`, `msoLineDashDot`, `msoLineDashDotDot`, `msoLineLongDash`, `msoLineLongDashDot`. (optional for: set-line) |
| `--visible` | (required for: set-shadow; optional for: set-line) |
| `--degrees` | Rotation in degrees clockwise from upright. (required for: set-rotation) |
| `--direction` | `horizontal` or `vertical`. (required for: flip) |
| `--z-order-command` | `bring-to-front`, `send-to-back`, `bring-forward`, or `send-backward`. (required for: set-z-order) |
| `--shape-indexes` | JSON array of 1-based shape indices, at least 2. (required for: group) |
| `--name` | (required for: set-name) |
| `--alt-text` | (required for: set-alt-text) |

### slide

Slide lifecycle commands: add, delete, count, duplicate, reorder, per-slide background color,
and section management. First domain built on top of the presentation lifecycle commands,
operating within an already-open session.

**Actions:** `add-blank`, `get-count`, `delete`, `duplicate`, `move-to`, `set-background-color`, `get-background-color`, `add-section`, `rename-section`, `delete-section`, `get-section-count`, `get-section-name`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required for: delete, duplicate, move-to, set-background-color, get-background-color) |
| `--to-position` | (required for: move-to) |
| `--red` / `--green` / `--blue` | (required for: set-background-color, 0-255 per channel) |
| `--section-index` | (required for: add-section, rename-section, delete-section, get-section-name) |
| `--section-name` | (required for: rename-section; optional for: add-section) |
| `--delete-slides` | Optional bool (delete-section); if true, deletes the section's slides too. PowerPoint disallows deleting section 1 without this. |

### table

Table commands: add a table shape, read/write cell text, insert/delete rows and columns, format
cell fill and borders, and merge cells. Operates within an already-open session, targeting a
specific slide and table shape by their 1-based indices.

**Actions:** `add-table`, `set-cell-text`, `get-cell-text`, `insert-row`, `delete-row`, `insert-column`, `delete-column`, `set-cell-fill`, `get-cell-fill`, `set-cell-border`, `get-cell-border`, `merge-cells`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--rows` | (required for: add-table) |
| `--columns` | (required for: add-table) |
| `--left` | (required for: add-table) |
| `--top` | (required for: add-table) |
| `--width` | (required for: add-table) |
| `--height` | (required for: add-table) |
| `--shape-index` | (required for: set-cell-text, get-cell-text, insert-row, delete-row, insert-column, delete-column, set-cell-fill, get-cell-fill, set-cell-border, get-cell-border, merge-cells) |
| `--row` | (required for: set-cell-text, get-cell-text, delete-row, set-cell-fill, get-cell-fill, set-cell-border, get-cell-border, merge-cells) |
| `--column` | (required for: set-cell-text, get-cell-text, delete-column, set-cell-fill, get-cell-fill, set-cell-border, get-cell-border, merge-cells) |
| `--text` | (required for: set-cell-text) |
| `--before-row` | Optional row to insert before (insert-row); omit to append. |
| `--before-column` | Optional column to insert before (insert-column); omit to append. |
| `--red` / `--green` / `--blue` | (required for: set-cell-fill); optional for set-cell-border (0-255 per channel). |
| `--border-type` | (required for: set-cell-border, get-cell-border) — `PpBorderType` name: `ppBorderTop`, `ppBorderLeft`, `ppBorderBottom`, `ppBorderRight`, `ppBorderDiagonalDown`, `ppBorderDiagonalUp`. |
| `--weight` | Optional border line weight (set-cell-border). |
| `--dash-style` | Optional `MsoLineDashStyle` name, e.g. `msoLineSolid`, `msoLineDash` (set-cell-border). |
| `--visible` | Optional border visibility (set-cell-border). |
| `--merge-to-row` / `--merge-to-column` | (required for: merge-cells) — the adjacent cell to merge into. |

### textframe

Text frame commands: set/get text and font/paragraph formatting (size, bold, italic, underline,
font name, color, alignment, bullets) for a shape's text range. Operates within an already-open
session, targeting a specific shape by its 1-based slide and shape index.

**Actions:** `set-text`, `get-text`, `set-font-size`, `set-bold`, `set-font-color`, `set-italic`, `get-italic`, `set-underline`, `get-underline`, `set-font-name`, `get-font-name`, `set-alignment`, `get-alignment`, `set-bullet`, `get-bullet`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required) |
| `--text` | (required for: set-text) |
| `--font-size` | (required for: set-font-size) |
| `--bold` | (required for: set-bold) |
| `--red` | (required for: set-font-color) — 0-255 |
| `--green` | (required for: set-font-color) — 0-255 |
| `--blue` | (required for: set-font-color) — 0-255 |
| `--italic` | (required for: set-italic) |
| `--underline` | (required for: set-underline) |
| `--font-name` | (required for: set-font-name) |
| `--alignment` | `PpParagraphAlignment` name: `ppAlignLeft`, `ppAlignCenter`, `ppAlignRight`, `ppAlignJustify`, `ppAlignDistribute`, `ppAlignThaiDistribute`, `ppAlignJustifyLow`. (required for: set-alignment) |
| `--enabled` | (required for: set-bullet) |
| `--character` | Single-character bullet glyph. (optional for: set-bullet) |

## Common Options (All Domain Commands)

| Option | Description |
|--------|-------------|
| `-h, --help` | Prints help information |
| `-s, --session <SESSION>` | Session ID from `session open`/`session create` |
| `-o, --output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as a binary file |
