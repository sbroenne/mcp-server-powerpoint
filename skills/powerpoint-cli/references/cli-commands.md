# CLI Command Reference

> Generated from `skills/powerpoint-cli/SKILL.md` by `scripts/Build-CliCommandReference.ps1`.
> Do not edit by hand — run the script instead. Use these exact parameter names.

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

### animation

Animation commands: add/delete entrance, emphasis, and exit effects on a shape's slide timeline (Slide.TimeLine.MainSequence), and read/set a slide's transition (Slide.SlideShowTransition). Operates within an already-open , targeting a specific slide (and, for shape effects, a specific shape) by 1-based index.

**Actions:** `add-effect`, `get-effect-count`, `delete-effect`, `get-transition`, `set-transition`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required for: add-effect) |
| `--effect-name` | (required for: add-effect) |
| `--is-exit` | When true, the effect is applied as the shape leaving the slide (exit) rather than the     default entrance/emphasis behavior. |
| `--trigger` | When the effect starts: "on-click" (default), "with-previous", or     "after-previous". |
| `--effect-index` | (required for: delete-effect) |
| `--transition-name` | (required for: set-transition) |
| `--duration-seconds` |  |
| `--advance-on-click` |  |
| `--advance-on-time` |  |
| `--advance-time-seconds` |  |

### chart

Chart lifecycle and data operations.

**Actions:** `add-chart`, `get-chart-data`, `add-series`, `set-chart-title`, `get-chart-title`, `set-axis-title`, `get-axis-title`, `set-legend-visibility`, `get-legend-visibility`, `replace-chart-data`

| Parameter | Description |
|------|-------------|
| `--slide-index` | 1-based slide index. (required) |
| `--chart-type` | Chart type: "bar", "line", or "pie". (required for: add-chart) |
| `--left` | Left position in points. (required for: add-chart) |
| `--top` | Top position in points. (required for: add-chart) |
| `--width` | Width in points. (required for: add-chart) |
| `--height` | Height in points. (required for: add-chart) |
| `--categories` | Category labels (x-axis / pie slice labels). (required for: add-chart, replace-chart-data) |
| `--series-name` | Name of the single data series. (required for: add-chart, add-series) |
| `--values` | Data values, one per category. (required for: add-chart, add-series) |
| `--shape-index` | (required for: get-chart-data, add-series, set-chart-title, get-chart-title, set-axis-title, get-axis-title, set-legend-visibility, get-legend-visibility, replace-chart-data) |
| `--title` | (required for: set-chart-title, set-axis-title) |
| `--axis-type` | (required for: set-axis-title, get-axis-title) |
| `--visible` | (required for: set-legend-visibility) |
| `--series-names` | (required for: replace-chart-data) |
| `--series-values` | (required for: replace-chart-data) |

### export

Export commands: render presentation slides to raster image files. Operates within an already-open .

**Actions:** `export-slide-to-image`, `export-all-slides-to-images`

| Parameter | Description |
|------|-------------|
| `--slide-index` | 1-based index of the slide to export. (required for: export-slide-to-image) |
| `--output-path` | Full path for the output image file (e.g. C:\output\slide1.png). (required for: export-slide-to-image) |
| `--format` | PowerPoint filter name for the image format (e.g. "PNG", "JPG", "GIF").     Defaults to "PNG". |
| `--width` | Optional output width in pixels; 0 or null uses PowerPoint's default. |
| `--height` | Optional output height in pixels; 0 or null uses PowerPoint's default. |
| `--output-directory` | Directory where slide images will be written. Created if it does not exist.     PowerPoint names the output files Slide1.{ext}, Slide2.{ext}, etc. (required for: export-all-slides-to-images) |

### image

Image commands: embed a picture file into a slide. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

**Actions:** `add-picture`, `set-brightness-contrast`, `get-brightness-contrast`, `set-recolor`, `get-recolor`, `set-crop`, `get-crop`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--image-path` | (required for: add-picture) |
| `--left` | (required for: add-picture) |
| `--top` | (required for: add-picture) |
| `--width` | (required for: add-picture) |
| `--height` | (required for: add-picture) |
| `--shape-index` | (required for: set-brightness-contrast, get-brightness-contrast, set-recolor, get-recolor, set-crop, get-crop) |
| `--brightness` | (required for: set-brightness-contrast) |
| `--contrast` | (required for: set-brightness-contrast) |
| `--color-type` | (required for: set-recolor) |
| `--crop-left` | (required for: set-crop) |
| `--crop-top` | (required for: set-crop) |
| `--crop-right` | (required for: set-crop) |
| `--crop-bottom` | (required for: set-crop) |

### layout

Slide layout commands: apply/read a slide's built-in layout. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

**Actions:** `set-layout`, `get-layout`, `list-layouts`, `delete-layout`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required for: set-layout, get-layout) |
| `--layout-name` | (required for: set-layout) |
| `--master-index` | (required for: list-layouts, delete-layout) |
| `--layout-index` | (required for: delete-layout) |

### master

Slide master commands: read/edit the title and body placeholder fonts on the presentation's slide master, and read/edit the slide master's background fill color. Operates within an already-open . Changes here apply to every slide that inherits from the master (i.e. any slide that does not itself override the property), which is the practical "edit the master, not each slide" workflow PowerPoint's COM object model supports safely.

**Actions:** `get-title-font`, `set-title-font`, `get-body-font`, `set-body-font`, `get-background-color`, `set-background-color`, `set-gradient-background`, `get-gradient-background`, `list-masters`, `delete-master`

| Parameter | Description |
|------|-------------|
| `--font-name` |  |
| `--font-size` |  |
| `--bold` |  |
| `--red` | (required for: set-background-color) |
| `--green` | (required for: set-background-color) |
| `--blue` | (required for: set-background-color) |
| `--red1` | (required for: set-gradient-background) |
| `--green1` | (required for: set-gradient-background) |
| `--blue1` | (required for: set-gradient-background) |
| `--red2` | (required for: set-gradient-background) |
| `--green2` | (required for: set-gradient-background) |
| `--blue2` | (required for: set-gradient-background) |
| `--gradient-style` |  |
| `--gradient-variant` |  |
| `--master-index` | (required for: delete-master) |

### notes

Speaker notes commands: set/get the notes text for a slide. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

**Actions:** `set-notes-text`, `get-notes-text`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--text` | (required for: set-notes-text) |

### shape

Shape commands: add rectangles/text boxes, count, delete, reposition/resize. Operates within an already-open IPresentationBatch, targeting a specific slide by its 1-based index.

**Actions:** `add-rectangle`, `add-text-box`, `add-auto-shape`, `add-line`, `add-connector`, `get-count`, `delete`, `set-position`, `set-size`, `set-fill`, `get-fill`, `set-line`, `get-line`, `set-rotation`, `get-rotation`, `flip`, `set-z-order`, `set-shadow`, `get-shadow`, `set-glow`, `get-glow`, `set-reflection`, `get-reflection`, `set-soft-edge`, `get-soft-edge`, `set-bevel`, `get-bevel`, `group`, `ungroup`, `set-name`, `get-name`, `set-alt-text`, `get-alt-text`, `set-hyperlink`, `get-hyperlink`, `remove-hyperlink`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--left` | (required for: add-rectangle, add-text-box, add-auto-shape, set-position) |
| `--top` | (required for: add-rectangle, add-text-box, add-auto-shape, set-position) |
| `--width` | (required for: add-rectangle, add-text-box, add-auto-shape, set-size) |
| `--height` | (required for: add-rectangle, add-text-box, add-auto-shape, set-size) |
| `--text` | (required for: add-text-box) |
| `--shape-type` | (required for: add-auto-shape) |
| `--begin-x` | (required for: add-line, add-connector) |
| `--begin-y` | (required for: add-line, add-connector) |
| `--end-x` | (required for: add-line, add-connector) |
| `--end-y` | (required for: add-line, add-connector) |
| `--connector-type` | (required for: add-connector) |
| `--shape-index` | (required for: delete, set-position, set-size, set-fill, get-fill, set-line, get-line, set-rotation, get-rotation, flip, set-z-order, set-shadow, get-shadow, set-glow, get-glow, set-reflection, get-reflection, set-soft-edge, get-soft-edge, set-bevel, get-bevel, ungroup, set-name, get-name, set-alt-text, get-alt-text, set-hyperlink, get-hyperlink, remove-hyperlink) |
| `--red` | (required for: set-fill, set-glow) |
| `--green` | (required for: set-fill, set-glow) |
| `--blue` | (required for: set-fill, set-glow) |
| `--weight` |  |
| `--dash-style` |  |
| `--visible` | (required for: set-shadow, set-reflection) |
| `--degrees` | (required for: set-rotation) |
| `--direction` | (required for: flip) |
| `--z-order-command` | (required for: set-z-order) |
| `--transparency` |  |
| `--blur` |  |
| `--offset-x` |  |
| `--offset-y` |  |
| `--radius` | (required for: set-glow, set-soft-edge) |
| `--size` |  |
| `--bevel-type` | (required for: set-bevel) |
| `--depth` |  |
| `--inset` |  |
| `--shape-indexes` | (required for: group) |
| `--name` | (required for: set-name) |
| `--alt-text` | (required for: set-alt-text) |
| `--address` | (required for: set-hyperlink) |
| `--screen-tip` |  |

### slide

Slide lifecycle commands: add, delete, count, duplicate, reorder, per-slide background color, and section management. First domain built on top of the presentation lifecycle commands, operating within an already-open IPresentationBatch.

**Actions:** `add-blank`, `get-count`, `delete`, `duplicate`, `move-to`, `set-background-color`, `get-background-color`, `set-gradient-background`, `get-gradient-background`, `add-section`, `rename-section`, `delete-section`, `get-section-count`, `get-section-name`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required for: delete, duplicate, move-to, set-background-color, get-background-color, set-gradient-background, get-gradient-background) |
| `--to-position` | (required for: move-to) |
| `--red` | (required for: set-background-color) |
| `--green` | (required for: set-background-color) |
| `--blue` | (required for: set-background-color) |
| `--red1` | (required for: set-gradient-background) |
| `--green1` | (required for: set-gradient-background) |
| `--blue1` | (required for: set-gradient-background) |
| `--red2` | (required for: set-gradient-background) |
| `--green2` | (required for: set-gradient-background) |
| `--blue2` | (required for: set-gradient-background) |
| `--gradient-style` |  |
| `--gradient-variant` |  |
| `--section-index` | (required for: add-section, rename-section, delete-section, get-section-name) |
| `--section-name` | (required for: rename-section) |
| `--delete-slides` |  |

### smartart

SmartArt commands: add a SmartArt diagram to a slide from PowerPoint's built-in layout gallery, and add/read/update/delete/count the diagram's nodes. Operates within an already-open , targeting a specific slide and shape by 1-based index.

**Actions:** `add-smart-art`, `add-node`, `add-child-node`, `set-node-text`, `get-node-text`, `delete-node`, `get-node-count`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--layout-name` | (required for: add-smart-art) |
| `--left` | (required for: add-smart-art) |
| `--top` | (required for: add-smart-art) |
| `--width` | (required for: add-smart-art) |
| `--height` | (required for: add-smart-art) |
| `--shape-index` | (required for: add-node, add-child-node, set-node-text, get-node-text, delete-node, get-node-count) |
| `--text` | (required for: add-node, add-child-node, set-node-text) |
| `--parent-node-index` | (required for: add-child-node) |
| `--node-index` | (required for: set-node-text, get-node-text, delete-node) |

### table

Table commands: add a table shape, read/write cell text, insert/delete rows and columns, format cell fill and borders, and merge cells. Operates within an already-open IPresentationBatch, targeting a specific slide and table shape by their 1-based indices.

**Actions:** `add-table`, `set-cell-text`, `get-cell-text`, `insert-row`, `delete-row`, `insert-column`, `delete-column`, `set-cell-fill`, `get-cell-fill`, `set-cell-border`, `get-cell-border`, `merge-cells`

| Parameter | Description |
|------|-------------|
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
| `--before-row` |  |
| `--before-column` |  |
| `--red` | (required for: set-cell-fill) |
| `--green` | (required for: set-cell-fill) |
| `--blue` | (required for: set-cell-fill) |
| `--border-type` | (required for: set-cell-border, get-cell-border) |
| `--weight` |  |
| `--dash-style` |  |
| `--visible` |  |
| `--merge-to-row` | (required for: merge-cells) |
| `--merge-to-column` | (required for: merge-cells) |

### textframe

Text frame commands: set/get text and basic font formatting (size, bold, italic, underline, font name, color, alignment, bullets) for a shape's text range. Operates within an already-open IPresentationBatch, targeting a specific shape by its 1-based slide and shape index.

**Actions:** `set-text`, `get-text`, `set-font-size`, `get-font-size`, `set-bold`, `get-bold`, `set-font-color`, `get-font-color`, `set-italic`, `get-italic`, `set-underline`, `get-underline`, `set-font-name`, `get-font-name`, `set-alignment`, `get-alignment`, `set-bullet`, `get-bullet`, `set-auto-size`, `get-auto-size`

| Parameter | Description |
|------|-------------|
| `--slide-index` | (required) |
| `--shape-index` | (required) |
| `--text` | (required for: set-text) |
| `--font-size` | (required for: set-font-size) |
| `--bold` | (required for: set-bold) |
| `--red` | (required for: set-font-color) |
| `--green` | (required for: set-font-color) |
| `--blue` | (required for: set-font-color) |
| `--italic` | (required for: set-italic) |
| `--underline` | (required for: set-underline) |
| `--font-name` | (required for: set-font-name) |
| `--alignment` | (required for: set-alignment) |
| `--enabled` | (required for: set-bullet) |
| `--character` |  |
| `--auto-size` | (required for: set-auto-size) |
