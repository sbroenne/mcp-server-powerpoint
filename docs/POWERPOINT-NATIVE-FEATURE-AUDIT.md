# PowerPoint-Native Feature Audit

## Scope

This audit compares the current 15 tools and 160 operations with the restored
`Microsoft.Office.Interop.PowerPoint` 15.0.4420.1018 assembly. It looks for useful PowerPoint
features that fit the existing live-session model. It does not copy Excel-only worksheet, Power
Query, Data Model, PivotTable, or calculation APIs.

Every accepted feature must use typed PowerPoint PIA members where available, release acquired COM
objects in `finally` blocks, keep 1-based indexes, propagate unexpected exceptions, and have a
real-PowerPoint integration test before implementation.

## Current coverage

The existing surface already covers the main editing workflow:

- presentation sessions, templates, themes, and document properties
- slides, sections, legacy comments, and slide import
- shapes, placeholders, hyperlinks, text, tables, images, charts, and SmartArt
- speaker notes, layouts, masters, page setup, animations, and transitions
- accessibility checks, reading order, PDF export, and rendered slide images

Modern Microsoft 365 threaded comments are not available through a reliable PowerPoint COM API.
The existing legacy comment operations are therefore correctly scoped.

## Recommended additions

| Priority | Capability | Proposed operations | Why it fits | Main risk |
| --- | --- | --- | --- | --- |
| 1 | Save As and Save Copy As | `presentation: save-as`, `save-copy-as` | High-value delivery workflow with direct Excel precedent | `save-as` must update the session path atomically |
| 2 | Mark as Final | `presentation: get-final`, `set-final` | Typed, reversible presentation state | It is advisory, not security |
| 3 | String tags | `set-tag`, `get-tag`, `list-tags`, `delete-tag` on presentation, slide, and shape | Safe PowerPoint-native metadata without XML complexity | Tag names are case-insensitive in PowerPoint and need normalization tests |
| 4 | Chart quick formatting | `get-style`, `set-style`, `get-color-style`, `set-color-style`, `get-data-table`, `set-data-table` | Useful visual control on the already acquired chart object | Style/color values are COM variants and need range tests against real PowerPoint |
| 5 | Linked pictures | optional `link_to_file` on `image: add-picture`; `shape: get-link-info`, `update-link`, `break-link`, `set-link-auto-update` | Enables linked-asset workflows and repair | `LinkFormat` is valid only for linked shapes |
| 6 | Audio and video | new `media` domain with `add-media`, `get-media-info` | Large PowerPoint-specific capability gap | Needs small licensed test fixtures and narrowly scoped Office enum handling |

### 1. Save As and Save Copy As — implemented

The restored PIA exposes:

- `_Presentation.SaveAs(string, PpSaveAsFileType, MsoTriState)`
- `_Presentation.SaveCopyAs(string, PpSaveAsFileType, MsoTriState)`

The restored signatures include an optional `Office.MsoTriState` parameter, but this project does
not reference `office.dll`, so the compiler cannot bind either method. The implementation keeps
late binding limited to these two invocations, passes typed `PpSaveAsFileType` values, and leaves
font embedding at PowerPoint's default.

`save-as` updates the batch and session registry path only after COM succeeds. `save-copy-as`
preserves the active session and original path. Both operations validate supported extensions,
format matches, destination directories, and overwrite intent before COM.

Real-PowerPoint tests cover:

- save as to `.pptx`, reopen the new path, and verify content
- save copy, verify both files exist, and verify the session still points to the original
- reject format/extension mismatches and existing destinations when overwrite is false
- verify a failed save does not change the tracked session path

References: [Presentation.SaveAs](https://learn.microsoft.com/office/vba/api/powerpoint.presentation.saveas),
[Presentation.SaveCopyAs](https://learn.microsoft.com/office/vba/api/powerpoint.presentation.savecopyas).

### 2. Mark as Final

`_Presentation.Final` is a typed `bool` property. Expose it as `get-final` and `set-final`, with
tool text stating clearly that it is an advisory editing flag, not access control.

Tests should set, save, reopen, read, clear, save, and reopen again.

Implemented with typed PIA access. PowerPoint persists the flag when it is set to `true` and then
makes the presentation read-only, so save-on-close remains accepted as a successful no-op while
the flag is set. Clearing the flag makes the presentation editable again and is persisted by the
normal save-on-close workflow. The flag is advisory only; it is not authentication, encryption,
or access control.

Reference: [Presentation.Final](https://learn.microsoft.com/office/vba/api/powerpoint.presentation.final).

### 3. Presentation, slide, and shape tags

The restored PIA exposes typed `Tags` collections on `_Presentation`, `_Slide`, and `Shape`.
`Tags.Add(name, value)`, `Tags.Item[name]`, `Tags.Name(index)`, `Tags.Value(index)`,
`Tags.Delete(name)`, and `Tags.Count` are sufficient for string metadata.

Use one shared Core helper for collection traversal and COM release, while keeping operations on
the three existing domains. Do not expose `AddBinary` or `BinaryValue`; arbitrary binary metadata
adds file-size and safety concerns without a clear automation need.

Tests should cover add/update, name-based lookup, 1-based enumeration, delete, missing tags, and
persistence after save/reopen for each owner type.

Reference: [Tags object](https://learn.microsoft.com/office/vba/api/powerpoint.tags).

### 4. Chart quick formatting

The restored `PowerPoint.Chart` type exposes `ChartStyle` and `ChartColor` as COM variant
properties and `HasDataTable` as `bool`. These properties operate on the chart object already
acquired by the current chart commands.

Add paired getters and setters so tests and callers can verify every write. Validate style and
color inputs against values accepted by the installed PowerPoint version rather than copying an
unverified numeric range from documentation.

Tests should set each property, read it back, save/reopen, and confirm invalid values fail without
altering the chart.

References: [Chart.ChartStyle](https://learn.microsoft.com/office/vba/api/powerpoint.chart.chartstyle),
[Chart.ChartColor](https://learn.microsoft.com/office/vba/api/powerpoint.chart.chartcolor),
[Chart.HasDataTable](https://learn.microsoft.com/office/vba/api/powerpoint.chart.hasdatatable).

### 5. Linked pictures and link management

`Shapes.AddPicture` accepts `MsoTriState` values for linking and embedding. The restored
`Shape.LinkFormat` property is typed and exposes `SourceFullName`, `AutoUpdate`,
`Update()`, and `BreakLink()`.

Keep embedding as the default. A linked picture must either be saved with the document or remain
resolvable by path; reject contradictory combinations before COM. Accessing `LinkFormat` on an
ordinary shape is an expected validation failure, not an unexpected exception.

Tests should create a linked image, read its source, change automatic update, update it, break the
link, and verify the image remains after the source file is removed.

References: [Shapes.AddPicture](https://learn.microsoft.com/office/vba/api/powerpoint.shapes.addpicture),
[LinkFormat](https://learn.microsoft.com/office/vba/api/powerpoint.linkformat).

### 6. Audio and video

`Shapes.AddMediaObject2` is the PowerPoint-native insertion method. It has the same Office enum
boundary as `AddPicture`. `Shape.MediaType` is a typed `PpMediaType` read-back property.

A dedicated `media` domain is clearer than treating playable content as an image. Start with
insertion and read-back only. Playback control and trimming require a separate audit because they
touch timing, codecs, and presentation-mode behavior.

Tests need small repository-owned audio and video fixtures, embedded and linked cases, media type
read-back, shape count, save/reopen, and cleanup.

References: [Shapes.AddMediaObject2](https://learn.microsoft.com/office/vba/api/powerpoint.shapes.addmediaobject2),
[Shape.MediaType](https://learn.microsoft.com/office/vba/api/powerpoint.shape.mediatype).

## Later, separately scoped work

- **Run-level text formatting and hyperlinks:** use `TextRange.Characters(start, length)` so one
  word or run can be formatted or linked without changing the whole shape. This changes many text
  operations and needs a dedicated contract design.
- **Advanced chart formatting:** data labels, axis scale, gridlines, series colors, and markers are
  useful, but several members cross into Excel chart types or COM variants. Implement only after
  the typed quick-formatting work establishes stable helpers and tests.
- **Persisted print options:** handout layout and color settings may be useful to a human who prints
  later. Add only after demonstrated demand and never call `PrintOut`.
- **Comment display preference:** `_Presentation.DisplayComments` is available but low value. It
  can accompany a future comments change rather than creating a separate feature.

## Explicitly rejected

| Capability | Decision | Reason |
| --- | --- | --- |
| Open-file passwords | Do not expose | `Presentations.Open` has no password parameter. Setting `_Presentation.Password` can create a file this server cannot reopen and may cause a blocking PowerPoint dialog. |
| Physical printing and print preview | Do not expose | Default-printer output and modal preview are unsafe for unattended automation. PDF export already provides deterministic output. |
| `CustomXMLParts` | Do not expose | It requires `Microsoft.Office.Core.CustomXMLParts`, adds complex XML ownership, and duplicates the useful metadata case covered by tags. |
| Window zoom, panes, and view chrome | Do not expose | The automation window is intentionally minimized and these controls do not improve the saved slide content. |
| Modern threaded comments | Unsupported upstream | PowerPoint COM does not provide a reliable modern threaded-comment surface; retain the documented legacy comment operations. |
| Bare `Presentation.UpdateLinks()` | Defer | It mainly updates linked OLE objects, which the current server cannot create. Shape-level link operations are testable and safer. |

The restored PIA confirms that `Presentations.Open` has only `FileName`, `ReadOnly`, `Untitled`,
and `WithWindow` parameters. It also confirms that `CustomXMLParts` and several print/media
parameters use Office types not currently referenced by the project.

References: [Presentations.Open](https://learn.microsoft.com/office/vba/api/powerpoint.presentations.open),
[Presentation.Password](https://learn.microsoft.com/office/vba/api/powerpoint.presentation.password),
[Presentation.PrintOut](https://learn.microsoft.com/office/vba/api/powerpoint.presentation.printout),
[DocumentWindow](https://learn.microsoft.com/office/vba/api/powerpoint.documentwindow).

## Delivery order

Each row below should be a separate pull request with generated CLI/MCP parity, skill updates, and
real-PowerPoint tests:

1. Save As and Save Copy As.
2. Mark as Final.
3. String tags.
4. Chart quick formatting.
5. Linked pictures and link management.
6. Audio and video.
7. Run-level text and advanced chart formatting only after separate contract reviews.

No accepted item requires a matching change in `mcp-server-excel`; Excel already has Save As,
copy, and workbook-link management, while tags and playable media are PowerPoint-specific. The
physical-printing rejection should remain consistent across all unattended Office and Windows
automation projects.
